// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.DependencyGraph;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Parsing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation;

/// <summary>
/// Implementation of formula evaluator feature.
/// </summary>
public class FormulaEvaluator : IFormulaEvaluator
{
    private readonly SpreadsheetDocument _document;
    private readonly Dictionary<string, Func<CellContext, CellValue>> _compiledFormulas = new();
    private readonly object _lock = new();
    private readonly FormulaParser _parser = new();
    private readonly FormulaCompiler _compiler = new();
    private readonly EvaluatorStatistics _statistics = new();

    /// <summary>
    /// Initializes a new instance of the <see cref="FormulaEvaluator"/> class.
    /// </summary>
    /// <param name="document">The spreadsheet document.</param>
    public FormulaEvaluator(SpreadsheetDocument document)
    {
        _document = document ?? throw new ArgumentNullException(nameof(document));
    }

    /// <inheritdoc/>
    public Result<CellValue> TryEvaluate(Worksheet worksheet, Cell cell)
    {
        _statistics.TotalEvaluations++;

        try
        {
            if (worksheet == null)
            {
                throw new ArgumentNullException(nameof(worksheet));
            }

            if (cell == null)
            {
                throw new ArgumentNullException(nameof(cell));
            }

            // Get the formula
            var cellFormula = cell.CellFormula?.Text;
            if (string.IsNullOrEmpty(cellFormula))
            {
                return Result<CellValue>.Failure(new ParserException("Cell does not contain a formula"));
            }

            // Get or compile the formula
            var compiledFormula = GetOrCompileFormula(cellFormula);

            // Get the shared string table part
            var sharedStringTablePart = _document.WorkbookPart?.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

            // Execute the formula
            var context = new CellContext(worksheet, sharedStringTablePart);
            var result = compiledFormula(context);

            _statistics.SuccessfulEvaluations++;
            return Result<CellValue>.Success(result);
        }
        catch (ParserException ex)
        {
            _statistics.FailedEvaluations++;
            return Result<CellValue>.Failure(ex);
        }
        catch (CompilationException ex)
        {
            _statistics.FailedEvaluations++;
            return Result<CellValue>.Failure(ex);
        }
        catch (UnsupportedFunctionException ex)
        {
            _statistics.FailedEvaluations++;
            return Result<CellValue>.Failure(ex);
        }
        catch (Exception ex)
        {
            _statistics.FailedEvaluations++;
            return Result<CellValue>.Failure(new ParserException($"Evaluation failed: {ex.Message}"));
        }
    }

    /// <inheritdoc/>
    public void RecalculateSheet(Worksheet worksheet)
    {
        // 1. Build dependency graph
        var graph = BuildDependencyGraph(worksheet);

        // 2. Get evaluation order (topological sort)
        var evalOrder = graph.GetEvaluationOrder();

        // 3. Evaluate cells in order
        foreach (var cellRef in evalOrder)
        {
            var cell = FindCellByReference(worksheet, cellRef);
            if (cell?.CellFormula != null)
            {
                var result = TryEvaluate(worksheet, cell);
                if (result.IsSuccess)
                {
                    // Update cached value
                    UpdateCellValue(cell, result.Value);
                }
            }
        }
    }

    /// <inheritdoc/>
    public void RecalculateDependents(Worksheet worksheet, params string[] changedCells)
    {
        if (worksheet == null)
        {
            throw new ArgumentNullException(nameof(worksheet));
        }

        if (changedCells == null || changedCells.Length == 0)
        {
            return;
        }

        // 1. Build dependency graph
        var graph = BuildDependencyGraph(worksheet);

        // 2. Find all dependents using BFS
        var dirtyCells = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var queue = new Queue<string>(changedCells);
        var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        while (queue.Count > 0)
        {
            var current = queue.Dequeue();
            if (visited.Contains(current))
            {
                continue;
            }

            visited.Add(current);

            // Get all cells that depend on this one
            var dependents = graph.GetDependents(current);
            foreach (var dependent in dependents)
            {
                dirtyCells.Add(dependent);
                queue.Enqueue(dependent);
            }
        }

        // 3. If no dependents, nothing to do
        if (dirtyCells.Count == 0)
        {
            return;
        }

        // 4. Get evaluation order for dirty cells only (topological sort)
        var evalOrder = graph.GetEvaluationOrder(dirtyCells);

        // 5. Evaluate cells in dependency order
        foreach (var cellRef in evalOrder)
        {
            var cell = FindCellByReference(worksheet, cellRef);
            if (cell?.CellFormula != null)
            {
                var result = TryEvaluate(worksheet, cell);
                if (result.IsSuccess)
                {
                    // Update cached value
                    UpdateCellValue(cell, result.Value);
                }
            }
        }
    }

    /// <inheritdoc/>
    public IDependencyGraph GetDependencyGraph(Worksheet worksheet)
    {
        return BuildDependencyGraph(worksheet);
    }

    /// <inheritdoc/>
    public EvaluatorStatistics GetStatistics()
    {
        lock (_lock)
        {
            _statistics.CompiledFormulaCount = _compiledFormulas.Count;
            _statistics.SupportedFunctionCount = SupportedFunctions.Count;
            return _statistics;
        }
    }

    /// <inheritdoc/>
    public bool IsFunctionSupported(string functionName)
    {
        return FunctionRegistry.TryGetFunction(functionName, out _);
    }

    /// <inheritdoc/>
    public HashSet<string> SupportedFunctions => GetSupportedFunctionNames();

    /// <inheritdoc/>
    public void Dispose()
    {
        _compiledFormulas.Clear();
    }

    private Func<CellContext, CellValue> GetOrCompileFormula(string formula)
    {
        lock (_lock)
        {
            if (_compiledFormulas.TryGetValue(formula, out var cached))
            {
                return cached;
            }

            // Parse the formula
            var ast = _parser.Parse(formula);

            // Compile to expression tree
            var expression = _compiler.Compile(ast);

            // Compile to delegate
            var compiled = expression.Compile();
            _compiledFormulas[formula] = compiled;
            return compiled;
        }
    }

    private static HashSet<string> GetSupportedFunctionNames()
    {
        // Get all function names from registry
        var functions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            // Phase 0 (3)
            "SUM", "AVERAGE", "IF",

            // Math (35)
            "COUNT", "COUNTA", "COUNTBLANK", "COUNTIF", "COUNTIFS", "MAX", "MIN",
            "ROUND", "ROUNDUP", "ROUNDDOWN", "ABS", "PRODUCT", "POWER", "SUMIF", "SUMIFS",
            "SQRT", "MOD", "INT", "CEILING", "FLOOR", "TRUNC",
            "SIGN", "EXP", "LN", "LOG", "LOG10", "PI", "RADIANS", "DEGREES",
            "SIN", "COS", "TAN",
            "COMBIN", "PERMUT", "MROUND", "QUOTIENT",

            // Logical (8)
            "AND", "OR", "NOT", "CHOOSE", "IFS", "SWITCH", "XOR", "IFNA",

            // Text (31)
            "CONCATENATE", "CONCAT", "TEXTJOIN", "LEFT", "RIGHT", "MID", "LEN", "TRIM", "TRIMALL",
            "UPPER", "LOWER", "PROPER", "TEXT", "VALUE",
            "FIND", "SEARCH", "SUBSTITUTE", "REPLACE", "REPT",
            "EXACT", "CHAR", "CODE", "UNICHAR", "UNICODE", "CLEAN", "T", "REVERSE",
            "FIXED", "DOLLAR", "NUMBERVALUE", "PHONETIC",

            // Lookup (11)
            "VLOOKUP", "HLOOKUP", "INDEX", "MATCH", "COLUMN", "ROW", "COLUMNS", "ROWS", "ADDRESS",
            "OFFSET", "INDIRECT",

            // Date/Time (22)
            "TODAY", "NOW", "DATE", "YEAR", "MONTH", "DAY",
            "HOUR", "MINUTE", "SECOND", "WEEKDAY", "WEEKNUM",
            "DAYS", "TIME", "TIMEVALUE", "DATEVALUE", "DAYS360", "EOMONTH", "EDATE",
            "NETWORKDAYS", "WORKDAY", "YEARFRAC", "DATEDIF",

            // Statistical (12)
            "MEDIAN", "MODE", "STDEV", "VAR", "RANK", "AVERAGEIF", "AVERAGEIFS",
            "MAXIFS", "MINIFS", "SKEW", "KURT", "FREQUENCY",

            // Information (13)
            "ISNUMBER", "ISTEXT", "IFERROR", "ISERROR", "ISNA", "ISERR", "ISBLANK",
            "ISEVEN", "ISODD", "ISLOGICAL", "ISNONTEXT", "TYPE", "N",

            // Financial (13)
            "PMT", "FV", "PV", "NPER", "RATE", "NPV", "IRR", "IPMT", "PPMT",
            "SLN", "DB", "DDB", "SYD",

            // Engineering (7)
            "CONVERT", "HEX2DEC", "DEC2HEX", "BIN2DEC", "DEC2BIN", "OCT2DEC", "DEC2OCT",
        };

        return functions;
    }

    private IDependencyGraph BuildDependencyGraph(Worksheet worksheet)
    {
        var graph = new DependencyGraph.DependencyGraph();

        foreach (var cell in worksheet.Descendants<Cell>())
        {
            if (cell.CellFormula != null && !string.IsNullOrEmpty(cell.CellFormula.Text))
            {
                try
                {
                    var formula = cell.CellFormula.Text;
                    var ast = _parser.Parse(formula);
                    var deps = ExtractDependencies(ast);
                    graph.AddCell(cell.CellReference?.Value ?? string.Empty, deps);
                }
                catch
                {
                    // Skip cells with invalid formulas
                }
            }
        }

        return graph;
    }

    private HashSet<string> ExtractDependencies(FormulaNode ast)
    {
        var deps = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        ExtractDependenciesRecursive(ast, deps);
        return deps;
    }

    private void ExtractDependenciesRecursive(FormulaNode node, HashSet<string> deps)
    {
        switch (node)
        {
            case CellReferenceNode cellRef:
                deps.Add(cellRef.Reference);
                break;

            case RangeNode range:
                // Expand range A1:A10 → A1, A2, ..., A10
                foreach (var cellRef in ExpandRange(range.Start, range.End))
                {
                    deps.Add(cellRef);
                }

                break;

            case BinaryOpNode bin:
                ExtractDependenciesRecursive(bin.Left, deps);
                ExtractDependenciesRecursive(bin.Right, deps);
                break;

            case UnaryOpNode un:
                ExtractDependenciesRecursive(un.Operand, deps);
                break;

            case FunctionCallNode func:
                foreach (var arg in func.Arguments)
                {
                    ExtractDependenciesRecursive(arg, deps);
                }

                break;
        }
    }

    private static IEnumerable<string> ExpandRange(string start, string end)
    {
        // Parse start and end cell references (e.g., A1, B10)
        // Generate all cell references in the range
        int startCol, startRow, endCol, endRow;
        ParseCellReference(start, out startCol, out startRow);
        ParseCellReference(end, out endCol, out endRow);

        for (int row = startRow; row <= endRow; row++)
        {
            for (int col = startCol; col <= endCol; col++)
            {
                yield return GetColumnLetter(col) + row.ToString(CultureInfo.InvariantCulture);
            }
        }
    }

    private static void ParseCellReference(string reference, out int col, out int row)
    {
        // Remove $ signs
        reference = reference.Replace("$", string.Empty);

        int i = 0;
        while (i < reference.Length && char.IsLetter(reference[i]))
        {
            i++;
        }

        var colPart = reference.Substring(0, i);
        var rowPart = reference.Substring(i);

        col = ParseColumnLetter(colPart);
        row = int.Parse(rowPart, CultureInfo.InvariantCulture);
    }

    private static int ParseColumnLetter(string column)
    {
        int result = 0;
        foreach (char c in column)
        {
            result = (result * 26) + (char.ToUpperInvariant(c) - 'A' + 1);
        }

        return result;
    }

    private static string GetColumnLetter(int column)
    {
        string result = string.Empty;
        while (column > 0)
        {
            int remainder = (column - 1) % 26;
            result = (char)('A' + remainder) + result;
            column = (column - 1) / 26;
        }

        return result;
    }

    private static Cell? FindCellByReference(Worksheet worksheet, string cellReference)
    {
        return worksheet.Descendants<Cell>()
            .FirstOrDefault(c => string.Equals(c.CellReference?.Value, cellReference, StringComparison.OrdinalIgnoreCase));
    }

    private static void UpdateCellValue(Cell cell, CellValue value)
    {
        switch (value.Type)
        {
            case CellValueType.Number:
                cell.DataType = null; // Number is the default
                cell.CellValue = new Spreadsheet.CellValue(value.NumericValue.ToString(CultureInfo.InvariantCulture));
                break;
            case CellValueType.Text:
                cell.DataType = new EnumValue<CellValues>(CellValues.String);
                cell.CellValue = new Spreadsheet.CellValue(value.StringValue);
                break;
            case CellValueType.Boolean:
                cell.DataType = new EnumValue<CellValues>(CellValues.Boolean);
                cell.CellValue = new Spreadsheet.CellValue(value.BoolValue ? "1" : "0");
                break;
            case CellValueType.Error:
                cell.DataType = new EnumValue<CellValues>(CellValues.Error);
                cell.CellValue = new Spreadsheet.CellValue(value.ErrorValue);
                break;
        }
    }
}
