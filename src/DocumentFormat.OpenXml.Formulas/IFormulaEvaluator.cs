// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;

using DocumentFormat.OpenXml.Features.FormulaEvaluation.DependencyGraph;
using DocumentFormat.OpenXml.Spreadsheet;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation;

/// <summary>
/// Feature interface for formula evaluation.
/// </summary>
public interface IFormulaEvaluator : IDisposable
{
    /// <summary>
    /// Evaluates a single cell formula.
    /// </summary>
    /// <param name="worksheet">The worksheet containing the cell.</param>
    /// <param name="cell">The cell containing the formula.</param>
    /// <returns>Result containing the evaluated cell value or an error.</returns>
    Result<CellValue> TryEvaluate(Worksheet worksheet, Cell cell);

    /// <summary>
    /// Recalculates all formulas in the worksheet in dependency order.
    /// </summary>
    void RecalculateSheet(Worksheet worksheet);

    /// <summary>
    /// Recalculates only the formulas that depend on the specified changed cells.
    /// This is significantly faster than RecalculateSheet for incremental updates.
    /// </summary>
    /// <param name="worksheet">The worksheet containing the cells.</param>
    /// <param name="changedCells">The cell references that were changed (e.g., "A1", "B5").</param>
    void RecalculateDependents(Worksheet worksheet, params string[] changedCells);

    /// <summary>
    /// Gets the dependency graph for the worksheet.
    /// </summary>
    IDependencyGraph GetDependencyGraph(Worksheet worksheet);

    /// <summary>
    /// Gets evaluation statistics.
    /// </summary>
    EvaluatorStatistics GetStatistics();

    /// <summary>
    /// Checks if a function is supported.
    /// </summary>
    bool IsFunctionSupported(string functionName);

    /// <summary>
    /// Gets the set of supported function names.
    /// </summary>
    HashSet<string> SupportedFunctions { get; }
}
