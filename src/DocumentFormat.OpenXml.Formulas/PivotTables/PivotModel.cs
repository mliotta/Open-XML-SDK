// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

using DocumentFormat.OpenXml.Features.FormulaEvaluation;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

namespace DocumentFormat.OpenXml.Features.PivotTables;

/// <summary>
/// The precomputed pivot grid for an arbitrary number of row fields, column fields, and value
/// fields (subtotals off). Aggregation reuses the built-in functions in <see cref="FunctionRegistry"/>.
/// </summary>
internal sealed class PivotModel
{
    private PivotModel(
        int[] rowFieldIndices,
        int[] columnFieldIndices,
        int[] valueFieldIndices,
        List<IList<string>> rowFieldMembers,
        List<IList<string>> columnFieldMembers,
        List<int[]> rowTuples,
        List<int[]> columnTuples,
        string[] valueDisplayNames,
        double?[][][] data,
        double?[][] rowTotals,
        double?[][] columnTotals,
        double?[] grandTotals)
    {
        RowFieldIndices = rowFieldIndices;
        ColumnFieldIndices = columnFieldIndices;
        ValueFieldIndices = valueFieldIndices;
        RowFieldMembers = rowFieldMembers;
        ColumnFieldMembers = columnFieldMembers;
        RowTuples = rowTuples;
        ColumnTuples = columnTuples;
        ValueDisplayNames = valueDisplayNames;
        Data = data;
        RowTotals = rowTotals;
        ColumnTotals = columnTotals;
        GrandTotals = grandTotals;
    }

    /// <summary>Gets the source column index of each row field.</summary>
    public int[] RowFieldIndices { get; }

    /// <summary>Gets the source column index of each column field.</summary>
    public int[] ColumnFieldIndices { get; }

    /// <summary>Gets the source column index of each value field.</summary>
    public int[] ValueFieldIndices { get; }

    /// <summary>Gets the ordered members of each row field.</summary>
    public List<IList<string>> RowFieldMembers { get; }

    /// <summary>Gets the ordered members of each column field.</summary>
    public List<IList<string>> ColumnFieldMembers { get; }

    /// <summary>Gets the ordered row tuples (each is a member index per row field).</summary>
    public List<int[]> RowTuples { get; }

    /// <summary>Gets the ordered column tuples (each is a member index per column field).</summary>
    public List<int[]> ColumnTuples { get; }

    /// <summary>Gets the value-field captions.</summary>
    public string[] ValueDisplayNames { get; }

    /// <summary>Gets the aggregated values indexed by <c>[rowTuple][columnTuple][value]</c>; null is blank.</summary>
    public double?[][][] Data { get; }

    /// <summary>Gets the per-row-tuple totals indexed by <c>[rowTuple][value]</c>.</summary>
    public double?[][] RowTotals { get; }

    /// <summary>Gets the per-column-tuple totals indexed by <c>[columnTuple][value]</c>.</summary>
    public double?[][] ColumnTotals { get; }

    /// <summary>Gets the overall grand totals indexed by <c>[value]</c>.</summary>
    public double?[] GrandTotals { get; }

    /// <summary>Gets the number of row fields.</summary>
    public int RowFieldCount => RowFieldIndices.Length;

    /// <summary>Gets the number of column fields.</summary>
    public int ColumnFieldCount => ColumnFieldIndices.Length;

    /// <summary>Gets the number of value fields.</summary>
    public int ValueCount => ValueFieldIndices.Length;

    /// <summary>Computes the pivot grid for the given plan and source data.</summary>
    /// <param name="source">The source data.</param>
    /// <param name="plan">The validated plan.</param>
    /// <param name="context">A cell context used to invoke the aggregation functions.</param>
    /// <returns>The computed model.</returns>
    public static PivotModel Compute(PivotSourceData source, PivotPlan plan, CellContext context)
    {
        var rowFieldIndices = plan.RowFields.Select(source.FieldIndex).ToArray();
        var columnFieldIndices = plan.ColumnFields.Select(source.FieldIndex).ToArray();
        var valueFieldIndices = plan.ValueFields.Select(v => source.FieldIndex(v.FieldName)).ToArray();
        var functions = plan.ValueFields.Select(v => ResolveFunction(v.Aggregate)).ToArray();
        var valueCount = valueFieldIndices.Length;

        var rowFieldMembers = BuildFieldMembers(source, rowFieldIndices);
        var columnFieldMembers = BuildFieldMembers(source, columnFieldIndices);
        var rowLookups = BuildFieldLookups(rowFieldMembers);
        var columnLookups = BuildFieldLookups(columnFieldMembers);

        var rowTuples = new List<int[]>();
        var rowTupleIndex = new Dictionary<string, int>(StringComparer.Ordinal);
        var columnTuples = new List<int[]>();
        var columnTupleIndex = new Dictionary<string, int>(StringComparer.Ordinal);

        var rows = FilterRows(source, plan);
        var rowAssignment = new int[rows.Count];
        var columnAssignment = new int[rows.Count];
        for (var i = 0; i < rows.Count; i++)
        {
            rowAssignment[i] = ResolveTuple(rows[i], rowFieldIndices, rowLookups, rowTuples, rowTupleIndex);
            columnAssignment[i] = ResolveTuple(rows[i], columnFieldIndices, columnLookups, columnTuples, columnTupleIndex);
        }

        SortTuples(rowTuples, ref rowAssignment);
        SortTuples(columnTuples, ref columnAssignment);

        var rowCount = rowTuples.Count;
        var columnCount = columnTuples.Count;

        var cellBuckets = CreateBuckets(rowCount, columnCount, valueCount);
        var rowBuckets = CreateBuckets(rowCount, valueCount);
        var columnBuckets = CreateBuckets(columnCount, valueCount);
        var grandBuckets = new List<FormulaResult>?[valueCount];

        for (var i = 0; i < rows.Count; i++)
        {
            var row = rows[i];
            var r = rowAssignment[i];
            var c = columnAssignment[i];
            for (var v = 0; v < valueCount; v++)
            {
                var cell = row[valueFieldIndices[v]];
                (cellBuckets[r][c][v] ??= new List<FormulaResult>()).Add(cell);
                (rowBuckets[r][v] ??= new List<FormulaResult>()).Add(cell);
                (columnBuckets[c][v] ??= new List<FormulaResult>()).Add(cell);
                (grandBuckets[v] ??= new List<FormulaResult>()).Add(cell);
            }
        }

        var data = new double?[rowCount][][];
        for (var r = 0; r < rowCount; r++)
        {
            data[r] = new double?[columnCount][];
            for (var c = 0; c < columnCount; c++)
            {
                data[r][c] = AggregateRow(functions, context, cellBuckets[r][c]);
            }
        }

        var rowTotals = new double?[rowCount][];
        for (var r = 0; r < rowCount; r++)
        {
            rowTotals[r] = AggregateRow(functions, context, rowBuckets[r]);
        }

        var columnTotals = new double?[columnCount][];
        for (var c = 0; c < columnCount; c++)
        {
            columnTotals[c] = AggregateRow(functions, context, columnBuckets[c]);
        }

        var grandTotals = AggregateRow(functions, context, grandBuckets);

        for (var v = 0; v < valueCount; v++)
        {
            ApplyShowAs(plan.ValueFields[v].ShowAs, data, rowTotals, columnTotals, grandTotals, v);
        }

        return new PivotModel(
            rowFieldIndices,
            columnFieldIndices,
            valueFieldIndices,
            rowFieldMembers,
            columnFieldMembers,
            rowTuples,
            columnTuples,
            plan.ValueFields.Select(v => v.DisplayName).ToArray(),
            data,
            rowTotals,
            columnTotals,
            grandTotals);
    }

    /// <summary>
    /// Returns the distinct values of a column in Excel's default order (numbers ascending, then
    /// text ascending), rendered as display text. Used for both the model and the cache shared items
    /// so their indices align.
    /// </summary>
    /// <param name="values">The column values.</param>
    /// <returns>The ordered distinct display texts.</returns>
    public static List<string> OrderedDistinct(IEnumerable<FormulaResult> values)
    {
        var seen = new HashSet<string>(StringComparer.Ordinal);
        var members = new List<string>();
        foreach (var value in values)
        {
            if (value.Type == FormulaResultType.Empty)
            {
                continue;
            }

            var text = PivotSource.ToText(value);
            if (seen.Add(text))
            {
                members.Add(text);
            }
        }

        members.Sort(CompareMembers);
        return members;
    }

    private static void ApplyShowAs(PivotShowAs showAs, double?[][][] data, double?[][] rowTotals, double?[][] columnTotals, double?[] grandTotals, int v)
    {
        if (showAs == PivotShowAs.Normal)
        {
            return;
        }

        var rowCount = data.Length;
        var columnCount = columnTotals.Length;
        var grand = grandTotals[v];

        switch (showAs)
        {
            case PivotShowAs.PercentOfTotal:
                for (var r = 0; r < rowCount; r++)
                {
                    for (var c = 0; c < columnCount; c++)
                    {
                        data[r][c][v] = Div(data[r][c][v], grand);
                    }

                    rowTotals[r][v] = Div(rowTotals[r][v], grand);
                }

                for (var c = 0; c < columnCount; c++)
                {
                    columnTotals[c][v] = Div(columnTotals[c][v], grand);
                }

                break;

            case PivotShowAs.PercentOfColumn:
                for (var c = 0; c < columnCount; c++)
                {
                    var columnTotal = columnTotals[c][v];
                    for (var r = 0; r < rowCount; r++)
                    {
                        data[r][c][v] = Div(data[r][c][v], columnTotal);
                    }
                }

                for (var r = 0; r < rowCount; r++)
                {
                    rowTotals[r][v] = Div(rowTotals[r][v], grand);
                }

                for (var c = 0; c < columnCount; c++)
                {
                    columnTotals[c][v] = Div(columnTotals[c][v], columnTotals[c][v]);
                }

                break;

            case PivotShowAs.PercentOfRow:
                for (var r = 0; r < rowCount; r++)
                {
                    var rowTotal = rowTotals[r][v];
                    for (var c = 0; c < columnCount; c++)
                    {
                        data[r][c][v] = Div(data[r][c][v], rowTotal);
                    }
                }

                for (var c = 0; c < columnCount; c++)
                {
                    columnTotals[c][v] = Div(columnTotals[c][v], grand);
                }

                for (var r = 0; r < rowCount; r++)
                {
                    rowTotals[r][v] = Div(rowTotals[r][v], rowTotals[r][v]);
                }

                break;
        }

        grandTotals[v] = Div(grand, grand);
    }

    private static double? Div(double? numerator, double? denominator)
    {
        if (numerator is null || denominator is null || denominator.Value == 0)
        {
            return null;
        }

        return numerator.Value / denominator.Value;
    }

    private static List<FormulaResult[]> FilterRows(PivotSourceData source, PivotPlan plan)
    {
        var specs = new List<KeyValuePair<int, string>>();
        foreach (var filter in plan.Filters)
        {
            if (filter.SelectedValue is not null)
            {
                specs.Add(new KeyValuePair<int, string>(source.FieldIndex(filter.FieldName), filter.SelectedValue));
            }
        }

        if (specs.Count == 0)
        {
            return source.Rows;
        }

        var result = new List<FormulaResult[]>();
        foreach (var row in source.Rows)
        {
            var matches = true;
            foreach (var spec in specs)
            {
                if (!string.Equals(PivotSource.ToText(row[spec.Key]), spec.Value, StringComparison.Ordinal))
                {
                    matches = false;
                    break;
                }
            }

            if (matches)
            {
                result.Add(row);
            }
        }

        return result;
    }

    private static List<IList<string>> BuildFieldMembers(PivotSourceData source, int[] fieldIndices)
    {
        var members = new List<IList<string>>(fieldIndices.Length);
        foreach (var index in fieldIndices)
        {
            members.Add(OrderedDistinct(source.Column(index)));
        }

        return members;
    }

    private static List<Dictionary<string, int>> BuildFieldLookups(List<IList<string>> fieldMembers)
    {
        var lookups = new List<Dictionary<string, int>>(fieldMembers.Count);
        foreach (var members in fieldMembers)
        {
            var lookup = new Dictionary<string, int>(StringComparer.Ordinal);
            for (var i = 0; i < members.Count; i++)
            {
                lookup[members[i]] = i;
            }

            lookups.Add(lookup);
        }

        return lookups;
    }

    private static int ResolveTuple(
        FormulaResult[] row,
        int[] fieldIndices,
        List<Dictionary<string, int>> lookups,
        List<int[]> tuples,
        Dictionary<string, int> tupleIndex)
    {
        var tuple = new int[fieldIndices.Length];
        for (var f = 0; f < fieldIndices.Length; f++)
        {
            tuple[f] = lookups[f][PivotSource.ToText(row[fieldIndices[f]])];
        }

        var key = TupleKey(tuple);
        if (!tupleIndex.TryGetValue(key, out var index))
        {
            index = tuples.Count;
            tuples.Add(tuple);
            tupleIndex[key] = index;
        }

        return index;
    }

    private static void SortTuples(List<int[]> tuples, ref int[] assignment)
    {
        var order = Enumerable.Range(0, tuples.Count).ToList();
        order.Sort((a, b) => CompareTuples(tuples[a], tuples[b]));

        var remap = new int[tuples.Count];
        var sorted = new List<int[]>(tuples.Count);
        for (var newIndex = 0; newIndex < order.Count; newIndex++)
        {
            remap[order[newIndex]] = newIndex;
            sorted.Add(tuples[order[newIndex]]);
        }

        tuples.Clear();
        tuples.AddRange(sorted);

        for (var i = 0; i < assignment.Length; i++)
        {
            assignment[i] = remap[assignment[i]];
        }
    }

    private static int CompareTuples(int[] a, int[] b)
    {
        var length = System.Math.Min(a.Length, b.Length);
        for (var i = 0; i < length; i++)
        {
            if (a[i] != b[i])
            {
                return a[i].CompareTo(b[i]);
            }
        }

        return a.Length.CompareTo(b.Length);
    }

    private static string TupleKey(int[] tuple)
    {
        var builder = new StringBuilder();
        foreach (var value in tuple)
        {
            builder.Append(value).Append(',');
        }

        return builder.ToString();
    }

    private static int CompareMembers(string a, string b)
    {
        var aNum = double.TryParse(a, NumberStyles.Any, CultureInfo.InvariantCulture, out var an);
        var bNum = double.TryParse(b, NumberStyles.Any, CultureInfo.InvariantCulture, out var bn);

        if (aNum && bNum)
        {
            return an.CompareTo(bn);
        }

        if (aNum != bNum)
        {
            return aNum ? -1 : 1;
        }

        return string.CompareOrdinal(a, b);
    }

    private static List<FormulaResult>?[][][] CreateBuckets(int rows, int columns, int values)
    {
        var buckets = new List<FormulaResult>?[rows][][];
        for (var r = 0; r < rows; r++)
        {
            buckets[r] = new List<FormulaResult>?[columns][];
            for (var c = 0; c < columns; c++)
            {
                buckets[r][c] = new List<FormulaResult>?[values];
            }
        }

        return buckets;
    }

    private static List<FormulaResult>?[][] CreateBuckets(int outer, int values)
    {
        var buckets = new List<FormulaResult>?[outer][];
        for (var i = 0; i < outer; i++)
        {
            buckets[i] = new List<FormulaResult>?[values];
        }

        return buckets;
    }

    private static double?[] AggregateRow(IFunctionImplementation[] functions, CellContext context, List<FormulaResult>?[] buckets)
    {
        var result = new double?[functions.Length];
        for (var v = 0; v < functions.Length; v++)
        {
            result[v] = Aggregate(functions[v], context, buckets[v]);
        }

        return result;
    }

    private static IFunctionImplementation ResolveFunction(PivotAggregate aggregate)
    {
        var name = PivotAggregateMap.FunctionName(aggregate);
        if (!FunctionRegistry.TryGetFunction(name, out var function) || function is null)
        {
            throw new InvalidOperationException($"The built-in function '{name}' required for aggregate '{aggregate}' is not registered.");
        }

        return function;
    }

    private static double? Aggregate(IFunctionImplementation function, CellContext context, List<FormulaResult>? values)
    {
        if (values is null || values.Count == 0)
        {
            return null;
        }

        var result = function.Execute(context, values.ToArray());
        return result.Type == FormulaResultType.Number ? result.NumericValue : null;
    }
}
