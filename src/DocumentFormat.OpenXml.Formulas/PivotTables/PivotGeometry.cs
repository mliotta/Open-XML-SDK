// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

namespace DocumentFormat.OpenXml.Features.PivotTables;

/// <summary>
/// Computes the cell rectangle and sub-region positions for a general pivot (N row fields,
/// N column fields, M value fields, subtotals off) so the rendered cells and the definition's
/// <c>Location</c> agree. All coordinates are one-based.
/// </summary>
internal sealed class PivotGeometry
{
    private PivotGeometry(PivotModel model, PivotPlan plan)
    {
        var origin = PivotRef.ParseCell(plan.TargetCell);
        OriginColumn = origin.Column;
        OriginRow = origin.Row;

        RowFieldCount = model.RowFieldCount;
        ColumnFieldCount = model.ColumnFieldCount;
        ValueCount = model.ValueCount;
        RowTupleCount = model.RowTuples.Count;
        ColumnTupleCount = model.ColumnTuples.Count;
        PageFieldCount = plan.Filters.Count;

        HasValueHeaderRow = ValueCount > 1 || ColumnFieldCount == 0;
        RowLabelColumns = RowFieldCount >= 1 ? RowFieldCount : 1;
        ColumnHeaderRows = ColumnFieldCount + (HasValueHeaderRow ? 1 : 0);
        ColumnLeafCount = ColumnTupleCount * (ValueCount > 1 ? ValueCount : 1);

        ShowGrandColumn = plan.ColumnGrandTotals && ColumnFieldCount >= 1;
        ShowGrandRow = plan.RowGrandTotals && RowFieldCount >= 1;
        GrandColumnCount = ShowGrandColumn ? (ValueCount > 1 ? ValueCount : 1) : 0;
    }

    /// <summary>Gets the one-based column of the top-left corner.</summary>
    public int OriginColumn { get; }

    /// <summary>Gets the one-based row of the top-left corner.</summary>
    public int OriginRow { get; }

    /// <summary>Gets the number of row fields.</summary>
    public int RowFieldCount { get; }

    /// <summary>Gets the number of column fields.</summary>
    public int ColumnFieldCount { get; }

    /// <summary>Gets the number of value fields.</summary>
    public int ValueCount { get; }

    /// <summary>Gets the number of row tuples (data rows).</summary>
    public int RowTupleCount { get; }

    /// <summary>Gets the number of column tuples.</summary>
    public int ColumnTupleCount { get; }

    /// <summary>Gets a value indicating whether a value-name header row is rendered.</summary>
    public bool HasValueHeaderRow { get; }

    /// <summary>Gets the number of row-label columns.</summary>
    public int RowLabelColumns { get; }

    /// <summary>Gets the number of column-header rows.</summary>
    public int ColumnHeaderRows { get; }

    /// <summary>Gets the number of data (non-total) columns.</summary>
    public int ColumnLeafCount { get; }

    /// <summary>Gets a value indicating whether a grand-total column group is rendered.</summary>
    public bool ShowGrandColumn { get; }

    /// <summary>Gets a value indicating whether a grand-total row is rendered.</summary>
    public bool ShowGrandRow { get; }

    /// <summary>Gets the number of grand-total columns (one per value field).</summary>
    public int GrandColumnCount { get; }

    /// <summary>Gets the number of report-filter (page) fields rendered above the body.</summary>
    public int PageFieldCount { get; }

    /// <summary>Gets the one-based row of the body's top-left corner (below any page-field rows).</summary>
    public int BodyOriginRow => OriginRow + (PageFieldCount > 0 ? PageFieldCount + 1 : 0);

    /// <summary>Gets the number of rows from the ref top down to the first column-header row.</summary>
    public int HeaderRowOffset => BodyOriginRow - OriginRow;

    /// <summary>Gets the one-based row of the value-name header (valid only when rendered).</summary>
    public int ValueHeaderRow => BodyOriginRow + ColumnFieldCount;

    /// <summary>Gets the one-based row of the first data row.</summary>
    public int FirstDataRow => BodyOriginRow + ColumnHeaderRows;

    /// <summary>Gets the one-based column of the first data column.</summary>
    public int FirstDataColumn => OriginColumn + RowLabelColumns;

    /// <summary>Gets the one-based column of the first grand-total column.</summary>
    public int FirstGrandColumn => FirstDataColumn + ColumnLeafCount;

    /// <summary>Gets the one-based row of the grand-total row (valid only when rendered).</summary>
    public int GrandRow => FirstDataRow + RowTupleCount;

    /// <summary>Gets the one-based last column of the pivot rectangle.</summary>
    public int LastColumn => FirstDataColumn + ColumnLeafCount + GrandColumnCount - 1;

    /// <summary>Gets the one-based last row of the pivot rectangle.</summary>
    public int LastRow => FirstDataRow + RowTupleCount - 1 + (ShowGrandRow ? 1 : 0);

    /// <summary>Gets the bounding range reference, e.g. <c>"A3:E12"</c>.</summary>
    public string Reference => PivotRef.Range(OriginColumn, OriginRow, LastColumn, LastRow);

    /// <summary>Gets the column of the data cell for a column tuple and value field.</summary>
    /// <param name="columnTupleIndex">The column tuple index.</param>
    /// <param name="valueIndex">The value field index.</param>
    /// <returns>The one-based column.</returns>
    public int DataColumn(int columnTupleIndex, int valueIndex)
    {
        var leaf = ValueCount > 1 ? (columnTupleIndex * ValueCount) + valueIndex : columnTupleIndex;
        return FirstDataColumn + leaf;
    }

    /// <summary>Builds the geometry for a model placed at the target cell.</summary>
    /// <param name="model">The computed model.</param>
    /// <param name="plan">The plan.</param>
    /// <returns>The geometry.</returns>
    public static PivotGeometry Create(PivotModel model, PivotPlan plan) => new(model, plan);
}
