// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Collections.Generic;

namespace DocumentFormat.OpenXml.Features.PivotTables;

/// <summary>An immutable, validated description of the pivot table to build.</summary>
internal sealed class PivotPlan
{
    /// <summary>Initializes a new instance of the <see cref="PivotPlan"/> class.</summary>
    /// <param name="name">The pivot table name.</param>
    /// <param name="rowFields">Row-axis field names.</param>
    /// <param name="columnFields">Column-axis field names.</param>
    /// <param name="filters">Filter (page) fields.</param>
    /// <param name="valueFields">Value fields.</param>
    /// <param name="layout">The label layout.</param>
    /// <param name="rowGrandTotals">Whether to show a grand-total row.</param>
    /// <param name="columnGrandTotals">Whether to show a grand-total column.</param>
    /// <param name="sourceReference">The source range reference.</param>
    /// <param name="targetCell">The top-left destination cell.</param>
    public PivotPlan(
        string name,
        IList<string> rowFields,
        IList<string> columnFields,
        IList<PivotFilter> filters,
        IList<PivotValueField> valueFields,
        PivotLayout layout,
        bool rowGrandTotals,
        bool columnGrandTotals,
        string sourceReference,
        string targetCell)
    {
        Name = name;
        Filters = filters;
        RowFields = rowFields;
        ColumnFields = columnFields;
        ValueFields = valueFields;
        Layout = layout;
        RowGrandTotals = rowGrandTotals;
        ColumnGrandTotals = columnGrandTotals;
        SourceReference = sourceReference;
        TargetCell = targetCell;
    }

    /// <summary>Gets the pivot table name.</summary>
    public string Name { get; }

    /// <summary>Gets the row-axis field names.</summary>
    public IList<string> RowFields { get; }

    /// <summary>Gets the column-axis field names.</summary>
    public IList<string> ColumnFields { get; }

    /// <summary>Gets the filter (page) fields and their selections.</summary>
    public IList<PivotFilter> Filters { get; }

    /// <summary>Gets the value fields.</summary>
    public IList<PivotValueField> ValueFields { get; }

    /// <summary>Gets the label layout.</summary>
    public PivotLayout Layout { get; }

    /// <summary>Gets a value indicating whether a grand-total row is shown.</summary>
    public bool RowGrandTotals { get; }

    /// <summary>Gets a value indicating whether a grand-total column is shown.</summary>
    public bool ColumnGrandTotals { get; }

    /// <summary>Gets the source range reference.</summary>
    public string SourceReference { get; }

    /// <summary>Gets the top-left destination cell.</summary>
    public string TargetCell { get; }
}
