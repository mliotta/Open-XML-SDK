// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace DocumentFormat.OpenXml.Features.PivotTables;

/// <summary>
/// Builds a native OOXML pivot table whose results are precomputed (so the values are visible
/// without recalculation in Excel) and writes the cache, definition, and rendered cells into the
/// host package.
/// </summary>
/// <remarks>
/// <para>Create an instance with <see cref="FromRange"/>, configure it fluently, then call
/// <see cref="PlaceAt"/> to materialize the pivot table on a destination worksheet.</para>
/// <para>Any number of row fields, column fields, and value fields is supported. Subtotals are not
/// emitted; grand totals are configurable. When more than one value field is present, the value
/// selector is placed on the column axis.</para>
/// </remarks>
public sealed class PivotTableBuilder
{
    private readonly WorksheetPart _sourceWorksheetPart;
    private readonly string _sourceReference;
    private readonly List<string> _rowFields = new();
    private readonly List<string> _columnFields = new();
    private readonly List<PivotFilter> _filters = new();
    private readonly List<PivotValueField> _valueFields = new();

    private string _name = "PivotTable1";
    private PivotLayout _layout = PivotLayout.Compact;
    private bool _rowGrandTotals = true;
    private bool _columnGrandTotals = true;

    private PivotTableBuilder(WorksheetPart sourceWorksheetPart, string sourceReference)
    {
        _sourceWorksheetPart = sourceWorksheetPart;
        _sourceReference = sourceReference;
    }

    /// <summary>
    /// Begins building a pivot table from a worksheet range whose first row holds the field headers.
    /// </summary>
    /// <param name="sourceWorksheetPart">The worksheet part containing the source data.</param>
    /// <param name="sourceReference">The source range, e.g. <c>"A1:D100"</c>.</param>
    /// <returns>A new <see cref="PivotTableBuilder"/>.</returns>
    public static PivotTableBuilder FromRange(WorksheetPart sourceWorksheetPart, string sourceReference)
    {
        if (sourceWorksheetPart is null)
        {
            throw new ArgumentNullException(nameof(sourceWorksheetPart));
        }

        if (IsBlank(sourceReference))
        {
            throw new ArgumentException("A source range is required.", nameof(sourceReference));
        }

        return new PivotTableBuilder(sourceWorksheetPart, sourceReference);
    }

    /// <summary>Sets the pivot table name.</summary>
    /// <param name="name">The name shown in Excel and used for the part.</param>
    /// <returns>This builder.</returns>
    public PivotTableBuilder Named(string name)
    {
        _name = IsBlank(name) ? _name : name;
        return this;
    }

    /// <summary>Adds a field to the row axis.</summary>
    /// <param name="fieldName">A header name from the source range.</param>
    /// <returns>This builder.</returns>
    public PivotTableBuilder Row(string fieldName)
    {
        _rowFields.Add(RequireField(fieldName));
        return this;
    }

    /// <summary>Adds a field to the column axis.</summary>
    /// <param name="fieldName">A header name from the source range.</param>
    /// <returns>This builder.</returns>
    public PivotTableBuilder Column(string fieldName)
    {
        _columnFields.Add(RequireField(fieldName));
        return this;
    }

    /// <summary>Adds a field to the report-filter (page) axis.</summary>
    /// <param name="fieldName">A header name from the source range.</param>
    /// <param name="selectedValue">The value to filter to; omit (or null) to include all values.</param>
    /// <returns>This builder.</returns>
    public PivotTableBuilder Filter(string fieldName, string? selectedValue = null)
    {
        _filters.Add(new PivotFilter(RequireField(fieldName), selectedValue));
        return this;
    }

    /// <summary>Adds a value (data) field with the given aggregate.</summary>
    /// <param name="fieldName">A header name from the source range.</param>
    /// <param name="aggregate">The consolidation function. Defaults to <see cref="PivotAggregate.Sum"/>.</param>
    /// <param name="displayName">Optional caption; defaults to e.g. "Sum of Sales".</param>
    /// <param name="showAs">How the value is displayed. Defaults to <see cref="PivotShowAs.Normal"/>.</param>
    /// <returns>This builder.</returns>
    public PivotTableBuilder Value(string fieldName, PivotAggregate aggregate = PivotAggregate.Sum, string? displayName = null, PivotShowAs showAs = PivotShowAs.Normal)
    {
        var field = RequireField(fieldName);
        _valueFields.Add(new PivotValueField(field, aggregate, displayName ?? PivotAggregateMap.DisplayName(aggregate, field), showAs));
        return this;
    }

    /// <summary>Sets the row/column label layout. Defaults to <see cref="PivotLayout.Compact"/>.</summary>
    /// <param name="layout">The layout.</param>
    /// <returns>This builder.</returns>
    public PivotTableBuilder Layout(PivotLayout layout)
    {
        _layout = layout;
        return this;
    }

    /// <summary>Enables or disables grand totals. Both are enabled by default.</summary>
    /// <param name="rows">Whether to show a grand-total row.</param>
    /// <param name="columns">Whether to show a grand-total column.</param>
    /// <returns>This builder.</returns>
    public PivotTableBuilder GrandTotals(bool rows = true, bool columns = true)
    {
        _rowGrandTotals = rows;
        _columnGrandTotals = columns;
        return this;
    }

    /// <summary>
    /// Computes the pivot table and writes the cache, definition, and rendered cells, returning the
    /// created <see cref="PivotTablePart"/>.
    /// </summary>
    /// <param name="targetWorksheetPart">The worksheet that will host the rendered pivot table.</param>
    /// <param name="targetCell">The top-left cell of the pivot table, e.g. <c>"A1"</c>.</param>
    /// <returns>The created pivot table part.</returns>
    public PivotTablePart PlaceAt(WorksheetPart targetWorksheetPart, string targetCell)
    {
        if (targetWorksheetPart is null)
        {
            throw new ArgumentNullException(nameof(targetWorksheetPart));
        }

        if (IsBlank(targetCell))
        {
            throw new ArgumentException("A target cell is required.", nameof(targetCell));
        }

        ValidateScope();

        var plan = new PivotPlan(
            _name,
            _rowFields,
            _columnFields,
            _filters,
            _valueFields,
            _layout,
            _rowGrandTotals,
            _columnGrandTotals,
            _sourceReference,
            targetCell);

        var workbookPart = ResolveWorkbookPart(_sourceWorksheetPart);
        var worksheet = _sourceWorksheetPart.Worksheet ?? throw new InvalidOperationException("The source worksheet has no content.");
        var context = new CellContext(worksheet, workbookPart.SharedStringTablePart);
        var source = PivotSource.Read(_sourceWorksheetPart, workbookPart, _sourceReference, context);
        var model = PivotModel.Compute(source, plan, context);

        var cache = PivotCacheFactory.Build(workbookPart, _sourceWorksheetPart, source, plan);
        var part = PivotDefinitionFactory.Build(targetWorksheetPart, source, cache, model, plan);
        PivotSheetWriter.Write(targetWorksheetPart, model, plan);

        return part;
    }

    private static WorkbookPart ResolveWorkbookPart(WorksheetPart worksheetPart)
    {
        if (worksheetPart.OpenXmlPackage is SpreadsheetDocument document && document.WorkbookPart is not null)
        {
            return document.WorkbookPart;
        }

        throw new InvalidOperationException("The source worksheet is not part of a SpreadsheetDocument with a workbook.");
    }

    private string RequireField(string fieldName)
    {
        if (IsBlank(fieldName))
        {
            throw new ArgumentException("A field name is required.", nameof(fieldName));
        }

        return fieldName;
    }

    private static bool IsBlank(string? value) => string.IsNullOrEmpty(value) || value!.Trim().Length == 0;

    private void ValidateScope()
    {
        if (_valueFields.Count < 1)
        {
            throw new InvalidOperationException("At least one value field is required.");
        }
    }
}
