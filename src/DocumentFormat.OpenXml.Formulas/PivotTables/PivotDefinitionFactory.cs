// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace DocumentFormat.OpenXml.Features.PivotTables;

/// <summary>
/// Builds the <see cref="PivotTableDefinition"/> (fields, axes, nested items, layout) for a general
/// pivot and attaches it to the host worksheet. Subtotals are disabled; the data field selector is
/// placed on the column axis when there is more than one value field.
/// </summary>
internal static class PivotDefinitionFactory
{
    /// <summary>Creates the pivot table part and definition.</summary>
    /// <param name="targetWorksheetPart">The worksheet that hosts the pivot table.</param>
    /// <param name="source">The parsed source data (for field indices).</param>
    /// <param name="cache">The cache artifacts.</param>
    /// <param name="model">The computed grid.</param>
    /// <param name="plan">The validated plan.</param>
    /// <returns>The created pivot table part.</returns>
    public static PivotTablePart Build(WorksheetPart targetWorksheetPart, PivotSourceData source, PivotCacheInfo cache, PivotModel model, PivotPlan plan)
    {
        var part = targetWorksheetPart.AddNewPart<PivotTablePart>();
        part.AddPart(cache.CacheDefinitionPart);

        var geometry = PivotGeometry.Create(model, plan);
        var compact = plan.Layout == PivotLayout.Compact;
        var outline = plan.Layout != PivotLayout.Tabular;
        var multipleValues = model.ValueCount > 1;

        var rowSet = new HashSet<int>(model.RowFieldIndices);
        var columnSet = new HashSet<int>(model.ColumnFieldIndices);
        var valueSet = new HashSet<int>(model.ValueFieldIndices);
        var filterIndices = plan.Filters.Select(f => source.FieldIndex(f.FieldName)).ToArray();
        var filterSet = new HashSet<int>(filterIndices);

        var definition = new PivotTableDefinition
        {
            Name = plan.Name,
            CacheId = cache.CacheId,
            DataCaption = "Values",
            RowGrandTotals = geometry.ShowGrandColumn,
            ColumnGrandTotals = geometry.ShowGrandRow,
            Compact = compact,
            CompactData = compact,
            Outline = outline,
        };

        definition.Append(new Location
        {
            Reference = geometry.Reference,
            FirstHeaderRow = (uint)System.Math.Max(geometry.HeaderRowOffset, 1),
            FirstDataRow = (uint)(geometry.HeaderRowOffset + geometry.ColumnHeaderRows),
            FirstDataColumn = (uint)geometry.RowLabelColumns,
        });

        definition.Append(BuildPivotFields(source, cache, compact, outline, rowSet, columnSet, filterSet, valueSet));

        if (model.RowFieldCount >= 1)
        {
            definition.Append(new RowFields(model.RowFieldIndices.Select(i => (OpenXmlElement)new Field { Index = i })) { Count = (uint)model.RowFieldCount });
            definition.Append(BuildRowItems(model, geometry));
        }

        var columnFieldElements = model.ColumnFieldIndices.Select(i => (OpenXmlElement)new Field { Index = i }).ToList();
        if (multipleValues)
        {
            columnFieldElements.Add(new Field { Index = -2 });
        }

        if (columnFieldElements.Count > 0)
        {
            definition.Append(new ColumnFields(columnFieldElements) { Count = (uint)columnFieldElements.Count });
            definition.Append(BuildColumnItems(model, geometry));
        }

        if (plan.Filters.Count > 0)
        {
            definition.Append(BuildPageFields(source, cache, plan));
        }

        definition.Append(BuildDataFields(model, plan));

        part.PivotTableDefinition = definition;
        return part;
    }

    private static PivotFields BuildPivotFields(
        PivotSourceData source,
        PivotCacheInfo cache,
        bool compact,
        bool outline,
        HashSet<int> rowSet,
        HashSet<int> columnSet,
        HashSet<int> filterSet,
        HashSet<int> valueSet)
    {
        var width = source.Headers.Length;
        var pivotFields = new PivotFields { Count = (uint)width };
        for (var f = 0; f < width; f++)
        {
            var field = new PivotField { Compact = compact, Outline = outline };
            if (rowSet.Contains(f))
            {
                field.Axis = PivotTableAxisValues.AxisRow;
                field.DefaultSubtotal = false;
                field.Append(BuildItems(cache.FieldMembers[f]!.Count));
            }
            else if (columnSet.Contains(f))
            {
                field.Axis = PivotTableAxisValues.AxisColumn;
                field.DefaultSubtotal = false;
                field.Append(BuildItems(cache.FieldMembers[f]!.Count));
            }
            else if (filterSet.Contains(f))
            {
                field.Axis = PivotTableAxisValues.AxisPage;
                field.DefaultSubtotal = false;
                field.Append(BuildItems(cache.FieldMembers[f]!.Count));
            }
            else if (valueSet.Contains(f))
            {
                field.DataField = true;
            }

            pivotFields.Append(field);
        }

        return pivotFields;
    }

    private static Items BuildItems(int memberCount)
    {
        var items = new Items { Count = (uint)memberCount };
        for (var i = 0; i < memberCount; i++)
        {
            items.Append(new Item { Index = (uint)i, ItemType = ItemValues.Data });
        }

        return items;
    }

    private static RowItems BuildRowItems(PivotModel model, PivotGeometry geometry)
    {
        var items = new RowItems();
        var count = 0;
        for (var rt = 0; rt < model.RowTuples.Count; rt++)
        {
            var repeat = rt == 0 ? 0 : FirstDifferingLevel(model.RowTuples[rt - 1], model.RowTuples[rt]);
            var item = new RowItem();
            if (repeat > 0)
            {
                item.RepeatedItemCount = (uint)repeat;
            }

            for (var f = repeat; f < model.RowFieldCount; f++)
            {
                item.Append(new MemberPropertyIndex { Val = model.RowTuples[rt][f] });
            }

            items.Append(item);
            count++;
        }

        if (geometry.ShowGrandRow)
        {
            items.Append(new RowItem(new MemberPropertyIndex { Val = 0 }) { ItemType = ItemValues.Grand });
            count++;
        }

        items.Count = (uint)count;
        return items;
    }

    private static ColumnItems BuildColumnItems(PivotModel model, PivotGeometry geometry)
    {
        var items = new ColumnItems();
        var multipleValues = model.ValueCount > 1;
        var valueLoop = multipleValues ? model.ValueCount : 1;
        var count = 0;

        for (var ct = 0; ct < model.ColumnTuples.Count; ct++)
        {
            var baseRepeat = ct == 0 ? 0 : FirstDifferingLevel(model.ColumnTuples[ct - 1], model.ColumnTuples[ct]);
            for (var v = 0; v < valueLoop; v++)
            {
                var item = new RowItem();
                if (v == 0)
                {
                    if (baseRepeat > 0)
                    {
                        item.RepeatedItemCount = (uint)baseRepeat;
                    }

                    for (var f = baseRepeat; f < model.ColumnFieldCount; f++)
                    {
                        item.Append(new MemberPropertyIndex { Val = model.ColumnTuples[ct][f] });
                    }
                }
                else
                {
                    item.RepeatedItemCount = (uint)model.ColumnFieldCount;
                }

                if (multipleValues)
                {
                    item.Index = (uint)v;
                }

                items.Append(item);
                count++;
            }
        }

        if (geometry.ShowGrandColumn)
        {
            for (var v = 0; v < valueLoop; v++)
            {
                var item = new RowItem(new MemberPropertyIndex { Val = 0 }) { ItemType = ItemValues.Grand };
                if (multipleValues)
                {
                    item.Index = (uint)v;
                }

                items.Append(item);
                count++;
            }
        }

        items.Count = (uint)count;
        return items;
    }

    private static PageFields BuildPageFields(PivotSourceData source, PivotCacheInfo cache, PivotPlan plan)
    {
        var pageFields = new PageFields { Count = (uint)plan.Filters.Count };
        foreach (var filter in plan.Filters)
        {
            var index = source.FieldIndex(filter.FieldName);
            var pageField = new PageField { Field = index };
            if (filter.SelectedValue is not null)
            {
                var members = cache.FieldMembers[index];
                var itemIndex = members?.IndexOf(filter.SelectedValue) ?? -1;
                if (itemIndex >= 0)
                {
                    pageField.Item = (uint)itemIndex;
                }
            }

            pageFields.Append(pageField);
        }

        return pageFields;
    }

    private static DataFields BuildDataFields(PivotModel model, PivotPlan plan)
    {
        var dataFields = new DataFields { Count = (uint)model.ValueCount };
        for (var v = 0; v < model.ValueCount; v++)
        {
            var dataField = new DataField
            {
                Name = plan.ValueFields[v].DisplayName,
                Field = (uint)model.ValueFieldIndices[v],
                Subtotal = PivotAggregateMap.Subtotal(plan.ValueFields[v].Aggregate),
            };

            var showDataAs = PivotAggregateMap.ShowDataAs(plan.ValueFields[v].ShowAs);
            if (showDataAs.HasValue)
            {
                dataField.ShowDataAs = showDataAs.Value;
                dataField.NumberFormatId = 10u;
            }

            dataFields.Append(dataField);
        }

        return dataFields;
    }

    private static int FirstDifferingLevel(int[] previous, int[] current)
    {
        for (var i = 0; i < current.Length; i++)
        {
            if (i >= previous.Length || previous[i] != current[i])
            {
                return i;
            }
        }

        return current.Length;
    }
}
