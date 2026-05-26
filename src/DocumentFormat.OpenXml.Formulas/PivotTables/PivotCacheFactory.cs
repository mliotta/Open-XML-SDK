// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Collections.Generic;
using System.Globalization;
using System.Linq;

using DocumentFormat.OpenXml.Features.FormulaEvaluation;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace DocumentFormat.OpenXml.Features.PivotTables;

/// <summary>Builds the pivot cache definition and records and links them into the workbook.</summary>
internal static class PivotCacheFactory
{
    /// <summary>Builds the cache for the source data.</summary>
    /// <param name="workbookPart">The owning workbook part.</param>
    /// <param name="sourceWorksheetPart">The worksheet holding the source range.</param>
    /// <param name="source">The parsed source data.</param>
    /// <param name="plan">The validated plan.</param>
    /// <returns>The cache artifacts.</returns>
    public static PivotCacheInfo Build(WorkbookPart workbookPart, WorksheetPart sourceWorksheetPart, PivotSourceData source, PivotPlan plan)
    {
        var width = source.Headers.Length;
        var valueIndex = source.FieldIndex(plan.ValueFields[0].FieldName);
        var valueIsNumeric = IsNumericColumn(source.Column(valueIndex));

        var members = new IList<string>?[width];
        var memberLookup = new Dictionary<string, int>?[width];

        var cacheFields = new CacheFields { Count = (uint)width };
        for (var f = 0; f < width; f++)
        {
            var enumerated = !(f == valueIndex && valueIsNumeric);
            var field = new CacheField { Name = source.Headers[f] };

            if (enumerated)
            {
                var ordered = PivotModel.OrderedDistinct(source.Column(f));
                members[f] = ordered;
                memberLookup[f] = BuildLookup(ordered);
                field.Append(BuildEnumeratedSharedItems(ordered));
            }
            else
            {
                field.Append(BuildNumericSharedItems(source.Column(f)));
            }

            cacheFields.Append(field);
        }

        var cacheDefinition = new PivotCacheDefinition
        {
            SaveData = true,
            RefreshOnLoad = false,
            CacheSource = new CacheSource(new WorksheetSource { Reference = source.Reference, Sheet = source.SheetName })
            {
                Type = SourceValues.Worksheet,
            },
        };
        cacheDefinition.Append(cacheFields);

        var records = BuildRecords(source, valueIndex, valueIsNumeric, memberLookup);

        var cacheDefinitionPart = workbookPart.AddNewPart<PivotTableCacheDefinitionPart>();
        cacheDefinitionPart.PivotCacheDefinition = cacheDefinition;
        var recordsPart = cacheDefinitionPart.AddNewPart<PivotTableCacheRecordsPart>();
        recordsPart.PivotCacheRecords = records;
        cacheDefinition.Id = cacheDefinitionPart.GetIdOfPart(recordsPart);

        var cacheId = LinkWorkbookCache(workbookPart, cacheDefinitionPart);
        return new PivotCacheInfo(cacheDefinitionPart, cacheId, members);
    }

    private static SharedItems BuildEnumeratedSharedItems(IList<string> members)
    {
        var shared = new SharedItems();
        var allNumeric = true;
        foreach (var member in members)
        {
            if (double.TryParse(member, NumberStyles.Any, CultureInfo.InvariantCulture, out var number))
            {
                shared.Append(new NumberItem { Val = number });
            }
            else
            {
                allNumeric = false;
                shared.Append(new StringItem { Val = member });
            }
        }

        if (allNumeric && members.Count > 0)
        {
            shared.ContainsString = false;
            shared.ContainsNumber = true;
        }

        return shared;
    }

    private static SharedItems BuildNumericSharedItems(IEnumerable<FormulaResult> values)
    {
        var numbers = values.Where(v => v.Type == FormulaResultType.Number).Select(v => v.NumericValue).ToList();
        var shared = new SharedItems
        {
            ContainsString = false,
            ContainsNumber = true,
        };

        if (numbers.Count > 0)
        {
            shared.MinValue = numbers.Min();
            shared.MaxValue = numbers.Max();
            if (numbers.All(n => n == System.Math.Floor(n)))
            {
                shared.ContainsInteger = true;
            }
        }

        return shared;
    }

    private static PivotCacheRecords BuildRecords(PivotSourceData source, int valueIndex, bool valueIsNumeric, IList<Dictionary<string, int>?> memberLookup)
    {
        var records = new PivotCacheRecords { Count = (uint)source.Rows.Count };
        foreach (var row in source.Rows)
        {
            var record = new PivotCacheRecord();
            for (var f = 0; f < source.Headers.Length; f++)
            {
                var cell = row[f];
                if (f == valueIndex && valueIsNumeric)
                {
                    record.Append(cell.Type == FormulaResultType.Number
                        ? new NumberItem { Val = cell.NumericValue }
                        : (OpenXmlElement)new MissingItem());
                }
                else if (cell.Type == FormulaResultType.Empty)
                {
                    record.Append(new MissingItem());
                }
                else
                {
                    var index = memberLookup[f]![PivotSource.ToText(cell)];
                    record.Append(new FieldItem { Val = (uint)index });
                }
            }

            records.Append(record);
        }

        return records;
    }

    private static uint LinkWorkbookCache(WorkbookPart workbookPart, PivotTableCacheDefinitionPart cacheDefinitionPart)
    {
        var workbook = workbookPart.Workbook ?? throw new System.InvalidOperationException("The workbook has no content.");
        var caches = workbook.GetFirstChild<PivotCaches>();
        if (caches is null)
        {
            caches = new PivotCaches();
            var anchor = (OpenXmlElement?)workbook.GetFirstChild<CalculationProperties>() ?? workbook.GetFirstChild<Sheets>();
            if (anchor is not null)
            {
                workbook.InsertAfter(caches, anchor);
            }
            else
            {
                workbook.Append(caches);
            }
        }

        var cacheId = caches.Elements<PivotCache>()
            .Select(c => c.CacheId?.Value ?? 0u)
            .DefaultIfEmpty(0u)
            .Max() + 1u;

        caches.Append(new PivotCache
        {
            CacheId = cacheId,
            Id = workbookPart.GetIdOfPart(cacheDefinitionPart),
        });

        return cacheId;
    }

    private static bool IsNumericColumn(IEnumerable<FormulaResult> values)
    {
        var any = false;
        foreach (var value in values)
        {
            if (value.Type == FormulaResultType.Empty)
            {
                continue;
            }

            any = true;
            if (value.Type != FormulaResultType.Number)
            {
                return false;
            }
        }

        return any;
    }

    private static Dictionary<string, int> BuildLookup(IList<string> members)
    {
        var lookup = new Dictionary<string, int>(System.StringComparer.Ordinal);
        for (var i = 0; i < members.Count; i++)
        {
            lookup[members[i]] = i;
        }

        return lookup;
    }
}
