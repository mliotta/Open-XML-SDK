// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the PIVOTBY function.
/// PIVOTBY(row_fields, col_fields, values, function, [field_headers], [row_total_depth], [col_total_depth], [sort_order], [filter_array])
/// Pivots data by grouping values into a cross-tabulation.
///
/// Phase 0 Implementation:
/// - Simplified cross-tabulation with basic grouping
/// - Supports common aggregation functions (SUM, AVERAGE, COUNT, MAX, MIN)
/// - Returns first cell of pivot table
/// - Full array support and all optional parameters require engine enhancements
/// </summary>
public sealed class PivotByFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly PivotByFunction Instance = new PivotByFunction();

    private PivotByFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "PIVOTBY";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        // Minimum args: row_fields, col_fields, values, function
        if (args.Length < 4)
        {
            return CellValue.Error("#VALUE!");
        }

        // Parse function parameter (should be a number representing aggregation type)
        // For Phase 0, we expect: 1=SUM, 2=AVERAGE, 3=COUNT, 4=MAX, 5=MIN
        var functionArg = args[args.Length - 1];
        if (functionArg.Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }

        var functionType = (int)functionArg.NumericValue;
        if (functionType < 1 || functionType > 5)
        {
            return CellValue.Error("#VALUE!");
        }

        // Calculate how many cells are in row_fields, col_fields, and values
        // Heuristic: Split remaining args evenly among the three arrays
        var totalDataArgs = args.Length - 1; // Exclude function parameter
        var sectionSize = totalDataArgs / 3;

        if (sectionSize == 0)
        {
            return CellValue.Error("#VALUE!");
        }

        var rowFieldCount = sectionSize;
        var colFieldCount = sectionSize;
        var valueCount = totalDataArgs - rowFieldCount - colFieldCount;

        // Check for errors in input arrays
        for (var i = 0; i < totalDataArgs; i++)
        {
            if (args[i].IsError)
            {
                return args[i];
            }
        }

        // Build pivot table structure: row key -> col key -> values
        var pivot = new Dictionary<string, Dictionary<string, List<CellValue>>>();

        // Simplified: assume each position represents one data point
        var numDataPoints = System.Math.Min(System.Math.Min(rowFieldCount, colFieldCount), valueCount);

        for (var i = 0; i < numDataPoints; i++)
        {
            var rowKey = CreateKey(args[i]);
            var colKey = CreateKey(args[rowFieldCount + i]);
            var value = args[rowFieldCount + colFieldCount + i];

            if (!pivot.ContainsKey(rowKey))
            {
                pivot[rowKey] = new Dictionary<string, List<CellValue>>();
            }

            if (!pivot[rowKey].ContainsKey(colKey))
            {
                pivot[rowKey][colKey] = new List<CellValue>();
            }

            pivot[rowKey][colKey].Add(value);
        }

        if (pivot.Count == 0)
        {
            return CellValue.Error("#CALC!");
        }

        // Get first row and first column
        Dictionary<string, List<CellValue>> firstRow = null;
        foreach (var row in pivot.Values)
        {
            firstRow = row;
            break;
        }

        if (firstRow == null || firstRow.Count == 0)
        {
            return CellValue.Error("#CALC!");
        }

        List<CellValue> firstCell = null;
        foreach (var cell in firstRow.Values)
        {
            firstCell = cell;
            break;
        }

        if (firstCell == null)
        {
            return CellValue.Error("#CALC!");
        }

        // Apply aggregation to first cell
        var result = ApplyAggregation(firstCell, functionType);

        return result;
    }

    private static string CreateKey(CellValue value)
    {
        switch (value.Type)
        {
            case CellValueType.Number:
                return "N:" + value.NumericValue.ToString();
            case CellValueType.Text:
                return "T:" + value.StringValue;
            case CellValueType.Boolean:
                return "B:" + value.BoolValue.ToString();
            case CellValueType.Empty:
                return "E:";
            case CellValueType.Error:
                return "ERR:" + value.ErrorValue;
            default:
                return "?";
        }
    }

    private static CellValue ApplyAggregation(List<CellValue> values, int functionType)
    {
        if (values.Count == 0)
        {
            return CellValue.Error("#CALC!");
        }

        // Extract numeric values
        var numbers = new List<double>();
        foreach (var val in values)
        {
            if (val.Type == CellValueType.Number)
            {
                numbers.Add(val.NumericValue);
            }
            else if (val.IsError)
            {
                return val; // Propagate error
            }
        }

        if (numbers.Count == 0 && functionType != 3) // COUNT can work with non-numeric
        {
            return CellValue.Error("#VALUE!");
        }

        switch (functionType)
        {
            case 1: // SUM
                {
                    var sum = 0.0;
                    foreach (var num in numbers)
                    {
                        sum += num;
                    }
                    return CellValue.FromNumber(sum);
                }
            case 2: // AVERAGE
                {
                    var sum = 0.0;
                    foreach (var num in numbers)
                    {
                        sum += num;
                    }
                    return CellValue.FromNumber(sum / numbers.Count);
                }
            case 3: // COUNT
                return CellValue.FromNumber(values.Count);
            case 4: // MAX
                {
                    var max = double.MinValue;
                    foreach (var num in numbers)
                    {
                        if (num > max)
                        {
                            max = num;
                        }
                    }
                    return CellValue.FromNumber(max);
                }
            case 5: // MIN
                {
                    var min = double.MaxValue;
                    foreach (var num in numbers)
                    {
                        if (num < min)
                        {
                            min = num;
                        }
                    }
                    return CellValue.FromNumber(min);
                }
            default:
                return CellValue.Error("#VALUE!");
        }
    }
}
