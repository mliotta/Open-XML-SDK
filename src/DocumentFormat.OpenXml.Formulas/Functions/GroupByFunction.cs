// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the GROUPBY function.
/// GROUPBY(row_fields, values, function, [field_headers], [total_depth], [sort_order], [filter_array])
/// Groups rows by values in specified columns and performs aggregation.
///
/// Phase 0 Implementation:
/// - Simplified to group by first unique value in row_fields
/// - Supports common aggregation functions (SUM, AVERAGE, COUNT, MAX, MIN)
/// - Returns first result value
/// - Full array support and all optional parameters require engine enhancements
/// </summary>
public sealed class GroupByFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly GroupByFunction Instance = new GroupByFunction();

    private GroupByFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "GROUPBY";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        // Minimum args: row_fields, values, function
        if (args.Length < 3)
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

        // Calculate how many cells are in row_fields and values
        // Heuristic: Split remaining args evenly between row_fields and values
        var totalDataArgs = args.Length - 1; // Exclude function parameter
        var fieldCount = totalDataArgs / 2;
        var valueCount = totalDataArgs - fieldCount;

        if (fieldCount == 0 || valueCount == 0)
        {
            return CellValue.Error("#VALUE!");
        }

        // Check for errors in input arrays
        for (var i = 0; i < totalDataArgs; i++)
        {
            if (args[i].IsError)
            {
                return args[i];
            }
        }

        // Group by first value in row_fields
        var groups = new Dictionary<string, List<CellValue>>();

        // Determine dimensions - assume row_fields and values are 1D arrays
        var numRows = fieldCount; // Simplified: one row per field value

        for (var i = 0; i < numRows; i++)
        {
            if (i >= fieldCount)
            {
                break;
            }

            var key = CreateKey(args[i]);
            var valueIndex = fieldCount + i;

            if (valueIndex >= totalDataArgs)
            {
                break;
            }

            var value = args[valueIndex];

            if (!groups.ContainsKey(key))
            {
                groups[key] = new List<CellValue>();
            }

            groups[key].Add(value);
        }

        if (groups.Count == 0)
        {
            return CellValue.Error("#CALC!");
        }

        // Apply aggregation function to first group
        List<CellValue> firstGroup = null;
        foreach (var group in groups.Values)
        {
            firstGroup = group;
            break;
        }

        if (firstGroup == null)
        {
            return CellValue.Error("#CALC!");
        }

        var result = ApplyAggregation(firstGroup, functionType);

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
