// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the FREQUENCY function.
/// FREQUENCY(data_array, bins_array) - Returns a frequency distribution (counts per bin).
/// Note: Phase 0 implementation returns the count of all data points as a single value.
/// Full array support will be added in a future phase.
/// </summary>
public sealed class FrequencyFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly FrequencyFunction Instance = new();

    private FrequencyFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "FREQUENCY";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 2)
        {
            return CellValue.Error("#VALUE!");
        }

        // Check for errors in arguments
        if (args[0].IsError)
        {
            return args[0];
        }

        if (args[1].IsError)
        {
            return args[1];
        }

        // Extract data values from first argument
        var dataValues = new List<double>();
        if (args[0].Type == CellValueType.Number)
        {
            dataValues.Add(args[0].NumericValue);
        }

        // Extract bin values from second argument
        var binValues = new List<double>();
        if (args[1].Type == CellValueType.Number)
        {
            binValues.Add(args[1].NumericValue);
        }

        // Sort bin values
        binValues.Sort();

        // Phase 0: Return simple count of data values
        // Full implementation would return an array of frequencies for each bin
        // For now, we return the total count of data values
        if (dataValues.Count == 0)
        {
            return CellValue.FromNumber(0);
        }

        // Calculate frequency distribution
        // Create bins: values <= binValues[0], values <= binValues[1], etc., and values > last bin
        var frequencies = new int[binValues.Count + 1];

        foreach (var value in dataValues)
        {
            var placed = false;
            for (int i = 0; i < binValues.Count; i++)
            {
                if (value <= binValues[i])
                {
                    frequencies[i]++;
                    placed = true;
                    break;
                }
            }

            // If not placed in any bin, put in the overflow bin (values > last bin)
            if (!placed)
            {
                frequencies[binValues.Count]++;
            }
        }

        // Phase 0: Return the first frequency (count of values <= first bin)
        // In a full implementation, this would return an array
        return CellValue.FromNumber(frequencies[0]);
    }
}
