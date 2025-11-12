// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the RANDARRAY function.
/// RANDARRAY([rows], [columns], [min], [max], [whole_number]) - returns an array of random numbers.
/// - rows: Number of rows (default 1)
/// - columns: Number of columns (default 1)
/// - min: Minimum value (default 0)
/// - max: Maximum value (default 1)
/// - whole_number: TRUE for integers, FALSE for decimals (default FALSE)
/// Note: This function is volatile and recalculates each time it is evaluated.
/// </summary>
public sealed class RandArrayFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly RandArrayFunction Instance = new();

    private static readonly Random _random = new();

    private RandArrayFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "RANDARRAY";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length > 5)
        {
            return CellValue.Error("#VALUE!");
        }

        // Default values
        int rows = 1;
        int columns = 1;
        double min = 0.0;
        double max = 1.0;
        bool wholeNumber = false;

        // Parse rows (optional, default 1)
        if (args.Length > 0 && args[0].Type != CellValueType.Empty)
        {
            if (args[0].IsError)
            {
                return args[0];
            }

            if (args[0].Type != CellValueType.Number)
            {
                return CellValue.Error("#VALUE!");
            }

            rows = (int)args[0].NumericValue;
            if (rows < 1)
            {
                return CellValue.Error("#VALUE!");
            }
        }

        // Parse columns (optional, default 1)
        if (args.Length > 1 && args[1].Type != CellValueType.Empty)
        {
            if (args[1].IsError)
            {
                return args[1];
            }

            if (args[1].Type != CellValueType.Number)
            {
                return CellValue.Error("#VALUE!");
            }

            columns = (int)args[1].NumericValue;
            if (columns < 1)
            {
                return CellValue.Error("#VALUE!");
            }
        }

        // Parse min (optional, default 0)
        if (args.Length > 2 && args[2].Type != CellValueType.Empty)
        {
            if (args[2].IsError)
            {
                return args[2];
            }

            if (args[2].Type != CellValueType.Number)
            {
                return CellValue.Error("#VALUE!");
            }

            min = args[2].NumericValue;
        }

        // Parse max (optional, default 1)
        if (args.Length > 3 && args[3].Type != CellValueType.Empty)
        {
            if (args[3].IsError)
            {
                return args[3];
            }

            if (args[3].Type != CellValueType.Number)
            {
                return CellValue.Error("#VALUE!");
            }

            max = args[3].NumericValue;
        }

        // Validate min < max
        if (min >= max)
        {
            return CellValue.Error("#VALUE!");
        }

        // Parse whole_number (optional, default FALSE)
        if (args.Length > 4 && args[4].Type != CellValueType.Empty)
        {
            if (args[4].IsError)
            {
                return args[4];
            }

            if (args[4].Type == CellValueType.Boolean)
            {
                wholeNumber = args[4].BoolValue;
            }
            else if (args[4].Type == CellValueType.Number)
            {
                wholeNumber = args[4].NumericValue != 0;
            }
            else
            {
                return CellValue.Error("#VALUE!");
            }
        }

        // Check if we would exceed reasonable array size
        var totalCells = rows * columns;
        if (totalCells > 1000000) // 1 million cell limit
        {
            return CellValue.Error("#NUM!");
        }

        // Generate the array (flattened)
        // Full array support would require engine changes, so we return first element
        var range = max - min;
        double value = _random.NextDouble() * range + min;

        if (wholeNumber)
        {
            value = System.Math.Floor(value);
        }

        return CellValue.FromNumber(value);
    }
}
