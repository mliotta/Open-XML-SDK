// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the SEQUENCE function.
/// SEQUENCE(rows, [columns], [start], [step]) - Generates a sequence of numbers.
/// rows: Number of rows to return
/// columns: Number of columns to return (default: 1)
/// start: First number in the sequence (default: 1)
/// step: Amount to increment each value (default: 1)
/// </summary>
public sealed class SequenceFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly SequenceFunction Instance = new SequenceFunction();

    private SequenceFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "SEQUENCE";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length == 0)
        {
            return CellValue.Error("#VALUE!");
        }

        // Check for errors in first argument
        if (args[0].IsError)
        {
            return args[0];
        }

        if (args[0].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }

        var rows = (int)args[0].NumericValue;
        if (rows <= 0)
        {
            return CellValue.Error("#VALUE!");
        }

        // Parse optional columns parameter
        var columns = 1;
        if (args.Length >= 2)
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
            if (columns <= 0)
            {
                return CellValue.Error("#VALUE!");
            }
        }

        // Parse optional start parameter
        var start = 1.0;
        if (args.Length >= 3)
        {
            if (args[2].IsError)
            {
                return args[2];
            }

            if (args[2].Type != CellValueType.Number)
            {
                return CellValue.Error("#VALUE!");
            }

            start = args[2].NumericValue;
        }

        // Parse optional step parameter
        var step = 1.0;
        if (args.Length >= 4)
        {
            if (args[3].IsError)
            {
                return args[3];
            }

            if (args[3].Type != CellValueType.Number)
            {
                return CellValue.Error("#VALUE!");
            }

            step = args[3].NumericValue;
        }

        // Check if we would exceed reasonable array size
        var totalCells = rows * columns;
        if (totalCells > 1000000) // 1 million cell limit
        {
            return CellValue.Error("#NUM!");
        }

        // Generate sequence
        var sequence = new CellValue[totalCells];
        var currentValue = start;

        for (var row = 0; row < rows; row++)
        {
            for (var col = 0; col < columns; col++)
            {
                sequence[row * columns + col] = CellValue.FromNumber(currentValue);
                currentValue += step;
            }
        }

        // Return first element (full array support would require engine changes)
        return sequence[0];
    }
}
