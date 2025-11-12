// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the ODD function.
/// ODD(number) - rounds a number up to the nearest odd integer.
/// </summary>
public sealed class OddFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly OddFunction Instance = new();

    private OddFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "ODD";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 1)
        {
            return CellValue.Error("#VALUE!");
        }

        if (args[0].IsError)
        {
            return args[0];
        }

        if (args[0].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }

        var number = args[0].NumericValue;

        // For positive numbers, round up to next odd
        // For negative numbers, round down (away from zero) to next odd
        double result;
        if (number >= 0)
        {
            result = System.Math.Ceiling(number);
            if (result % 2 == 0)
            {
                result += 1;
            }
        }
        else
        {
            result = System.Math.Floor(number);
            if (result % 2 == 0)
            {
                result -= 1;
            }
        }

        return CellValue.FromNumber(result);
    }
}
