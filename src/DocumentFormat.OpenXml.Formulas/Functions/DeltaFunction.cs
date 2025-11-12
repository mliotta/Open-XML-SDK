// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the DELTA function.
/// DELTA(number1, [number2]) - tests whether two values are equal. Returns 1 if equal, 0 otherwise.
/// </summary>
public sealed class DeltaFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly DeltaFunction Instance = new();

    private DeltaFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "DELTA";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 1 || args.Length > 2)
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

        var number1 = args[0].NumericValue;
        var number2 = 0.0;

        if (args.Length == 2)
        {
            if (args[1].IsError)
            {
                return args[1];
            }

            if (args[1].Type != CellValueType.Number)
            {
                return CellValue.Error("#VALUE!");
            }

            number2 = args[1].NumericValue;
        }

        // Return 1 if equal, 0 otherwise
        return CellValue.FromNumber(System.Math.Abs(number1 - number2) < 1e-10 ? 1 : 0);
    }
}
