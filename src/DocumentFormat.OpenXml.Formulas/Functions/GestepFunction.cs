// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the GESTEP function.
/// GESTEP(number, [step]) - tests whether a number is greater than a threshold value.
/// Returns 1 if number >= step, 0 otherwise.
/// </summary>
public sealed class GestepFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly GestepFunction Instance = new();

    private GestepFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "GESTEP";

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

        var number = args[0].NumericValue;
        var step = 0.0;

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

            step = args[1].NumericValue;
        }

        // Return 1 if number >= step, 0 otherwise
        return CellValue.FromNumber(number >= step ? 1 : 0);
    }
}
