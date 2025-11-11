// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the MAX function.
/// MAX(number1, [number2], ...) - returns maximum value.
/// </summary>
public sealed class MaxFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly MaxFunction Instance = new();

    private MaxFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "MAX";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        var max = double.MinValue;
        var hasValue = false;

        foreach (var arg in args)
        {
            if (arg.IsError)
            {
                return arg; // Propagate errors
            }

            if (arg.Type == CellValueType.Number)
            {
                max = System.Math.Max(max, arg.NumericValue);
                hasValue = true;
            }
        }

        if (!hasValue)
        {
            return CellValue.FromNumber(0);
        }

        return CellValue.FromNumber(max);
    }
}
