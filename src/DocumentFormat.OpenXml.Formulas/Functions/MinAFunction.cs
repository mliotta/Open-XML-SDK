// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the MINA function.
/// MINA(value1, [value2], ...) - Minimum value including text and logical values.
/// Text evaluates as 0, TRUE as 1, FALSE as 0, empty values are ignored.
/// </summary>
public sealed class MinAFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly MinAFunction Instance = new();

    private MinAFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "MINA";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        var min = double.MaxValue;
        var hasValue = false;

        foreach (var arg in args)
        {
            if (arg.IsError)
            {
                return arg; // Propagate errors
            }

            if (arg.Type == CellValueType.Number)
            {
                min = System.Math.Min(min, arg.NumericValue);
                hasValue = true;
            }
            else if (arg.Type == CellValueType.Boolean)
            {
                min = System.Math.Min(min, arg.BoolValue ? 1.0 : 0.0);
                hasValue = true;
            }
            else if (arg.Type == CellValueType.Text)
            {
                // Text values count as 0
                min = System.Math.Min(min, 0.0);
                hasValue = true;
            }
            // Empty values are ignored
        }

        if (!hasValue)
        {
            return CellValue.FromNumber(0);
        }

        return CellValue.FromNumber(min);
    }
}
