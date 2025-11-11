// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the DEGREES function.
/// DEGREES(radians) - converts radians to degrees.
/// </summary>
public sealed class DegreesFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly DegreesFunction Instance = new();

    private DegreesFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "DEGREES";

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

        var radians = args[0].NumericValue;
        var degrees = radians * 180.0 / System.Math.PI;

        return CellValue.FromNumber(degrees);
    }
}
