// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the RADIANS function.
/// RADIANS(degrees) - converts degrees to radians.
/// </summary>
public sealed class RadiansFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly RadiansFunction Instance = new();

    private RadiansFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "RADIANS";

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

        var degrees = args[0].NumericValue;
        var radians = degrees * System.Math.PI / 180.0;

        return CellValue.FromNumber(radians);
    }
}
