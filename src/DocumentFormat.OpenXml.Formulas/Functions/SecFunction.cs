// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the SEC function.
/// SEC(number) - returns the secant of an angle (number in radians).
/// SEC(x) = 1/COS(x)
/// </summary>
public sealed class SecFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly SecFunction Instance = new();

    private SecFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "SEC";

    /// <inheritdoc/>
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        if (args.Length != 1)
        {
            return FormulaResult.Error("#VALUE!");
        }

        if (args[0].IsError)
        {
            return args[0];
        }

        if (args[0].Type != FormulaResultType.Number)
        {
            return FormulaResult.Error("#VALUE!");
        }

        var cosValue = System.Math.Cos(args[0].NumericValue);

        // Check if cos is zero (would cause division by zero)
        // Use a larger threshold to catch values like PI/2 which are very close to zero
        if (System.Math.Abs(cosValue) < 1e-10)
        {
            return FormulaResult.Error("#DIV/0!");
        }

        var result = 1.0 / cosValue;

        if (double.IsInfinity(result) || double.IsNaN(result))
        {
            return FormulaResult.Error("#NUM!");
        }

        return FormulaResult.FromNumber(result);
    }
}
