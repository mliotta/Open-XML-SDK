// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the CSCH function.
/// CSCH(number) - returns the hyperbolic cosecant of a number.
/// CSCH(x) = 1/SINH(x) = 2/(e^x - e^(-x))
/// </summary>
public sealed class CschFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly CschFunction Instance = new();

    private CschFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "CSCH";

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

        var sinhValue = System.Math.Sinh(args[0].NumericValue);

        // Check if sinh is zero (would cause division by zero)
        if (System.Math.Abs(sinhValue) < double.Epsilon)
        {
            return FormulaResult.Error("#DIV/0!");
        }

        var result = 1.0 / sinhValue;

        if (double.IsInfinity(result) || double.IsNaN(result))
        {
            return FormulaResult.Error("#NUM!");
        }

        return FormulaResult.FromNumber(result);
    }
}
