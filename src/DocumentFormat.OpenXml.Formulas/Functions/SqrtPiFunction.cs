// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the SQRTPI function.
/// SQRTPI(number) - returns the square root of (number * pi).
/// </summary>
public sealed class SqrtPiFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly SqrtPiFunction Instance = new();

    private SqrtPiFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "SQRTPI";

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

        var number = args[0].NumericValue;

        if (number < 0)
        {
            return FormulaResult.Error("#NUM!");
        }

        var result = System.Math.Sqrt(number * System.Math.PI);

        // Check for overflow
        if (double.IsInfinity(result) || double.IsNaN(result))
        {
            return FormulaResult.Error("#NUM!");
        }

        return FormulaResult.FromNumber(result);
    }
}
