// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the EXP function.
/// EXP(number) - returns e raised to the power of number.
/// </summary>
public sealed class ExpFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly ExpFunction Instance = new();

    private ExpFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "EXP";

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

        var result = System.Math.Exp(args[0].NumericValue);

        if (double.IsInfinity(result) || double.IsNaN(result))
        {
            return FormulaResult.Error("#NUM!");
        }

        return FormulaResult.FromNumber(result);
    }
}
