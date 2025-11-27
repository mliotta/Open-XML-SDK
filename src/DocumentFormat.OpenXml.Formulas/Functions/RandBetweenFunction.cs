// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the RANDBETWEEN function.
/// RANDBETWEEN(bottom, top) - returns a random integer between bottom and top.
/// Note: This function is volatile and recalculates each time it is evaluated.
/// </summary>
public sealed class RandBetweenFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly RandBetweenFunction Instance = new();

    private static readonly Random _random = new();

    private RandBetweenFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "RANDBETWEEN";

    /// <inheritdoc/>
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        if (args.Length != 2)
        {
            return FormulaResult.Error("#VALUE!");
        }

        if (args[0].IsError)
        {
            return args[0];
        }

        if (args[1].IsError)
        {
            return args[1];
        }

        if (args[0].Type != FormulaResultType.Number || args[1].Type != FormulaResultType.Number)
        {
            return FormulaResult.Error("#VALUE!");
        }

        var bottom = (int)System.Math.Floor(args[0].NumericValue);
        var top = (int)System.Math.Floor(args[1].NumericValue);

        if (bottom > top)
        {
            return FormulaResult.Error("#NUM!");
        }

        // Random.Next is exclusive on upper bound, so add 1
        var result = _random.Next(bottom, top + 1);
        return FormulaResult.FromNumber(result);
    }
}
