// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the LOG function.
/// LOG(number, [base]) - returns the logarithm of a number to the specified base (default 10).
/// </summary>
public sealed class LogFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly LogFunction Instance = new();

    private LogFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "LOG";

    /// <inheritdoc/>
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        if (args.Length < 1 || args.Length > 2)
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

        if (number <= 0)
        {
            return FormulaResult.Error("#NUM!");
        }

        double baseValue = 10.0;

        if (args.Length == 2)
        {
            if (args[1].IsError)
            {
                return args[1];
            }

            if (args[1].Type != FormulaResultType.Number)
            {
                return FormulaResult.Error("#VALUE!");
            }

            baseValue = args[1].NumericValue;

            if (baseValue <= 0 || baseValue == 1)
            {
                return FormulaResult.Error("#NUM!");
            }
        }

        var result = System.Math.Log(number, baseValue);

        if (double.IsNaN(result) || double.IsInfinity(result))
        {
            return FormulaResult.Error("#NUM!");
        }

        return FormulaResult.FromNumber(result);
    }
}
