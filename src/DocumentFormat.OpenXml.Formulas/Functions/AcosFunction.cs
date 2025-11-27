// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the ACOS function.
/// ACOS(number) - returns the arccosine (inverse cosine) of a number in radians.
/// Number must be between -1 and 1.
/// </summary>
public sealed class AcosFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly AcosFunction Instance = new();

    private AcosFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "ACOS";

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

        if (number < -1 || number > 1)
        {
            return FormulaResult.Error("#NUM!");
        }

        var result = System.Math.Acos(number);
        return FormulaResult.FromNumber(result);
    }
}
