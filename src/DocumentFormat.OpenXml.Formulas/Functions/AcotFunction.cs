// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the ACOT function.
/// ACOT(number) - returns the inverse cotangent (arccotangent) of a number.
/// ACOT(x) = PI/2 - ATAN(x).
/// Returns value in radians between 0 and PI.
/// </summary>
public sealed class AcotFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly AcotFunction Instance = new();

    private AcotFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "ACOT";

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

        var x = args[0].NumericValue;

        // ACOT(x) = PI/2 - ATAN(x)
        // This is the standard mathematical definition
        var result = System.Math.PI / 2 - System.Math.Atan(x);

        if (double.IsInfinity(result) || double.IsNaN(result))
        {
            return FormulaResult.Error("#NUM!");
        }

        return FormulaResult.FromNumber(result);
    }
}
