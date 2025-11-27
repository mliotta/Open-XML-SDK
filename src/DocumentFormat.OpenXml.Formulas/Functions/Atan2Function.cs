// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the ATAN2 function.
/// ATAN2(x_num, y_num) - returns the arctangent of the specified x and y coordinates in radians.
/// </summary>
public sealed class Atan2Function : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly Atan2Function Instance = new();

    private Atan2Function()
    {
    }

    /// <inheritdoc/>
    public string Name => "ATAN2";

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

        var xNum = args[0].NumericValue;
        var yNum = args[1].NumericValue;

        if (xNum == 0 && yNum == 0)
        {
            return FormulaResult.Error("#DIV/0!");
        }

        var result = System.Math.Atan2(yNum, xNum);
        return FormulaResult.FromNumber(result);
    }
}
