// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the PERCENTOF function.
/// PERCENTOF(subset, total) - Returns subset as a percentage of total.
/// </summary>
public sealed class PercentOfFunction : IFunctionImplementation
{
    public static readonly PercentOfFunction Instance = new();

    private PercentOfFunction()
    {
    }

    public string Name => "PERCENTOF";

    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 2)
        {
            return CellValue.Error("#VALUE!");
        }

        foreach (var arg in args)
        {
            if (arg.IsError)
            {
                return arg;
            }
        }

        if (args[0].Type != CellValueType.Number || args[1].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }

        var subset = args[0].NumericValue;
        var total = args[1].NumericValue;

        if (total == 0)
        {
            return CellValue.Error("#DIV/0!");
        }

        return CellValue.FromNumber(subset / total);
    }
}
