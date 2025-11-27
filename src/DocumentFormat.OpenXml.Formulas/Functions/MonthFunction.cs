// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the MONTH function.
/// MONTH(date) - extracts month (1-12).
/// </summary>
public sealed class MonthFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly MonthFunction Instance = new();

    private MonthFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "MONTH";

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

        try
        {
            var date = DateTime.FromOADate(args[0].NumericValue);
            return FormulaResult.FromNumber(date.Month);
        }
        catch
        {
            return FormulaResult.Error("#NUM!");
        }
    }
}
