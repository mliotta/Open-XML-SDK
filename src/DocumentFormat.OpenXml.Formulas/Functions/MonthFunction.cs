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
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 1)
        {
            return CellValue.Error("#VALUE!");
        }

        if (args[0].IsError)
        {
            return args[0];
        }

        if (args[0].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }

        try
        {
            var date = DateTime.FromOADate(args[0].NumericValue);
            return CellValue.FromNumber(date.Month);
        }
        catch
        {
            return CellValue.Error("#NUM!");
        }
    }
}
