// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the YEAR function.
/// YEAR(date) - extracts year from date.
/// </summary>
public sealed class YearFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly YearFunction Instance = new();

    private YearFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "YEAR";

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
            return CellValue.FromNumber(date.Year);
        }
        catch
        {
            return CellValue.Error("#NUM!");
        }
    }
}
