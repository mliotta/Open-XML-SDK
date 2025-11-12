// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the DAYS function.
/// DAYS(end_date, start_date) - returns the number of days between two dates.
/// </summary>
public sealed class DaysFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly DaysFunction Instance = new();

    private DaysFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "DAYS";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 2)
        {
            return CellValue.Error("#VALUE!");
        }

        if (args[0].IsError)
        {
            return args[0];
        }

        if (args[1].IsError)
        {
            return args[1];
        }

        if (args[0].Type != CellValueType.Number || args[1].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }

        try
        {
            // DAYS(end_date, start_date) = end_date - start_date
            var endDate = args[0].NumericValue;
            var startDate = args[1].NumericValue;
            var days = endDate - startDate;
            return CellValue.FromNumber(days);
        }
        catch
        {
            return CellValue.Error("#NUM!");
        }
    }
}
