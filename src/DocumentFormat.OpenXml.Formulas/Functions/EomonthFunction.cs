// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the EOMONTH function.
/// EOMONTH(start_date, months) - returns the last day of the month, months in future/past.
/// </summary>
public sealed class EomonthFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly EomonthFunction Instance = new();

    private EomonthFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "EOMONTH";

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
            var startDate = DateTime.FromOADate(args[0].NumericValue);
            var monthsToAdd = (int)args[1].NumericValue;

            // Add months to the start date
            var targetDate = startDate.AddMonths(monthsToAdd);

            // Get the last day of that month
            var lastDay = DateTime.DaysInMonth(targetDate.Year, targetDate.Month);
            var endOfMonth = new DateTime(targetDate.Year, targetDate.Month, lastDay);

            var serialDate = endOfMonth.ToOADate();
            return CellValue.FromNumber(serialDate);
        }
        catch
        {
            return CellValue.Error("#NUM!");
        }
    }
}
