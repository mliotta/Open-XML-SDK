// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Globalization;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the ISOWEEKNUM function.
/// ISOWEEKNUM(date) - returns the ISO week number of the year for a given date.
/// ISO 8601 defines week 1 as the first week with Thursday in the new year.
/// </summary>
public sealed class IsoweeknumFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly IsoweeknumFunction Instance = new();

    private IsoweeknumFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "ISOWEEKNUM";

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

            // ISO 8601: Week starts on Monday, week 1 contains the first Thursday of the year
            var calendar = CultureInfo.InvariantCulture.Calendar;
            var weekNum = calendar.GetWeekOfYear(date, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);

            return CellValue.FromNumber(weekNum);
        }
        catch
        {
            return CellValue.Error("#NUM!");
        }
    }
}
