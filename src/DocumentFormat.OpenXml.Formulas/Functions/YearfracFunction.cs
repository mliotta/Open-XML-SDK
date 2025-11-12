// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the YEARFRAC function.
/// YEARFRAC(start_date, end_date, [basis]) - calculates the fraction of year between two dates.
/// </summary>
public sealed class YearfracFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly YearfracFunction Instance = new();

    private YearfracFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "YEARFRAC";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 2 || args.Length > 3)
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

        // Default to US (NASD) 30/360 method (basis = 0)
        var basis = 0;

        if (args.Length == 3)
        {
            if (args[2].IsError)
            {
                return args[2];
            }

            if (args[2].Type == CellValueType.Number)
            {
                basis = (int)args[2].NumericValue;
                if (basis < 0 || basis > 4)
                {
                    return CellValue.Error("#NUM!");
                }
            }
            else
            {
                return CellValue.Error("#VALUE!");
            }
        }

        try
        {
            var startDate = DateTime.FromOADate(args[0].NumericValue);
            var endDate = DateTime.FromOADate(args[1].NumericValue);

            // Validate start_date < end_date
            if (startDate > endDate)
            {
                return CellValue.Error("#NUM!");
            }

            double fraction = basis switch
            {
                0 => CalculateBasis0(startDate, endDate), // US (NASD) 30/360
                1 => CalculateBasis1(startDate, endDate), // Actual/actual
                2 => CalculateBasis2(startDate, endDate), // Actual/360
                3 => CalculateBasis3(startDate, endDate), // Actual/365
                4 => CalculateBasis4(startDate, endDate), // European 30/360
                _ => 0.0,
            };

            return CellValue.FromNumber(fraction);
        }
        catch
        {
            return CellValue.Error("#NUM!");
        }
    }

    /// <summary>
    /// Basis 0: US (NASD) 30/360.
    /// </summary>
    private static double CalculateBasis0(DateTime startDate, DateTime endDate)
    {
        int startYear = startDate.Year;
        int startMonth = startDate.Month;
        int startDay = startDate.Day;
        int endYear = endDate.Year;
        int endMonth = endDate.Month;
        int endDay = endDate.Day;

        // US/NASD method (30US/360)
        if (startDay == 31)
        {
            startDay = 30;
        }

        if (endDay == 31 && startDay >= 30)
        {
            endDay = 30;
        }

        // Handle February special case
        if (startMonth == 2 && IsLastDayOfFebruary(startDate))
        {
            startDay = 30;
        }

        if (endMonth == 2 && IsLastDayOfFebruary(endDate))
        {
            endDay = 30;
        }

        var days = ((endYear - startYear) * 360) + ((endMonth - startMonth) * 30) + (endDay - startDay);
        return days / 360.0;
    }

    /// <summary>
    /// Basis 1: Actual/actual.
    /// </summary>
    private static double CalculateBasis1(DateTime startDate, DateTime endDate)
    {
        // Count actual days in the period
        var totalDays = (endDate - startDate).TotalDays;

        // Determine the year basis (weighted average of days in each year)
        var startYear = startDate.Year;
        var endYear = endDate.Year;

        if (startYear == endYear)
        {
            // Same year - use days in that year
            var daysInYear = DateTime.IsLeapYear(startYear) ? 366.0 : 365.0;
            return totalDays / daysInYear;
        }
        else
        {
            // Multiple years - use weighted average approach
            double yearFraction = 0.0;
            var currentDate = startDate;

            while (currentDate.Year <= endYear)
            {
                var yearStart = new DateTime(currentDate.Year, 1, 1);
                var yearEnd = new DateTime(currentDate.Year, 12, 31);

                var periodStart = currentDate > yearStart ? currentDate : yearStart;
                var periodEnd = endDate < yearEnd ? endDate : yearEnd;

                if (periodStart <= periodEnd)
                {
                    var daysInThisYear = (periodEnd - periodStart).TotalDays;
                    var totalDaysInYear = DateTime.IsLeapYear(currentDate.Year) ? 366.0 : 365.0;
                    yearFraction += daysInThisYear / totalDaysInYear;
                }

                currentDate = new DateTime(currentDate.Year + 1, 1, 1);
            }

            return yearFraction;
        }
    }

    /// <summary>
    /// Basis 2: Actual/360.
    /// </summary>
    private static double CalculateBasis2(DateTime startDate, DateTime endDate)
    {
        var totalDays = (endDate - startDate).TotalDays;
        return totalDays / 360.0;
    }

    /// <summary>
    /// Basis 3: Actual/365.
    /// </summary>
    private static double CalculateBasis3(DateTime startDate, DateTime endDate)
    {
        var totalDays = (endDate - startDate).TotalDays;
        return totalDays / 365.0;
    }

    /// <summary>
    /// Basis 4: European 30/360.
    /// </summary>
    private static double CalculateBasis4(DateTime startDate, DateTime endDate)
    {
        int startYear = startDate.Year;
        int startMonth = startDate.Month;
        int startDay = startDate.Day;
        int endYear = endDate.Year;
        int endMonth = endDate.Month;
        int endDay = endDate.Day;

        // European method (30E/360)
        if (startDay == 31)
        {
            startDay = 30;
        }

        if (endDay == 31)
        {
            endDay = 30;
        }

        var days = ((endYear - startYear) * 360) + ((endMonth - startMonth) * 30) + (endDay - startDay);
        return days / 360.0;
    }

    private static bool IsLastDayOfFebruary(DateTime date)
    {
        return date.Month == 2 && date.Day == DateTime.DaysInMonth(date.Year, 2);
    }
}
