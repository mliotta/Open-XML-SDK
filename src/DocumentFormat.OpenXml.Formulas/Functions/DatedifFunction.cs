// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the DATEDIF function.
/// DATEDIF(start_date, end_date, unit) - calculates the difference between two dates in various units.
/// </summary>
public sealed class DatedifFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly DatedifFunction Instance = new();

    private DatedifFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "DATEDIF";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 3)
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

        if (args[2].IsError)
        {
            return args[2];
        }

        if (args[0].Type != CellValueType.Number || args[1].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }

        if (args[2].Type != CellValueType.Text)
        {
            return CellValue.Error("#VALUE!");
        }

        try
        {
            var startDate = DateTime.FromOADate(args[0].NumericValue);
            var endDate = DateTime.FromOADate(args[1].NumericValue);

            // Validate start_date <= end_date
            if (startDate > endDate)
            {
                return CellValue.Error("#NUM!");
            }

            var unit = args[2].StringValue.ToUpperInvariant();

            var result = unit switch
            {
                "Y" => CalculateYears(startDate, endDate),
                "M" => CalculateMonths(startDate, endDate),
                "D" => CalculateDays(startDate, endDate),
                "YM" => CalculateMonthsExcludingYears(startDate, endDate),
                "YD" => CalculateDaysExcludingYears(startDate, endDate),
                "MD" => CalculateDaysExcludingMonthsAndYears(startDate, endDate),
                _ => throw new ArgumentException($"Invalid unit: {unit}"),
            };

            return CellValue.FromNumber(result);
        }
        catch
        {
            return CellValue.Error("#NUM!");
        }
    }

    /// <summary>
    /// Calculates complete years between two dates.
    /// </summary>
    private static int CalculateYears(DateTime startDate, DateTime endDate)
    {
        int years = endDate.Year - startDate.Year;

        // Check if we need to subtract a year (end date hasn't reached anniversary)
        if (endDate.Month < startDate.Month ||
            (endDate.Month == startDate.Month && endDate.Day < startDate.Day))
        {
            years--;
        }

        return years;
    }

    /// <summary>
    /// Calculates complete months between two dates.
    /// </summary>
    private static int CalculateMonths(DateTime startDate, DateTime endDate)
    {
        int months = ((endDate.Year - startDate.Year) * 12) + (endDate.Month - startDate.Month);

        // Check if we need to subtract a month (end date hasn't reached anniversary)
        if (endDate.Day < startDate.Day)
        {
            months--;
        }

        return months;
    }

    /// <summary>
    /// Calculates days between two dates.
    /// </summary>
    private static int CalculateDays(DateTime startDate, DateTime endDate)
    {
        return (int)(endDate - startDate).TotalDays;
    }

    /// <summary>
    /// Calculates months excluding years.
    /// </summary>
    private static int CalculateMonthsExcludingYears(DateTime startDate, DateTime endDate)
    {
        // Adjust start date to be in the same year as end date
        var adjustedStartDate = new DateTime(endDate.Year, startDate.Month, startDate.Day);
        if (adjustedStartDate > endDate)
        {
            adjustedStartDate = adjustedStartDate.AddYears(-1);
        }

        int months = endDate.Month - adjustedStartDate.Month;

        // Check if we need to subtract a month
        if (endDate.Day < adjustedStartDate.Day)
        {
            months--;
        }

        // Ensure result is non-negative
        if (months < 0)
        {
            months += 12;
        }

        return months;
    }

    /// <summary>
    /// Calculates days excluding years.
    /// </summary>
    private static int CalculateDaysExcludingYears(DateTime startDate, DateTime endDate)
    {
        // Move start date to the same year as end date
        var adjustedStartDate = new DateTime(endDate.Year, startDate.Month, startDate.Day);

        // If the adjusted date is after the end date, move it back one year
        if (adjustedStartDate > endDate)
        {
            adjustedStartDate = adjustedStartDate.AddYears(-1);
        }

        return (int)(endDate - adjustedStartDate).TotalDays;
    }

    /// <summary>
    /// Calculates days excluding months and years.
    /// </summary>
    private static int CalculateDaysExcludingMonthsAndYears(DateTime startDate, DateTime endDate)
    {
        // Move start date to the same month and year as end date
        var adjustedStartDate = new DateTime(endDate.Year, endDate.Month, startDate.Day);

        // If the adjusted date is after the end date, move it back one month
        if (adjustedStartDate > endDate)
        {
            adjustedStartDate = adjustedStartDate.AddMonths(-1);
        }

        return (int)(endDate - adjustedStartDate).TotalDays;
    }
}
