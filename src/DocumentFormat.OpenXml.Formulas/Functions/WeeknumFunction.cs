// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Globalization;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the WEEKNUM function.
/// WEEKNUM(serial_number, [return_type]) - returns the week number of a date.
/// Return_type determines the week start day: 1 (default, Sunday), 2 (Monday), etc.
/// </summary>
public sealed class WeeknumFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly WeeknumFunction Instance = new();

    private WeeknumFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "WEEKNUM";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 1 || args.Length > 2)
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

        var returnType = 1; // Default: week starts on Sunday

        if (args.Length == 2)
        {
            if (args[1].IsError)
            {
                return args[1];
            }

            if (args[1].Type != CellValueType.Number)
            {
                return CellValue.Error("#VALUE!");
            }

            returnType = (int)args[1].NumericValue;

            // Validate return_type
            if (returnType < 1 || returnType > 21)
            {
                return CellValue.Error("#NUM!");
            }
        }

        try
        {
            var date = DateTime.FromOADate(args[0].NumericValue);

            // Calculate week number based on return_type
            int weekNum;

            // Excel WEEKNUM function compatibility
            if (returnType == 1 || returnType == 17)
            {
                // Week starts on Sunday (type 1) or Saturday (type 17)
                // Use simple calculation: week 1 contains Jan 1
                var jan1 = new DateTime(date.Year, 1, 1);
                var dayOffset = returnType == 1 ? 0 : 1; // Sunday=0, Saturday=1

                // Adjust for the first week
                var firstWeekDay = (int)jan1.DayOfWeek;
                var daysToAdd = (7 - firstWeekDay + dayOffset) % 7;

                if (daysToAdd > 0)
                {
                    jan1 = jan1.AddDays(-firstWeekDay + dayOffset);
                }

                var daysSinceJan1 = (date - jan1).Days;
                weekNum = (daysSinceJan1 / 7) + 1;
            }
            else if (returnType == 2 || returnType == 11)
            {
                // Week starts on Monday (type 2) or Monday (type 11 - ISO week)
                // For type 11, use ISO 8601 standard
                if (returnType == 11)
                {
                    // ISO 8601: Week 1 is the first week with Thursday
                    var calendar = CultureInfo.InvariantCulture.Calendar;
                    weekNum = calendar.GetWeekOfYear(date, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
                }
                else
                {
                    // Type 2: Simple Monday-based week counting
                    var jan1 = new DateTime(date.Year, 1, 1);
                    var firstMonday = jan1;

                    // Find first Monday of the year
                    while (firstMonday.DayOfWeek != DayOfWeek.Monday)
                    {
                        firstMonday = firstMonday.AddDays(1);
                    }

                    if (date < firstMonday)
                    {
                        weekNum = 1;
                    }
                    else
                    {
                        var daysSinceFirstMonday = (date - firstMonday).Days;
                        weekNum = (daysSinceFirstMonday / 7) + 2;
                    }
                }
            }
            else if (returnType == 21)
            {
                // ISO 8601 week number (same as type 11)
                var calendar = CultureInfo.InvariantCulture.Calendar;
                weekNum = calendar.GetWeekOfYear(date, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
            }
            else
            {
                // For other types (12-17), map to appropriate day
                // 12=Tuesday, 13=Wednesday, 14=Thursday, 15=Friday, 16=Saturday, 17=Sunday
                var jan1 = new DateTime(date.Year, 1, 1);
                var targetDayOfWeek = (DayOfWeek)((returnType - 12 + 2) % 7);

                var firstTargetDay = jan1;
                while (firstTargetDay.DayOfWeek != targetDayOfWeek)
                {
                    firstTargetDay = firstTargetDay.AddDays(1);
                }

                if (date < firstTargetDay)
                {
                    weekNum = 1;
                }
                else
                {
                    var daysSinceFirstTarget = (date - firstTargetDay).Days;
                    weekNum = (daysSinceFirstTarget / 7) + 2;
                }
            }

            return CellValue.FromNumber(weekNum);
        }
        catch
        {
            return CellValue.Error("#NUM!");
        }
    }
}
