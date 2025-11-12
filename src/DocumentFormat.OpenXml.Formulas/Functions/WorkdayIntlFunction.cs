// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the WORKDAY.INTL function.
/// WORKDAY.INTL(start_date, days, [weekend], [holidays]) - returns a date that is the specified number
/// of working days from the start date, with customizable weekend days.
/// Weekend parameter can be a number (1-17) or a 7-character string of 0s and 1s.
/// </summary>
public sealed class WorkdayIntlFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly WorkdayIntlFunction Instance = new();

    private WorkdayIntlFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "WORKDAY.INTL";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 2 || args.Length > 4)
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

        // Parse weekend parameter (default is 1 = Saturday/Sunday)
        var weekendMask = new bool[7]; // Sunday through Saturday
        var weekendType = 1;

        if (args.Length >= 3 && !args[2].IsError)
        {
            if (args[2].Type == CellValueType.Number)
            {
                weekendType = (int)args[2].NumericValue;
                if (weekendType < 1 || weekendType > 17)
                {
                    return CellValue.Error("#NUM!");
                }

                weekendMask = GetWeekendMaskFromType(weekendType);
            }
            else if (args[2].Type == CellValueType.Text)
            {
                // Custom weekend string: 7 characters, 0=workday, 1=weekend
                var weekendString = args[2].StringValue;
                if (weekendString.Length != 7)
                {
                    return CellValue.Error("#VALUE!");
                }

                for (int i = 0; i < 7; i++)
                {
                    if (weekendString[i] == '1')
                    {
                        weekendMask[i] = true;
                    }
                    else if (weekendString[i] != '0')
                    {
                        return CellValue.Error("#VALUE!");
                    }
                }
            }
            else
            {
                return CellValue.Error("#VALUE!");
            }
        }
        else
        {
            // Default: Saturday and Sunday
            weekendMask = GetWeekendMaskFromType(1);
        }

        // Parse optional holidays (currently only supports single holiday value)
        var holidays = new HashSet<int>();
        if (args.Length == 4)
        {
            if (args[3].IsError)
            {
                return args[3];
            }

            if (args[3].Type == CellValueType.Number)
            {
                // Single holiday date
                holidays.Add((int)System.Math.Floor(args[3].NumericValue));
            }
        }

        try
        {
            var startDate = DateTime.FromOADate(args[0].NumericValue);
            var daysToAdd = (int)args[1].NumericValue;

            // Determine direction (forward or backward)
            int direction = daysToAdd >= 0 ? 1 : -1;
            int remainingDays = System.Math.Abs(daysToAdd);

            var currentDate = startDate;

            // Add/subtract working days
            while (remainingDays > 0)
            {
                currentDate = currentDate.AddDays(direction);
                var dayOfWeek = (int)currentDate.DayOfWeek; // Sunday=0, Monday=1, ..., Saturday=6
                var serialDate = (int)System.Math.Floor(currentDate.ToOADate());

                // Count if not a weekend day and not a holiday
                if (!weekendMask[dayOfWeek] && !holidays.Contains(serialDate))
                {
                    remainingDays--;
                }
            }

            return CellValue.FromNumber(currentDate.ToOADate());
        }
        catch
        {
            return CellValue.Error("#NUM!");
        }
    }

    private static bool[] GetWeekendMaskFromType(int weekendType)
    {
        // Returns a 7-element array where true indicates a weekend day
        // Index: 0=Sunday, 1=Monday, 2=Tuesday, 3=Wednesday, 4=Thursday, 5=Friday, 6=Saturday
        var mask = new bool[7];

        switch (weekendType)
        {
            case 1: // Saturday, Sunday
                mask[0] = true; // Sunday
                mask[6] = true; // Saturday
                break;
            case 2: // Sunday, Monday
                mask[0] = true;
                mask[1] = true;
                break;
            case 3: // Monday, Tuesday
                mask[1] = true;
                mask[2] = true;
                break;
            case 4: // Tuesday, Wednesday
                mask[2] = true;
                mask[3] = true;
                break;
            case 5: // Wednesday, Thursday
                mask[3] = true;
                mask[4] = true;
                break;
            case 6: // Thursday, Friday
                mask[4] = true;
                mask[5] = true;
                break;
            case 7: // Friday, Saturday
                mask[5] = true;
                mask[6] = true;
                break;
            case 11: // Sunday only
                mask[0] = true;
                break;
            case 12: // Monday only
                mask[1] = true;
                break;
            case 13: // Tuesday only
                mask[2] = true;
                break;
            case 14: // Wednesday only
                mask[3] = true;
                break;
            case 15: // Thursday only
                mask[4] = true;
                break;
            case 16: // Friday only
                mask[5] = true;
                break;
            case 17: // Saturday only
                mask[6] = true;
                break;
        }

        return mask;
    }
}
