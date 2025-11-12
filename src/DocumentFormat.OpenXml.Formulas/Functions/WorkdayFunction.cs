// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the WORKDAY function.
/// WORKDAY(start_date, days, [holidays]) - returns a date that is the specified number of working days from the start date.
/// Working days exclude weekends (Saturday and Sunday) and optionally specified holidays.
/// </summary>
public sealed class WorkdayFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly WorkdayFunction Instance = new();

    private WorkdayFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "WORKDAY";

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

        // Parse optional holidays (currently only supports single holiday value)
        var holidays = new HashSet<int>();
        if (args.Length == 3)
        {
            if (args[2].IsError)
            {
                return args[2];
            }

            if (args[2].Type == CellValueType.Number)
            {
                // Single holiday date
                holidays.Add((int)System.Math.Floor(args[2].NumericValue));
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
                var dayOfWeek = currentDate.DayOfWeek;
                var serialDate = (int)System.Math.Floor(currentDate.ToOADate());

                // Count if not a weekend and not a holiday
                if (dayOfWeek != DayOfWeek.Saturday &&
                    dayOfWeek != DayOfWeek.Sunday &&
                    !holidays.Contains(serialDate))
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
}
