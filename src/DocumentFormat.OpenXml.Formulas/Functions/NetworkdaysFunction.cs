// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the NETWORKDAYS function.
/// NETWORKDAYS(start_date, end_date, [holidays]) - returns the number of working days between two dates.
/// Working days exclude weekends (Saturday and Sunday) and optionally specified holidays.
/// </summary>
public sealed class NetworkdaysFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly NetworkdaysFunction Instance = new();

    private NetworkdaysFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "NETWORKDAYS";

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
            var endDate = DateTime.FromOADate(args[1].NumericValue);

            // Ensure start is before end
            if (startDate > endDate)
            {
                var temp = startDate;
                startDate = endDate;
                endDate = temp;
            }

            int workingDays = 0;
            var currentDate = startDate;

            while (currentDate <= endDate)
            {
                var dayOfWeek = currentDate.DayOfWeek;
                var serialDate = (int)System.Math.Floor(currentDate.ToOADate());

                // Count if not a weekend and not a holiday
                if (dayOfWeek != DayOfWeek.Saturday &&
                    dayOfWeek != DayOfWeek.Sunday &&
                    !holidays.Contains(serialDate))
                {
                    workingDays++;
                }

                currentDate = currentDate.AddDays(1);
            }

            // If original order was reversed, return negative count
            if (DateTime.FromOADate(args[0].NumericValue) > DateTime.FromOADate(args[1].NumericValue))
            {
                workingDays = -workingDays;
            }

            return CellValue.FromNumber(workingDays);
        }
        catch
        {
            return CellValue.Error("#NUM!");
        }
    }
}
