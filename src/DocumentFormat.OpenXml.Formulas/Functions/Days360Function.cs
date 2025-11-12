// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the DAYS360 function.
/// DAYS360(start_date, end_date, [method]) - calculates days between dates using 360-day year.
/// </summary>
public sealed class Days360Function : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly Days360Function Instance = new();

    private Days360Function()
    {
    }

    /// <inheritdoc/>
    public string Name => "DAYS360";

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

        // Default to US/NASD method (FALSE)
        var useEuropeanMethod = false;

        if (args.Length == 3)
        {
            if (args[2].IsError)
            {
                return args[2];
            }

            if (args[2].Type == CellValueType.Boolean)
            {
                useEuropeanMethod = args[2].BoolValue;
            }
            else if (args[2].Type == CellValueType.Number)
            {
                useEuropeanMethod = args[2].NumericValue != 0;
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

            int startYear = startDate.Year;
            int startMonth = startDate.Month;
            int startDay = startDate.Day;
            int endYear = endDate.Year;
            int endMonth = endDate.Month;
            int endDay = endDate.Day;

            if (useEuropeanMethod)
            {
                // European method (30E/360)
                if (startDay == 31)
                {
                    startDay = 30;
                }

                if (endDay == 31)
                {
                    endDay = 30;
                }
            }
            else
            {
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
            }

            // Calculate days using 360-day year formula
            var days = ((endYear - startYear) * 360) + ((endMonth - startMonth) * 30) + (endDay - startDay);
            return CellValue.FromNumber(days);
        }
        catch
        {
            return CellValue.Error("#NUM!");
        }
    }

    private static bool IsLastDayOfFebruary(DateTime date)
    {
        return date.Month == 2 && date.Day == DateTime.DaysInMonth(date.Year, 2);
    }
}
