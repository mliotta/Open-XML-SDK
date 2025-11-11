// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the WEEKDAY function.
/// WEEKDAY(date, [return_type]) - returns day of week.
/// </summary>
public sealed class WeekdayFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly WeekdayFunction Instance = new();

    private WeekdayFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "WEEKDAY";

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

        var returnType = 1; // Default: Sunday=1, Saturday=7

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
        }

        try
        {
            var date = DateTime.FromOADate(args[0].NumericValue);
            var dayOfWeek = (int)date.DayOfWeek; // 0=Sunday, 6=Saturday

            // Convert based on return_type
            int result = returnType switch
            {
                1 => dayOfWeek + 1, // Sunday=1, Saturday=7
                2 => dayOfWeek == 0 ? 7 : dayOfWeek, // Monday=1, Sunday=7
                3 => dayOfWeek == 0 ? 6 : dayOfWeek - 1, // Monday=0, Sunday=6
                _ => dayOfWeek + 1, // Default to type 1
            };

            return CellValue.FromNumber(result);
        }
        catch
        {
            return CellValue.Error("#NUM!");
        }
    }
}
