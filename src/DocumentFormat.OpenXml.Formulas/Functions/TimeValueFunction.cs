// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Globalization;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the TIMEVALUE function.
/// TIMEVALUE(time_text) - converts text to time (fraction of day).
/// </summary>
public sealed class TimeValueFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly TimeValueFunction Instance = new();

    private TimeValueFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "TIMEVALUE";

    /// <inheritdoc/>
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        if (args.Length != 1)
        {
            return FormulaResult.Error("#VALUE!");
        }

        if (args[0].IsError)
        {
            return args[0];
        }

        if (args[0].Type != FormulaResultType.Text)
        {
            return FormulaResult.Error("#VALUE!");
        }

        try
        {
            var timeText = args[0].StringValue;

            // Try to parse as DateTime (handles various time formats)
            if (DateTime.TryParse(timeText, CultureInfo.InvariantCulture, DateTimeStyles.None, out var dateTime))
            {
                // Return only the time portion (fractional part of OADate)
                var oaDate = dateTime.ToOADate();
                var timeValue = oaDate - System.Math.Floor(oaDate);
                return FormulaResult.FromNumber(timeValue);
            }

            // If parsing fails, return error
            return FormulaResult.Error("#VALUE!");
        }
        catch
        {
            return FormulaResult.Error("#VALUE!");
        }
    }
}
