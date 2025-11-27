// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the TIME function.
/// TIME(hour, minute, second) - creates a time value from components.
/// </summary>
public sealed class TimeFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly TimeFunction Instance = new();

    private TimeFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "TIME";

    /// <inheritdoc/>
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        if (args.Length != 3)
        {
            return FormulaResult.Error("#VALUE!");
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

        if (args[0].Type != FormulaResultType.Number || args[1].Type != FormulaResultType.Number || args[2].Type != FormulaResultType.Number)
        {
            return FormulaResult.Error("#VALUE!");
        }

        try
        {
            var hour = (int)args[0].NumericValue;
            var minute = (int)args[1].NumericValue;
            var second = (int)args[2].NumericValue;

            // Handle negative values and overflow
            if (hour < 0 || minute < 0 || second < 0)
            {
                return FormulaResult.Error("#NUM!");
            }

            // Time is stored as fraction of day
            // 1 hour = 1/24 day, 1 minute = 1/1440 day, 1 second = 1/86400 day
            var timeValue = (hour / 24.0) + (minute / 1440.0) + (second / 86400.0);

            // Excel allows time values > 1 (represents multiple days)
            return FormulaResult.FromNumber(timeValue);
        }
        catch
        {
            return FormulaResult.Error("#NUM!");
        }
    }
}
