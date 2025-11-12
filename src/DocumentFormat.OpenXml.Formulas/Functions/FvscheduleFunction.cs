// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the FVSCHEDULE function.
/// FVSCHEDULE(principal, schedule) - calculates the future value of an initial principal after applying a series of compound interest rates.
/// </summary>
public sealed class FvscheduleFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly FvscheduleFunction Instance = new();

    private FvscheduleFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "FVSCHEDULE";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 2)
        {
            return CellValue.Error("#VALUE!");
        }

        // Check for errors in principal argument
        if (args[0].IsError)
        {
            return args[0];
        }

        // Validate principal argument is a number
        if (args[0].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }

        var principal = args[0].NumericValue;

        // Extract schedule values (all remaining arguments)
        var scheduleCount = args.Length - 1;
        var schedule = new double[scheduleCount];

        for (int i = 0; i < scheduleCount; i++)
        {
            if (args[i + 1].IsError)
            {
                return args[i + 1];
            }

            if (args[i + 1].Type != CellValueType.Number)
            {
                return CellValue.Error("#VALUE!");
            }

            schedule[i] = args[i + 1].NumericValue;
        }

        // Calculate future value by compounding with each rate in the schedule
        double fv = principal;

        foreach (var rate in schedule)
        {
            fv *= (1 + rate);

            if (double.IsNaN(fv) || double.IsInfinity(fv))
            {
                return CellValue.Error("#NUM!");
            }
        }

        return CellValue.FromNumber(fv);
    }
}
