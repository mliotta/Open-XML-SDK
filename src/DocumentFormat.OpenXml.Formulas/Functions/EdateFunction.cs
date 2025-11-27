// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the EDATE function.
/// EDATE(start_date, months) - returns a date that is months in future/past.
/// </summary>
public sealed class EdateFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly EdateFunction Instance = new();

    private EdateFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "EDATE";

    /// <inheritdoc/>
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        if (args.Length != 2)
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

        if (args[0].Type != FormulaResultType.Number || args[1].Type != FormulaResultType.Number)
        {
            return FormulaResult.Error("#VALUE!");
        }

        try
        {
            var startDate = DateTime.FromOADate(args[0].NumericValue);
            var monthsToAdd = (int)args[1].NumericValue;

            // Add months to the start date
            var resultDate = startDate.AddMonths(monthsToAdd);

            var serialDate = resultDate.ToOADate();
            return FormulaResult.FromNumber(serialDate);
        }
        catch
        {
            return FormulaResult.Error("#NUM!");
        }
    }
}
