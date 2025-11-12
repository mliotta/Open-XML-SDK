// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the ACCRINTM function.
/// ACCRINTM(issue, settlement, rate, par, [basis]) - returns the accrued interest for a security that pays interest at maturity.
/// </summary>
public sealed class AccrintmFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly AccrintmFunction Instance = new();

    private AccrintmFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "ACCRINTM";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 4 || args.Length > 5)
        {
            return CellValue.Error("#VALUE!");
        }

        // Check for errors in required arguments
        for (int i = 0; i < 4; i++)
        {
            if (args[i].IsError)
            {
                return args[i];
            }

            if (args[i].Type != CellValueType.Number)
            {
                return CellValue.Error("#VALUE!");
            }
        }

        var basis = 0;
        if (args.Length == 5 && args[4].Type != CellValueType.Empty)
        {
            if (args[4].IsError)
            {
                return args[4];
            }

            if (args[4].Type == CellValueType.Number)
            {
                basis = (int)args[4].NumericValue;
                if (!DayCountHelper.IsValidBasis(basis))
                {
                    return CellValue.Error("#NUM!");
                }
            }
            else
            {
                return CellValue.Error("#VALUE!");
            }
        }

        try
        {
            var issue = DateTime.FromOADate(args[0].NumericValue);
            var settlement = DateTime.FromOADate(args[1].NumericValue);
            var rate = args[2].NumericValue;
            var par = args[3].NumericValue;

            // Validate inputs
            if (rate <= 0 || par <= 0)
            {
                return CellValue.Error("#NUM!");
            }

            if (issue >= settlement)
            {
                return CellValue.Error("#NUM!");
            }

            // Calculate accrued interest
            var dayCount = DayCountHelper.DayCountFraction(issue, settlement, basis);
            var accruedInterest = par * rate * dayCount;

            if (double.IsNaN(accruedInterest) || double.IsInfinity(accruedInterest))
            {
                return CellValue.Error("#NUM!");
            }

            return CellValue.FromNumber(accruedInterest);
        }
        catch
        {
            return CellValue.Error("#NUM!");
        }
    }
}
