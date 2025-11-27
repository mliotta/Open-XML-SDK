// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the YIELDDISC function.
/// YIELDDISC(settlement, maturity, pr, redemption, [basis]) - returns the annual yield of a discounted security.
/// </summary>
public sealed class YielddiscFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly YielddiscFunction Instance = new();

    private YielddiscFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "YIELDDISC";

    /// <inheritdoc/>
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        if (args.Length < 4 || args.Length > 5)
        {
            return FormulaResult.Error("#VALUE!");
        }

        // Check for errors in required arguments
        for (int i = 0; i < 4; i++)
        {
            if (args[i].IsError)
            {
                return args[i];
            }

            if (args[i].Type != FormulaResultType.Number)
            {
                return FormulaResult.Error("#VALUE!");
            }
        }

        var basis = 0;
        if (args.Length == 5)
        {
            if (args[4].IsError)
            {
                return args[4];
            }

            if (args[4].Type == FormulaResultType.Number)
            {
                basis = (int)args[4].NumericValue;
                if (!DayCountHelper.IsValidBasis(basis))
                {
                    return FormulaResult.Error("#NUM!");
                }
            }
            else
            {
                return FormulaResult.Error("#VALUE!");
            }
        }

        try
        {
            var settlement = DateTime.FromOADate(args[0].NumericValue);
            var maturity = DateTime.FromOADate(args[1].NumericValue);
            var pr = args[2].NumericValue;
            var redemption = args[3].NumericValue;

            // Validate inputs
            if (settlement >= maturity || pr <= 0 || redemption <= 0)
            {
                return FormulaResult.Error("#NUM!");
            }

            // Calculate fraction of year
            var yearFraction = DayCountHelper.DayCountFraction(settlement, maturity, basis);

            // YIELDDISC formula: ((redemption - pr) / pr) * (B / DSM)
            // where B is year basis, DSM is days from settlement to maturity
            var yieldDisc = ((redemption - pr) / pr) / yearFraction;

            if (double.IsNaN(yieldDisc) || double.IsInfinity(yieldDisc))
            {
                return FormulaResult.Error("#NUM!");
            }

            return FormulaResult.FromNumber(yieldDisc);
        }
        catch
        {
            return FormulaResult.Error("#NUM!");
        }
    }
}
