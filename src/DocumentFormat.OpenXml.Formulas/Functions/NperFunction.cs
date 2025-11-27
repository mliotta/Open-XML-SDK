// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the NPER function.
/// NPER(rate, pmt, pv, [fv], [type]) - calculates the number of periods for an investment based on periodic, constant payments and a constant interest rate.
/// </summary>
public sealed class NperFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly NperFunction Instance = new();

    private NperFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "NPER";

    /// <inheritdoc/>
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        if (args.Length < 3 || args.Length > 5)
        {
            return FormulaResult.Error("#VALUE!");
        }

        // Check for errors in required arguments
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

        // Validate required arguments are numbers
        if (args[0].Type != FormulaResultType.Number || args[1].Type != FormulaResultType.Number || args[2].Type != FormulaResultType.Number)
        {
            return FormulaResult.Error("#VALUE!");
        }

        var rate = args[0].NumericValue;
        var pmt = args[1].NumericValue;
        var pv = args[2].NumericValue;
        var fv = 0.0;
        var type = 0.0;

        // Optional fv parameter
        if (args.Length >= 4)
        {
            if (args[3].IsError)
            {
                return args[3];
            }

            if (args[3].Type != FormulaResultType.Number)
            {
                return FormulaResult.Error("#VALUE!");
            }

            fv = args[3].NumericValue;
        }

        // Optional type parameter
        if (args.Length == 5)
        {
            if (args[4].IsError)
            {
                return args[4];
            }

            if (args[4].Type != FormulaResultType.Number)
            {
                return FormulaResult.Error("#VALUE!");
            }

            type = args[4].NumericValue;
        }

        // Validate type is 0 or 1
        if (type != 0.0 && type != 1.0)
        {
            return FormulaResult.Error("#NUM!");
        }

        // Validate pmt is not zero when rate is zero
        if (rate == 0.0 && pmt == 0.0)
        {
            return FormulaResult.Error("#NUM!");
        }

        double nper;

        // Special case: rate = 0
        // When rate is 0, PV + PMT*n + FV = 0, so n = -(PV + FV) / PMT
        // But Excel uses: n = (PV - FV) / PMT (accounting for sign conventions)
        if (rate == 0.0)
        {
            nper = (pv - fv) / pmt;
        }
        else
        {
            // Standard NPER formula
            var pmtWithType = pmt * (1 + rate * type);

            // Check for valid inputs to avoid log of negative number
            if (pmtWithType == 0.0)
            {
                return FormulaResult.Error("#NUM!");
            }

            var numerator = pmtWithType - fv * rate;
            var denominator = pmtWithType + pv * rate;

            // The ratio must be positive for log to be valid
            // This is true when numerator and denominator have the same sign
            var ratio = numerator / denominator;
            if (ratio <= 0.0)
            {
                return FormulaResult.Error("#NUM!");
            }

            nper = System.Math.Log(ratio) / System.Math.Log(1 + rate);
        }

        if (double.IsNaN(nper) || double.IsInfinity(nper) || nper < 0.0)
        {
            return FormulaResult.Error("#NUM!");
        }

        return FormulaResult.FromNumber(nper);
    }
}
