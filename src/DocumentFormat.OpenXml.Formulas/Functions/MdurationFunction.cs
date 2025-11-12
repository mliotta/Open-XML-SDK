// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the MDURATION function.
/// MDURATION(settlement, maturity, coupon, yld, frequency, [basis]) - returns the modified Macaulay duration for a security with an assumed par value of $100.
/// </summary>
public sealed class MdurationFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly MdurationFunction Instance = new();

    private MdurationFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "MDURATION";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 5 || args.Length > 6)
        {
            return CellValue.Error("#VALUE!");
        }

        // Use DURATION function to calculate Macaulay duration
        var durationResult = DurationFunction.Instance.Execute(context, args);

        if (durationResult.IsError)
        {
            return durationResult;
        }

        try
        {
            var yld = args[3].NumericValue;
            var frequency = (int)args[4].NumericValue;
            var macaulayDuration = durationResult.NumericValue;

            // Modified duration = Macaulay duration / (1 + yield/frequency)
            var modifiedDuration = macaulayDuration / (1 + yld / frequency);

            if (double.IsNaN(modifiedDuration) || double.IsInfinity(modifiedDuration) || modifiedDuration < 0)
            {
                return CellValue.Error("#NUM!");
            }

            return CellValue.FromNumber(modifiedDuration);
        }
        catch
        {
            return CellValue.Error("#NUM!");
        }
    }
}
