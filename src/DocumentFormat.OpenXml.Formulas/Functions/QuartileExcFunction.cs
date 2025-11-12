// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the QUARTILE.EXC function.
/// QUARTILE.EXC(array, quart) - returns the quartile (1, 2, or 3).
/// Uses PERCENTILE.EXC internally: 1=25%, 2=50%, 3=75%.
/// Note: 0 and 4 are not valid for the exclusive method.
/// </summary>
public sealed class QuartileExcFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly QuartileExcFunction Instance = new();

    private QuartileExcFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "QUARTILE.EXC";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 2)
        {
            return CellValue.Error("#VALUE!");
        }

        // Propagate errors
        if (args[0].IsError)
        {
            return args[0];
        }

        if (args[1].IsError)
        {
            return args[1];
        }

        // Get quart value
        if (args[1].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }

        var quart = (int)args[1].NumericValue;

        // quart must be 1, 2, or 3 (exclusive method doesn't support 0 and 4)
        if (quart < 1 || quart > 3)
        {
            return CellValue.Error("#NUM!");
        }

        // Map quartile to percentile
        double percentile;
        switch (quart)
        {
            case 1:
                percentile = 0.25; // First quartile
                break;
            case 2:
                percentile = 0.5; // Median
                break;
            case 3:
                percentile = 0.75; // Third quartile
                break;
            default:
                return CellValue.Error("#NUM!");
        }

        // Use PERCENTILE.EXC function to calculate the result
        var percentileArgs = new[]
        {
            args[0],
            CellValue.FromNumber(percentile),
        };

        return PercentileExcFunction.Instance.Execute(context, percentileArgs);
    }
}
