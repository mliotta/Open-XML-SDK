// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the QUARTILE function.
/// QUARTILE(array, quart) - returns the quartile (0, 1, 2, 3, or 4).
/// Uses PERCENTILE internally: 0=min, 1=25%, 2=50%, 3=75%, 4=max.
/// </summary>
public sealed class QuartileFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly QuartileFunction Instance = new();

    private QuartileFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "QUARTILE";

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

        // quart must be 0, 1, 2, 3, or 4
        if (quart < 0 || quart > 4)
        {
            return CellValue.Error("#NUM!");
        }

        // Map quartile to percentile
        double percentile;
        switch (quart)
        {
            case 0:
                percentile = 0.0; // Minimum
                break;
            case 1:
                percentile = 0.25; // First quartile
                break;
            case 2:
                percentile = 0.5; // Median
                break;
            case 3:
                percentile = 0.75; // Third quartile
                break;
            case 4:
                percentile = 1.0; // Maximum
                break;
            default:
                return CellValue.Error("#NUM!");
        }

        // Use PERCENTILE function to calculate the result
        var percentileArgs = new[]
        {
            args[0],
            CellValue.FromNumber(percentile),
        };

        return PercentileFunction.Instance.Execute(context, percentileArgs);
    }
}
