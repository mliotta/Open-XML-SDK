// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the FORECAST.LINEAR function.
/// FORECAST.LINEAR(x, known_y's, known_x's) - calculates a future value using linear regression.
/// This is the modern name for the FORECAST function with identical behavior.
/// </summary>
public sealed class ForecastLinearFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly ForecastLinearFunction Instance = new();

    private ForecastLinearFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "FORECAST.LINEAR";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        // FORECAST.LINEAR is identical to FORECAST
        return ForecastFunction.Instance.Execute(context, args);
    }
}
