// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the FORECAST.ETS.STAT function.
/// FORECAST.ETS.STAT(values, timeline, statistic_type, [seasonality], [data_completion], [aggregation])
/// Returns statistical values from exponential smoothing forecasting.
/// Statistic types: 1=Alpha, 2=Beta, 3=Gamma, 4=MASE, 5=SMAPE, 6=MAE, 7=RMSE, 8=Step size
/// </summary>
public sealed class ForecastEtsStatFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly ForecastEtsStatFunction Instance = new();

    private ForecastEtsStatFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "FORECAST.ETS.STAT";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 3 || args.Length > 6)
        {
            return CellValue.Error("#VALUE!");
        }

        // Check for errors in required arguments
        for (int i = 0; i < System.Math.Min(3, args.Length); i++)
        {
            if (args[i].IsError)
            {
                return args[i];
            }
        }

        // Extract values array
        var values = new List<double>();
        if (args[0].Type == CellValueType.Number)
        {
            values.Add(args[0].NumericValue);
        }
        else if (args[0].IsError)
        {
            return args[0];
        }
        else
        {
            return CellValue.Error("#VALUE!");
        }

        // Extract timeline array
        var timeline = new List<double>();
        if (args[1].Type == CellValueType.Number)
        {
            timeline.Add(args[1].NumericValue);
        }
        else if (args[1].IsError)
        {
            return args[1];
        }
        else
        {
            return CellValue.Error("#VALUE!");
        }

        // Check arrays have same length
        if (values.Count != timeline.Count)
        {
            return CellValue.Error("#VALUE!");
        }

        // Need at least 2 data points
        if (values.Count < 2)
        {
            return CellValue.Error("#N/A");
        }

        // Get statistic_type
        if (args[2].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }
        int statisticType = (int)args[2].NumericValue;
        if (statisticType < 1 || statisticType > 8)
        {
            return CellValue.Error("#NUM!");
        }

        // Get optional seasonality parameter (default: 0 = auto-detect)
        int seasonality = 0;
        if (args.Length > 3 && args[3].Type == CellValueType.Number)
        {
            seasonality = (int)args[3].NumericValue;
            if (seasonality < 0)
            {
                return CellValue.Error("#NUM!");
            }
        }

        // Optional parameters: data_completion and aggregation
        // For Phase 0, we ignore these parameters

        try
        {
            // Sort timeline and values together
            var sorted = SortByTimeline(timeline.ToArray(), values.ToArray());
            double[] sortedTimeline = sorted.Timeline;
            double[] sortedValues = sorted.Values;

            // Validate timeline is strictly increasing
            for (int i = 1; i < sortedTimeline.Length; i++)
            {
                if (sortedTimeline[i] <= sortedTimeline[i - 1])
                {
                    return CellValue.Error("#VALUE!");
                }
            }

            // Perform Holt-Winters forecast
            var etsResult = ForecastHelper.HoltWintersForecast(
                sortedValues,
                seasonality,
                1); // Only need to fit, not forecast far ahead

            // Return the requested statistic
            double result = statisticType switch
            {
                1 => etsResult.Alpha,              // Alpha (base/level smoothing)
                2 => etsResult.Beta,               // Beta (trend smoothing)
                3 => etsResult.Gamma,              // Gamma (seasonality smoothing)
                4 => etsResult.MASE,               // Mean Absolute Scaled Error
                5 => etsResult.SMAPE,              // Symmetric Mean Absolute Percentage Error
                6 => etsResult.MAE,                // Mean Absolute Error
                7 => etsResult.RMSE,               // Root Mean Square Error
                8 => CalculateAverageStep(sortedTimeline), // Step size (average timeline interval)
                _ => throw new ArgumentException("Invalid statistic type"),
            };

            return CellValue.FromNumber(result);
        }
        catch (ArgumentException)
        {
            return CellValue.Error("#VALUE!");
        }
        catch (Exception)
        {
            return CellValue.Error("#N/A");
        }
    }

    /// <summary>
    /// Sorts timeline and values arrays together by timeline.
    /// </summary>
    private static SortedArrays SortByTimeline(double[] timeline, double[] values)
    {
        var pairs = new List<TimeValuePair>();
        for (int i = 0; i < timeline.Length; i++)
        {
            pairs.Add(new TimeValuePair { Time = timeline[i], Value = values[i] });
        }

        pairs.Sort((a, b) => a.Time.CompareTo(b.Time));

        double[] sortedTimeline = new double[pairs.Count];
        double[] sortedValues = new double[pairs.Count];

        for (int i = 0; i < pairs.Count; i++)
        {
            sortedTimeline[i] = pairs[i].Time;
            sortedValues[i] = pairs[i].Value;
        }

        return new SortedArrays { Timeline = sortedTimeline, Values = sortedValues };
    }

    private class TimeValuePair
    {
        public double Time { get; set; }
        public double Value { get; set; }
    }

    private class SortedArrays
    {
        public double[] Timeline { get; set; } = new double[0];
        public double[] Values { get; set; } = new double[0];
    }

    /// <summary>
    /// Calculates the average step size in the timeline.
    /// </summary>
    private static double CalculateAverageStep(double[] timeline)
    {
        if (timeline.Length < 2)
        {
            return 1.0;
        }

        double sum = 0.0;
        for (int i = 1; i < timeline.Length; i++)
        {
            sum += timeline[i] - timeline[i - 1];
        }

        return sum / (timeline.Length - 1);
    }
}
