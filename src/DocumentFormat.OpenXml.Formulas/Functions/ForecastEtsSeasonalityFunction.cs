// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the FORECAST.ETS.SEASONALITY function.
/// FORECAST.ETS.SEASONALITY(values, timeline, [data_completion], [aggregation])
/// Returns the length of the detected seasonal pattern in the data.
/// </summary>
public sealed class ForecastEtsSeasonalityFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly ForecastEtsSeasonalityFunction Instance = new();

    private ForecastEtsSeasonalityFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "FORECAST.ETS.SEASONALITY";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 2 || args.Length > 4)
        {
            return CellValue.Error("#VALUE!");
        }

        // Check for errors in required arguments
        for (int i = 0; i < System.Math.Min(2, args.Length); i++)
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

        // Need at least 4 data points for seasonality detection
        if (values.Count < 4)
        {
            return CellValue.FromNumber(1); // No seasonality
        }

        // Optional parameters: data_completion and aggregation
        // For Phase 0, we ignore these parameters

        try
        {
            // Sort timeline and values together
            var sorted = SortByTimeline(timeline.ToArray(), values.ToArray());
            double[] sortedValues = sorted.Values;

            // Detect seasonality
            int seasonalPeriod = ForecastHelper.DetectSeasonality(sortedValues);

            // Return 1 if no seasonality detected (as per Excel behavior)
            if (seasonalPeriod == 0)
            {
                seasonalPeriod = 1;
            }

            return CellValue.FromNumber(seasonalPeriod);
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
}
