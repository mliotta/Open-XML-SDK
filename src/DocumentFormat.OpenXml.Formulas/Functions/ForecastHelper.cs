// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Linq;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Helper class for exponential smoothing forecasting (ETS) calculations.
/// Implements Holt-Winters Exponential Triple Smoothing for time series forecasting.
/// </summary>
internal static class ForecastHelper
{
    /// <summary>
    /// Result of Holt-Winters exponential smoothing forecast.
    /// </summary>
    internal class ETSResult
    {
        public double Alpha { get; set; }
        public double Beta { get; set; }
        public double Gamma { get; set; }
        public double[] Level { get; set; } = new double[0];
        public double[] Trend { get; set; } = new double[0];
        public double[] Seasonal { get; set; } = new double[0];
        public double[] FittedValues { get; set; } = new double[0];
        public double[] Residuals { get; set; } = new double[0];
        public int SeasonalPeriod { get; set; }
        public double MAE { get; set; }
        public double RMSE { get; set; }
        public double MASE { get; set; }
        public double SMAPE { get; set; }
    }

    /// <summary>
    /// Detects the seasonal period in a time series using autocorrelation.
    /// </summary>
    /// <param name="values">The time series values.</param>
    /// <param name="maxPeriod">Maximum period to search (default: min(values.Length/2, 24)).</param>
    /// <returns>The detected seasonal period (0 or 1 if no seasonality detected).</returns>
    public static int DetectSeasonality(double[] values, int? maxPeriod = null)
    {
        if (values.Length < 4)
        {
            return 0; // Not enough data for seasonality detection
        }

        int maxP = maxPeriod ?? System.Math.Min(values.Length / 2, 24);
        if (maxP < 2)
        {
            return 0;
        }

        // Calculate autocorrelation for different lags
        double maxCorrelation = 0.0;
        int bestPeriod = 0;

        for (int lag = 2; lag <= maxP; lag++)
        {
            double correlation = CalculateAutocorrelation(values, lag);
            if (correlation > maxCorrelation && correlation > 0.5) // Threshold for significance
            {
                maxCorrelation = correlation;
                bestPeriod = lag;
            }
        }

        // Return 0 if no significant seasonality found
        return bestPeriod > 0 ? bestPeriod : 0;
    }

    /// <summary>
    /// Calculates autocorrelation at a specific lag.
    /// </summary>
    private static double CalculateAutocorrelation(double[] values, int lag)
    {
        if (lag >= values.Length)
        {
            return 0.0;
        }

        double mean = values.Average();
        double variance = 0.0;
        double covariance = 0.0;

        for (int i = 0; i < values.Length; i++)
        {
            variance += (values[i] - mean) * (values[i] - mean);
        }

        for (int i = 0; i < values.Length - lag; i++)
        {
            covariance += (values[i] - mean) * (values[i + lag] - mean);
        }

        if (variance == 0.0)
        {
            return 0.0;
        }

        return covariance / variance;
    }

    /// <summary>
    /// Performs Holt-Winters exponential smoothing forecast.
    /// </summary>
    /// <param name="values">The time series values.</param>
    /// <param name="seasonalPeriod">The seasonal period (0 for auto-detect, 1 for no seasonality).</param>
    /// <param name="stepsAhead">Number of steps to forecast ahead.</param>
    /// <param name="alpha">Smoothing parameter for level (null for auto-optimization).</param>
    /// <param name="beta">Smoothing parameter for trend (null for auto-optimization).</param>
    /// <param name="gamma">Smoothing parameter for seasonality (null for auto-optimization).</param>
    /// <returns>ETS result containing forecast and statistics.</returns>
    public static ETSResult HoltWintersForecast(
        double[] values,
        int seasonalPeriod = 0,
        int stepsAhead = 1,
        double? alpha = null,
        double? beta = null,
        double? gamma = null)
    {
        if (values.Length < 2)
        {
            throw new ArgumentException("Need at least 2 data points for forecasting");
        }

        // Auto-detect seasonality if requested
        if (seasonalPeriod == 0)
        {
            seasonalPeriod = DetectSeasonality(values);
            if (seasonalPeriod == 0)
            {
                seasonalPeriod = 1; // No seasonality
            }
        }

        bool hasSeason = seasonalPeriod > 1 && values.Length >= 2 * seasonalPeriod;

        // Use default parameters if not provided
        double alphaVal = alpha ?? 0.3;
        double betaVal = beta ?? 0.1;
        double gammaVal = gamma ?? (hasSeason ? 0.1 : 0.0);

        // If parameters not provided, optimize them
        if (!alpha.HasValue || !beta.HasValue || (!gamma.HasValue && hasSeason))
        {
            OptimizeParameters(values, seasonalPeriod, out alphaVal, out betaVal, out gammaVal);
        }

        return ApplyHoltWinters(values, seasonalPeriod, stepsAhead, alphaVal, betaVal, gammaVal);
    }

    /// <summary>
    /// Applies Holt-Winters smoothing with given parameters.
    /// </summary>
    private static ETSResult ApplyHoltWinters(
        double[] values,
        int seasonalPeriod,
        int stepsAhead,
        double alpha,
        double beta,
        double gamma)
    {
        int n = values.Length;
        bool hasSeason = seasonalPeriod > 1 && n >= 2 * seasonalPeriod;

        double[] level = new double[n];
        double[] trend = new double[n];
        double[] seasonal = new double[n + (hasSeason ? seasonalPeriod : 0)];
        double[] fitted = new double[n];
        double[] residuals = new double[n];

        // Initialize level, trend, and seasonal components
        InitializeComponents(values, seasonalPeriod, level, trend, seasonal, hasSeason);

        // Apply Holt-Winters smoothing
        for (int t = 0; t < n; t++)
        {
            // Calculate fitted value
            if (hasSeason)
            {
                fitted[t] = (level[t] + trend[t]) * seasonal[t];
            }
            else
            {
                fitted[t] = level[t] + trend[t];
            }

            residuals[t] = values[t] - fitted[t];

            // Update components for next iteration (except last)
            if (t < n - 1)
            {
                double prevLevel = level[t];
                double prevTrend = trend[t];

                if (hasSeason)
                {
                    // Multiplicative seasonality
                    level[t + 1] = alpha * (values[t] / seasonal[t]) + (1 - alpha) * (prevLevel + prevTrend);
                    trend[t + 1] = beta * (level[t + 1] - prevLevel) + (1 - beta) * prevTrend;
                    seasonal[t + seasonalPeriod] = gamma * (values[t] / level[t + 1]) + (1 - gamma) * seasonal[t];
                }
                else
                {
                    // No seasonality (Holt's method)
                    level[t + 1] = alpha * values[t] + (1 - alpha) * (prevLevel + prevTrend);
                    trend[t + 1] = beta * (level[t + 1] - prevLevel) + (1 - beta) * prevTrend;
                }
            }
        }

        // Calculate error metrics
        double mae = CalculateMAE(residuals);
        double rmse = CalculateRMSE(residuals);
        double mase = CalculateMASE(values, residuals);
        double smape = CalculateSMAPE(values, fitted);

        return new ETSResult
        {
            Alpha = alpha,
            Beta = beta,
            Gamma = gamma,
            Level = level,
            Trend = trend,
            Seasonal = seasonal,
            FittedValues = fitted,
            Residuals = residuals,
            SeasonalPeriod = seasonalPeriod,
            MAE = mae,
            RMSE = rmse,
            MASE = mase,
            SMAPE = smape,
        };
    }

    /// <summary>
    /// Initializes level, trend, and seasonal components.
    /// </summary>
    private static void InitializeComponents(
        double[] values,
        int seasonalPeriod,
        double[] level,
        double[] trend,
        double[] seasonal,
        bool hasSeason)
    {
        int n = values.Length;

        if (hasSeason)
        {
            // Initialize level as average of first season
            level[0] = values.Take(seasonalPeriod).Average();

            // Initialize trend as average difference between first two seasons
            if (n >= 2 * seasonalPeriod)
            {
                double sum1 = values.Take(seasonalPeriod).Sum();
                double sum2 = values.Skip(seasonalPeriod).Take(seasonalPeriod).Sum();
                trend[0] = (sum2 - sum1) / (seasonalPeriod * seasonalPeriod);
            }
            else
            {
                trend[0] = 0.0;
            }

            // Initialize seasonal indices
            for (int i = 0; i < seasonalPeriod; i++)
            {
                double sum = 0.0;
                int count = 0;
                for (int j = i; j < n; j += seasonalPeriod)
                {
                    sum += values[j];
                    count++;
                }
                double seasonalAvg = sum / count;
                seasonal[i] = seasonalAvg / level[0];
            }

            // Normalize seasonal indices to sum to seasonalPeriod
            double seasonalSum = seasonal.Take(seasonalPeriod).Sum();
            if (seasonalSum > 0)
            {
                for (int i = 0; i < seasonalPeriod; i++)
                {
                    seasonal[i] = seasonal[i] * seasonalPeriod / seasonalSum;
                }
            }
        }
        else
        {
            // Simple initialization without seasonality
            level[0] = values[0];
            trend[0] = n > 1 ? (values[1] - values[0]) : 0.0;
        }
    }

    /// <summary>
    /// Optimizes smoothing parameters using grid search.
    /// </summary>
    private static void OptimizeParameters(
        double[] values,
        int seasonalPeriod,
        out double alpha,
        out double beta,
        out double gamma)
    {
        bool hasSeason = seasonalPeriod > 1 && values.Length >= 2 * seasonalPeriod;

        double bestAlpha = 0.3;
        double bestBeta = 0.1;
        double bestGamma = 0.1;
        double bestError = double.MaxValue;

        // Grid search (simplified for Phase 0)
        double[] alphaRange = { 0.1, 0.3, 0.5 };
        double[] betaRange = { 0.05, 0.1, 0.2 };
        double[] gammaRange = hasSeason ? new[] { 0.05, 0.1, 0.2 } : new[] { 0.0 };

        foreach (double a in alphaRange)
        {
            foreach (double b in betaRange)
            {
                foreach (double g in gammaRange)
                {
                    try
                    {
                        var result = ApplyHoltWinters(values, seasonalPeriod, 1, a, b, g);
                        double error = result.RMSE;
                        if (error < bestError)
                        {
                            bestError = error;
                            bestAlpha = a;
                            bestBeta = b;
                            bestGamma = g;
                        }
                    }
                    catch
                    {
                        // Skip invalid parameter combinations
                    }
                }
            }
        }

        alpha = bestAlpha;
        beta = bestBeta;
        gamma = bestGamma;
    }

    /// <summary>
    /// Forecasts future values using an ETS result.
    /// </summary>
    /// <param name="etsResult">The ETS result from training.</param>
    /// <param name="stepsAhead">Number of steps to forecast.</param>
    /// <returns>Array of forecasted values.</returns>
    public static double[] ForecastValues(ETSResult etsResult, int stepsAhead)
    {
        double[] forecasts = new double[stepsAhead];
        int n = etsResult.Level.Length;
        double lastLevel = etsResult.Level[n - 1];
        double lastTrend = etsResult.Trend[n - 1];
        bool hasSeason = etsResult.SeasonalPeriod > 1;

        for (int h = 1; h <= stepsAhead; h++)
        {
            if (hasSeason)
            {
                int seasonalIndex = (n - etsResult.SeasonalPeriod + h - 1) % etsResult.SeasonalPeriod;
                forecasts[h - 1] = (lastLevel + h * lastTrend) * etsResult.Seasonal[seasonalIndex];
            }
            else
            {
                forecasts[h - 1] = lastLevel + h * lastTrend;
            }
        }

        return forecasts;
    }

    /// <summary>
    /// Calculates Mean Absolute Error.
    /// </summary>
    private static double CalculateMAE(double[] residuals)
    {
        if (residuals.Length == 0)
        {
            return 0.0;
        }

        return residuals.Select(System.Math.Abs).Average();
    }

    /// <summary>
    /// Calculates Root Mean Square Error.
    /// </summary>
    private static double CalculateRMSE(double[] residuals)
    {
        if (residuals.Length == 0)
        {
            return 0.0;
        }

        return System.Math.Sqrt(residuals.Select(r => r * r).Average());
    }

    /// <summary>
    /// Calculates Mean Absolute Scaled Error (MASE).
    /// </summary>
    private static double CalculateMASE(double[] values, double[] residuals)
    {
        if (values.Length < 2 || residuals.Length == 0)
        {
            return 0.0;
        }

        // Calculate MAE of residuals
        double mae = CalculateMAE(residuals);

        // Calculate MAE of naive forecast (one-step ahead)
        double naiveMae = 0.0;
        for (int i = 1; i < values.Length; i++)
        {
            naiveMae += System.Math.Abs(values[i] - values[i - 1]);
        }
        naiveMae /= (values.Length - 1);

        if (naiveMae == 0.0)
        {
            return 0.0;
        }

        return mae / naiveMae;
    }

    /// <summary>
    /// Calculates Symmetric Mean Absolute Percentage Error (SMAPE).
    /// </summary>
    private static double CalculateSMAPE(double[] actual, double[] fitted)
    {
        if (actual.Length == 0 || fitted.Length == 0 || actual.Length != fitted.Length)
        {
            return 0.0;
        }

        double sum = 0.0;
        int count = 0;

        for (int i = 0; i < actual.Length; i++)
        {
            double denominator = System.Math.Abs(actual[i]) + System.Math.Abs(fitted[i]);
            if (denominator > 0)
            {
                sum += System.Math.Abs(actual[i] - fitted[i]) / denominator;
                count++;
            }
        }

        if (count == 0)
        {
            return 0.0;
        }

        return 200.0 * sum / count;
    }

    /// <summary>
    /// Calculates confidence interval for forecast.
    /// </summary>
    /// <param name="etsResult">The ETS result.</param>
    /// <param name="stepsAhead">Steps ahead for forecast.</param>
    /// <param name="confidenceLevel">Confidence level (0-1, default 0.95).</param>
    /// <returns>Half-width of confidence interval.</returns>
    public static double CalculateConfidenceInterval(ETSResult etsResult, int stepsAhead, double confidenceLevel)
    {
        if (confidenceLevel <= 0 || confidenceLevel >= 1)
        {
            throw new ArgumentException("Confidence level must be between 0 and 1");
        }

        // Calculate standard error from residuals
        double se = etsResult.RMSE;

        // Adjust for forecast horizon (uncertainty increases with distance)
        double adjustedSE = se * System.Math.Sqrt(1.0 + stepsAhead * 0.1);

        // Use normal approximation for confidence interval
        // For 95% confidence: z = 1.96, for 90%: z = 1.645, etc.
        double z = StatisticalHelper.NormSInv(0.5 + confidenceLevel / 2.0);

        return z * adjustedSE;
    }
}
