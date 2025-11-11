// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation;

/// <summary>
/// Statistics about formula evaluation performance.
/// </summary>
public class EvaluatorStatistics
{
    /// <summary>
    /// Total number of formulas evaluated since creation.
    /// </summary>
    public long TotalEvaluations { get; set; }

    /// <summary>
    /// Number of successful evaluations.
    /// </summary>
    public long SuccessfulEvaluations { get; set; }

    /// <summary>
    /// Number of failed evaluations.
    /// </summary>
    public long FailedEvaluations { get; set; }

    /// <summary>
    /// Success rate (0.0 to 1.0).
    /// </summary>
    public double SuccessRate =>
        TotalEvaluations > 0 ? (double)SuccessfulEvaluations / TotalEvaluations : 0;

    /// <summary>
    /// Number of unique formulas compiled and cached.
    /// </summary>
    public int CompiledFormulaCount { get; set; }

    /// <summary>
    /// Number of functions currently supported.
    /// </summary>
    public int SupportedFunctionCount { get; set; }

    /// <summary>
    /// Average evaluation time in microseconds.
    /// </summary>
    public double AvgEvaluationTimeMicros { get; set; }
}
