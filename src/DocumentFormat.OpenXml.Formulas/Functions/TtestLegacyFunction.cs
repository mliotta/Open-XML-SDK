// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the TTEST function (legacy compatibility function).
/// TTEST(array1, array2, tails, type) - returns the probability associated with a Student's t-test.
/// This is a legacy function that delegates to T.TEST.
/// </summary>
public sealed class TtestLegacyFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly TtestLegacyFunction Instance = new();

    private TtestLegacyFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "TTEST";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        // TTEST is identical to T.TEST
        return TTestFunction.Instance.Execute(context, args);
    }
}
