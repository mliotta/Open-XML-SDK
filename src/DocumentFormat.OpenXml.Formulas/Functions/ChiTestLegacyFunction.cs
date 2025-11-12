// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the CHITEST function (legacy Excel 2007 compatibility).
/// CHITEST(actual_range, expected_range) - returns the test for independence.
/// This is the legacy version; modern Excel uses CHISQ.TEST.
/// </summary>
public sealed class ChiTestLegacyFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly ChiTestLegacyFunction Instance = new();

    private ChiTestLegacyFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "CHITEST";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        // CHITEST is equivalent to CHISQ.TEST
        return ChiSqTestFunction.Instance.Execute(context, args);
    }
}
