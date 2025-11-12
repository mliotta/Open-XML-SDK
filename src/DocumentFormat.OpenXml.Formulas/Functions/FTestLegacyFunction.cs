// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the FTEST function (legacy Excel 2007 compatibility).
/// FTEST(array1, array2) - returns the result of an F-test.
/// This is the legacy version; modern Excel uses F.TEST.
/// </summary>
public sealed class FTestLegacyFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly FTestLegacyFunction Instance = new();

    private FTestLegacyFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "FTEST";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        // FTEST has the same signature as F.TEST
        // Delegate directly to F.TEST
        return FTestFunction.Instance.Execute(context, args);
    }
}
