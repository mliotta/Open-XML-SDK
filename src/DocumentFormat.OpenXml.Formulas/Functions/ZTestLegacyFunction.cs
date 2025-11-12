// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the ZTEST function (legacy version, same as Z.TEST).
/// ZTEST(array, x, [sigma]) - returns the one-tailed P-value of a z-test.
/// </summary>
public sealed class ZTestLegacyFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly ZTestLegacyFunction Instance = new();

    private ZTestLegacyFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "ZTEST";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        // ZTEST is the same as Z.TEST
        return ZTestFunction.Instance.Execute(context, args);
    }
}
