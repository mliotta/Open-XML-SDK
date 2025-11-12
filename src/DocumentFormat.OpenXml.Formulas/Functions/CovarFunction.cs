// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the COVAR function (legacy compatibility function).
/// COVAR(array1, array2) - returns population covariance.
/// This is a legacy function that delegates to COVARIANCE.P.
/// </summary>
public sealed class CovarFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly CovarFunction Instance = new();

    private CovarFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "COVAR";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        // COVAR is identical to COVARIANCE.P
        return CovariancePFunction.Instance.Execute(context, args);
    }
}
