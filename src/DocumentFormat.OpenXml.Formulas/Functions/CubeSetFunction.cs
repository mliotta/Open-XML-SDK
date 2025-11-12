// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the CUBESET function.
/// CUBESET(connection, set_expression, [caption], [sort_order], [sort_by]) - defines a calculated set of members or tuples.
/// Note: This function requires an OLAP connection and is not supported in this implementation.
/// </summary>
public sealed class CubeSetFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly CubeSetFunction Instance = new();

    private CubeSetFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "CUBESET";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        // CUBESET requires an OLAP connection which is not available in this context
        // Return #REF! error as this function cannot be evaluated without OLAP data
        return CellValue.Error("#REF!");
    }
}
