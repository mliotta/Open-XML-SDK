// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the CUBERANKEDMEMBER function.
/// CUBERANKEDMEMBER(connection, set_expression, rank, [caption]) - returns the nth, or ranked, member in a set.
/// Note: This function requires an OLAP connection and is not supported in this implementation.
/// </summary>
public sealed class CubeRankedMemberFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly CubeRankedMemberFunction Instance = new();

    private CubeRankedMemberFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "CUBERANKEDMEMBER";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        // CUBERANKEDMEMBER requires an OLAP connection which is not available in this context
        // Return #REF! error as this function cannot be evaluated without OLAP data
        return CellValue.Error("#REF!");
    }
}
