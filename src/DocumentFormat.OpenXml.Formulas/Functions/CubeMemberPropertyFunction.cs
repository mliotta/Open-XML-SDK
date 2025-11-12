// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the CUBEMEMBERPROPERTY function.
/// CUBEMEMBERPROPERTY(connection, member_expression, property) - returns the value of a member property from the cube.
/// Note: This function requires an OLAP connection and is not supported in this implementation.
/// </summary>
public sealed class CubeMemberPropertyFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly CubeMemberPropertyFunction Instance = new();

    private CubeMemberPropertyFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "CUBEMEMBERPROPERTY";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        // CUBEMEMBERPROPERTY requires an OLAP connection which is not available in this context
        // Return #REF! error as this function cannot be evaluated without OLAP data
        return CellValue.Error("#REF!");
    }
}
