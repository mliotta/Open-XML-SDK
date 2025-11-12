// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the CUBEVALUE function.
/// CUBEVALUE(connection, [member_expression1], [member_expression2], ...) - returns an aggregated value from the cube.
/// Note: This function requires an OLAP connection and is not supported in this implementation.
/// </summary>
public sealed class CubeValueFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly CubeValueFunction Instance = new();

    private CubeValueFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "CUBEVALUE";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        // CUBEVALUE requires an OLAP connection which is not available in this context
        // Return #REF! error as this function cannot be evaluated without OLAP data
        return CellValue.Error("#REF!");
    }
}
