// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the CUBEKPIMEMBER function.
/// CUBEKPIMEMBER(connection, kpi_name, kpi_property, [caption]) - returns a key performance indicator (KPI) property.
/// Note: This function requires an OLAP connection and is not supported in this implementation.
/// </summary>
public sealed class CubeKpiMemberFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly CubeKpiMemberFunction Instance = new();

    private CubeKpiMemberFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "CUBEKPIMEMBER";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        // CUBEKPIMEMBER requires an OLAP connection which is not available in this context
        // Return #REF! error as this function cannot be evaluated without OLAP data
        return CellValue.Error("#REF!");
    }
}
