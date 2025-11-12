// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the GETPIVOTDATA function.
/// GETPIVOTDATA(data_field, pivot_table, [field1, item1], [field2, item2], ...) - Extracts data from a pivot table.
/// </summary>
public sealed class GetPivotDataFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly GetPivotDataFunction Instance = new();

    private GetPivotDataFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "GETPIVOTDATA";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 2)
        {
            return CellValue.Error("#VALUE!");
        }

        var dataField = args[0];
        var pivotTable = args[1];

        // Check for errors
        if (dataField.IsError)
        {
            return dataField;
        }

        if (pivotTable.IsError)
        {
            return pivotTable;
        }

        // GETPIVOTDATA requires access to pivot table structures which is complex
        // For now, return #REF! to indicate this feature is not fully implemented
        return CellValue.Error("#REF!");
    }
}
