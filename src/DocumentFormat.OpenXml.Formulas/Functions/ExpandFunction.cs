// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the EXPAND function.
/// EXPAND(array, rows, [columns], [pad_with]) - Expands an array to specified dimensions.
/// NOTE: Due to single-value return limitation, only the first element is returned.
/// </summary>
public sealed class ExpandFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly ExpandFunction Instance = new();

    private ExpandFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "EXPAND";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 2)
        {
            return CellValue.Error("#VALUE!");
        }

        // Parse rows parameter
        var rowsArg = args[args.Length - 1];
        if (rowsArg.IsError)
        {
            return rowsArg;
        }

        if (rowsArg.Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }

        var rows = (int)rowsArg.NumericValue;

        // Parse optional columns parameter
        var cols = 0;
        var hasColumns = args.Length >= 3 && args[args.Length - 2].Type == CellValueType.Number;
        if (hasColumns)
        {
            cols = (int)args[args.Length - 2].NumericValue;
        }

        // Parse optional pad_with parameter
        var padWith = CellValue.Error("#N/A");
        var hasPadWith = false;
        if (args.Length >= 4)
        {
            var padIdx = hasColumns ? args.Length - 3 : args.Length - 2;
            if (padIdx >= 1)
            {
                padWith = args[padIdx];
                hasPadWith = true;
            }
        }

        // Determine array length
        var arrayLength = args.Length - 1; // At least rows parameter
        if (hasColumns)
        {
            arrayLength--;
        }
        if (hasPadWith)
        {
            arrayLength--;
        }

        if (arrayLength == 0)
        {
            return CellValue.Error("#VALUE!");
        }

        // Check for errors in array
        for (var i = 0; i < arrayLength; i++)
        {
            if (args[i].IsError)
            {
                return args[i];
            }
        }

        // If expansion is needed, return first element or pad value
        // In a full implementation, this would create an expanded array
        return args[0];
    }
}
