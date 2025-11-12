// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the WRAPCOLS function.
/// WRAPCOLS(vector, wrap_count, [pad_with]) - Wraps a row or column of values into columns.
/// NOTE: Due to single-value return limitation, only the first element is returned.
/// </summary>
public sealed class WrapColsFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly WrapColsFunction Instance = new();

    private WrapColsFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "WRAPCOLS";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 2)
        {
            return CellValue.Error("#VALUE!");
        }

        // Parse wrap_count parameter
        var wrapCountArg = args[args.Length - 1];
        if (wrapCountArg.IsError)
        {
            return wrapCountArg;
        }

        if (wrapCountArg.Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }

        var wrapCount = (int)wrapCountArg.NumericValue;

        if (wrapCount <= 0)
        {
            return CellValue.Error("#VALUE!");
        }

        // Parse optional pad_with parameter
        var hasPadWith = args.Length >= 3;
        var arrayLength = hasPadWith ? args.Length - 2 : args.Length - 1;

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

        // Return first element
        // In a full implementation, this would wrap values into columns
        return args[0];
    }
}
