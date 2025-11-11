// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the CHAR function.
/// CHAR(number) - returns character for ASCII code.
/// </summary>
public sealed class CharFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly CharFunction Instance = new();

    private CharFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "CHAR";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 1)
        {
            return CellValue.Error("#VALUE!");
        }

        if (args[0].IsError)
        {
            return args[0];
        }

        if (args[0].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }

        var number = (int)args[0].NumericValue;

        // Valid character codes are 1-255 in Excel
        if (number < 1 || number > 255)
        {
            return CellValue.Error("#VALUE!");
        }

        var character = ((char)number).ToString();

        return CellValue.FromString(character);
    }
}
