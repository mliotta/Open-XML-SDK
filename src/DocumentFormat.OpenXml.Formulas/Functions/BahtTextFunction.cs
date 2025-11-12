// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the BAHTTEXT function.
/// BAHTTEXT(number) - Converts a number to Thai text and adds the Baht currency suffix.
/// Phase 0: Simplified implementation returns number with " บาท" suffix.
/// Full implementation would require Thai number-to-text conversion.
/// </summary>
public sealed class BahtTextFunction : IFunctionImplementation
{
    public static readonly BahtTextFunction Instance = new();

    private BahtTextFunction()
    {
    }

    public string Name => "BAHTTEXT";

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

        var number = args[0].NumericValue;

        // Phase 0 simplified: return number with Baht symbol
        // Full implementation would convert number to Thai words
        return CellValue.FromString(number.ToString("F2") + " บาท");
    }
}
