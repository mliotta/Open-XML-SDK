// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the BAHTTEXT function.
/// BAHTTEXT(number) - Converts a number to Thai text and adds the Baht currency suffix.
///
/// NOTE: This function is not fully implemented. A complete implementation would require
/// converting numbers to Thai words (e.g., 1234.56 → "หนึ่งพันสองร้อยสามสิบสี่บาทห้าสิบหกสตางค์").
/// Currently throws UnsupportedFunctionException to avoid returning incorrect results.
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

        // BAHTTEXT requires complex Thai number-to-text conversion which is not yet implemented.
        // Throwing exception rather than returning incorrect results.
        throw new UnsupportedFunctionException(
            "BAHTTEXT is not fully implemented. Full implementation requires Thai number-to-text conversion.");
    }
}
