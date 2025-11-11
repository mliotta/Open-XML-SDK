// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the MINUTE function.
/// MINUTE(time) - extracts minute (0-59).
/// </summary>
public sealed class MinuteFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly MinuteFunction Instance = new();

    private MinuteFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "MINUTE";

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

        try
        {
            var dateTime = DateTime.FromOADate(args[0].NumericValue);
            return CellValue.FromNumber(dateTime.Minute);
        }
        catch
        {
            return CellValue.Error("#NUM!");
        }
    }
}
