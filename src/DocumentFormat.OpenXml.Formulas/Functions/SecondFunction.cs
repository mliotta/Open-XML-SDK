// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the SECOND function.
/// SECOND(time) - extracts second (0-59).
/// </summary>
public sealed class SecondFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly SecondFunction Instance = new();

    private SecondFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "SECOND";

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
            return CellValue.FromNumber(dateTime.Second);
        }
        catch
        {
            return CellValue.Error("#NUM!");
        }
    }
}
