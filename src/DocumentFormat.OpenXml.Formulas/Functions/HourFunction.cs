// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the HOUR function.
/// HOUR(time) - extracts hour (0-23).
/// </summary>
public sealed class HourFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly HourFunction Instance = new();

    private HourFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "HOUR";

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
            return CellValue.FromNumber(dateTime.Hour);
        }
        catch
        {
            return CellValue.Error("#NUM!");
        }
    }
}
