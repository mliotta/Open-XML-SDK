// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the DATE function.
/// DATE(year, month, day) - returns date serial number.
/// </summary>
public sealed class DateFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly DateFunction Instance = new();

    private DateFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "DATE";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 3)
        {
            return CellValue.Error("#VALUE!");
        }

        if (args[0].IsError)
        {
            return args[0];
        }

        if (args[1].IsError)
        {
            return args[1];
        }

        if (args[2].IsError)
        {
            return args[2];
        }

        if (args[0].Type != CellValueType.Number || args[1].Type != CellValueType.Number || args[2].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }

        var year = (int)args[0].NumericValue;
        var month = (int)args[1].NumericValue;
        var day = (int)args[2].NumericValue;

        try
        {
            var date = new DateTime(year, month, day);
            var serialDate = date.ToOADate();
            return CellValue.FromNumber(serialDate);
        }
        catch
        {
            return CellValue.Error("#NUM!");
        }
    }
}
