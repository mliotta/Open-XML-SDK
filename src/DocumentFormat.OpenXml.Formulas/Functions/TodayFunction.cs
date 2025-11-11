// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the TODAY function.
/// TODAY() - returns current date (no arguments).
/// </summary>
public sealed class TodayFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly TodayFunction Instance = new();

    private TodayFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "TODAY";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 0)
        {
            return CellValue.Error("#VALUE!");
        }

        // Excel stores dates as OLE Automation dates (days since 1/1/1900)
        var serialDate = DateTime.Today.ToOADate();
        return CellValue.FromNumber(serialDate);
    }
}
