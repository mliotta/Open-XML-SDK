// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the NOW function.
/// NOW() - returns current date and time (no arguments).
/// </summary>
public sealed class NowFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly NowFunction Instance = new();

    private NowFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "NOW";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 0)
        {
            return CellValue.Error("#VALUE!");
        }

        // Excel stores dates as OLE Automation dates (days since 1/1/1900)
        var serialDate = DateTime.Now.ToOADate();
        return CellValue.FromNumber(serialDate);
    }
}
