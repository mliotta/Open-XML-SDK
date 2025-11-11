// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the PI function.
/// PI() - returns the value of π (pi).
/// </summary>
public sealed class PiFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly PiFunction Instance = new();

    private PiFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "PI";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 0)
        {
            return CellValue.Error("#VALUE!");
        }

        return CellValue.FromNumber(System.Math.PI);
    }
}
