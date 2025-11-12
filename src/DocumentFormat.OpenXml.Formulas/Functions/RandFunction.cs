// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the RAND function.
/// RAND() - returns a random number between 0 and 1.
/// Note: This function is volatile and recalculates each time it is evaluated.
/// </summary>
public sealed class RandFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly RandFunction Instance = new();

    private static readonly Random _random = new();

    private RandFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "RAND";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 0)
        {
            return CellValue.Error("#VALUE!");
        }

        return CellValue.FromNumber(_random.NextDouble());
    }
}
