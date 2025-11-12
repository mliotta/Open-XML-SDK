// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the FISHER function.
/// FISHER(x) - returns the Fisher transformation of x.
/// </summary>
public sealed class FisherFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly FisherFunction Instance = new();

    private FisherFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "FISHER";

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

        double x = args[0].NumericValue;

        // x must be between -1 and 1 (exclusive)
        if (x <= -1 || x >= 1)
        {
            return CellValue.Error("#NUM!");
        }

        // Fisher transformation: 0.5 * ln((1 + x) / (1 - x))
        double result = 0.5 * System.Math.Log((1 + x) / (1 - x));

        return CellValue.FromNumber(result);
    }
}
