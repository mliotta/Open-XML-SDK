// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Linq;

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the SUM function.
/// </summary>
public sealed class SumFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly SumFunction Instance = new();

    private SumFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "SUM";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        var sum = 0.0;

        foreach (var arg in args)
        {
            if (arg.Type == CellValueType.Number)
            {
                sum += arg.NumericValue;
            }
            else if (arg.IsError)
            {
                return arg; // Propagate errors
            }
        }

        return CellValue.FromNumber(sum);
    }
}
