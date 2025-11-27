// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Linq;

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the AVERAGE function.
/// </summary>
public sealed class AverageFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly AverageFunction Instance = new();

    private AverageFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "AVERAGE";

    /// <inheritdoc/>
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        var sum = 0.0;
        var count = 0;

        foreach (var arg in args)
        {
            if (arg.Type == FormulaResultType.Number)
            {
                sum += arg.NumericValue;
                count++;
            }
            else if (arg.IsError)
            {
                return arg; // Propagate errors
            }
        }

        if (count == 0)
        {
            return FormulaResult.Error("#DIV/0!");
        }

        return FormulaResult.FromNumber(sum / count);
    }
}
