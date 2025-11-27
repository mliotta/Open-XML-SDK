// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the AVERAGEA function.
/// AVERAGEA(value1, [value2], ...) - Average including text and logical values.
/// Text evaluates as 0, TRUE as 1, FALSE as 0, empty values are ignored.
/// </summary>
public sealed class AverageAFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly AverageAFunction Instance = new();

    private AverageAFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "AVERAGEA";

    /// <inheritdoc/>
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        var sum = 0.0;
        var count = 0;

        foreach (var arg in args)
        {
            if (arg.IsError)
            {
                return arg; // Propagate errors
            }

            if (arg.Type == FormulaResultType.Number)
            {
                sum += arg.NumericValue;
                count++;
            }
            else if (arg.Type == FormulaResultType.Boolean)
            {
                sum += arg.BoolValue ? 1.0 : 0.0;
                count++;
            }
            else if (arg.Type == FormulaResultType.Text)
            {
                // Text values count as 0
                sum += 0.0;
                count++;
            }
            // Empty values are ignored
        }

        if (count == 0)
        {
            return FormulaResult.Error("#DIV/0!");
        }

        return FormulaResult.FromNumber(sum / count);
    }
}
