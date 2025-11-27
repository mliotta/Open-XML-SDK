// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Text;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the REPT function.
/// REPT(text, number_times) - repeats text a given number of times.
/// </summary>
public sealed class ReptFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly ReptFunction Instance = new();

    private ReptFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "REPT";

    /// <inheritdoc/>
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        if (args.Length != 2)
        {
            return FormulaResult.Error("#VALUE!");
        }

        if (args[0].IsError)
        {
            return args[0];
        }

        if (args[1].IsError)
        {
            return args[1];
        }

        var text = args[0].StringValue;

        if (args[1].Type != FormulaResultType.Number)
        {
            return FormulaResult.Error("#VALUE!");
        }

        var numberTimes = (int)args[1].NumericValue;

        if (numberTimes < 0)
        {
            return FormulaResult.Error("#VALUE!");
        }

        if (numberTimes == 0)
        {
            return FormulaResult.FromString(string.Empty);
        }

        // Use StringBuilder for efficient string concatenation
        var sb = new StringBuilder(text.Length * numberTimes);
        for (int i = 0; i < numberTimes; i++)
        {
            sb.Append(text);
        }

        return FormulaResult.FromString(sb.ToString());
    }
}
