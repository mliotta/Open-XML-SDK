// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the XOR function.
/// XOR(logical1, [logical2], ...) - Returns TRUE if an odd number of arguments evaluate to TRUE.
/// Returns FALSE if an even number (including zero) of arguments evaluate to TRUE.
/// </summary>
public sealed class XorFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly XorFunction Instance = new();

    private XorFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "XOR";

    /// <inheritdoc/>
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        if (args.Length == 0)
        {
            return FormulaResult.Error("#VALUE!");
        }

        int trueCount = 0;

        foreach (var arg in args)
        {
            // Propagate errors
            if (arg.IsError)
            {
                return arg;
            }

            // Evaluate as boolean
            var isTrue = arg.Type switch
            {
                FormulaResultType.Boolean => arg.BoolValue,
                FormulaResultType.Number => arg.NumericValue != 0,
                FormulaResultType.Text => !string.IsNullOrEmpty(arg.StringValue),
                FormulaResultType.Empty => false,
                _ => false,
            };

            if (isTrue)
            {
                trueCount++;
            }
        }

        // XOR returns TRUE if odd number of TRUE values
        return FormulaResult.FromBool(trueCount % 2 == 1);
    }
}
