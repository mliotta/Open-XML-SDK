// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the AND function.
/// AND(logical1, [logical2], ...) - TRUE if all arguments are true.
/// </summary>
public sealed class AndFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly AndFunction Instance = new();

    private AndFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "AND";

    /// <inheritdoc/>
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        if (args.Length == 0)
        {
            return FormulaResult.Error("#VALUE!");
        }

        foreach (var arg in args)
        {
            if (arg.IsError)
            {
                return arg; // Propagate errors
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

            if (!isTrue)
            {
                return FormulaResult.FromBool(false);
            }
        }

        return FormulaResult.FromBool(true);
    }
}
