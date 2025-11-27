// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the IF function.
/// </summary>
public sealed class IfFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly IfFunction Instance = new();

    private IfFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "IF";

    /// <inheritdoc/>
    // TODO: Phase 0 limitation - eagerly evaluates both branches.
    // Excel's IF is lazy (only evaluates the taken branch).
    // Phase 1 should make IF a special form with conditional compilation.
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        if (args.Length < 3)
        {
            return FormulaResult.Error("#VALUE!");
        }

        var condition = args[0];

        if (condition.IsError)
        {
            return condition; // Propagate errors
        }

        // Evaluate condition
        var isTrue = condition.Type switch
        {
            FormulaResultType.Boolean => condition.BoolValue,
            FormulaResultType.Number => condition.NumericValue != 0,
            FormulaResultType.Text => !string.IsNullOrEmpty(condition.StringValue),
            _ => false,
        };

        return isTrue ? args[1] : args[2];
    }
}
