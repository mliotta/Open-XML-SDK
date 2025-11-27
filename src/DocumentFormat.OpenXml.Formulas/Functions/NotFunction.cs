// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the NOT function.
/// NOT(logical) - reverses the logical value.
/// </summary>
public sealed class NotFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly NotFunction Instance = new();

    private NotFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "NOT";

    /// <inheritdoc/>
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        if (args.Length != 1)
        {
            return FormulaResult.Error("#VALUE!");
        }

        if (args[0].IsError)
        {
            return args[0]; // Propagate errors
        }

        // Evaluate as boolean
        var isTrue = args[0].Type switch
        {
            FormulaResultType.Boolean => args[0].BoolValue,
            FormulaResultType.Number => args[0].NumericValue != 0,
            FormulaResultType.Text => !string.IsNullOrEmpty(args[0].StringValue),
            FormulaResultType.Empty => false,
            _ => false,
        };

        return FormulaResult.FromBool(!isTrue);
    }
}
