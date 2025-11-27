// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the N function.
/// N(value) - converts value to number (numbers unchanged, TRUE=1, FALSE=0, dates to serial, text to 0).
/// </summary>
public sealed class NFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly NFunction Instance = new();

    private NFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "N";

    /// <inheritdoc/>
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        if (args.Length != 1)
        {
            return FormulaResult.Error("#VALUE!");
        }

        // Propagate errors
        if (args[0].IsError)
        {
            return args[0];
        }

        // Convert to number based on type
        var result = args[0].Type switch
        {
            FormulaResultType.Number => args[0].NumericValue,
            FormulaResultType.Boolean => args[0].BoolValue ? 1.0 : 0.0,
            FormulaResultType.Text => 0.0, // Text converts to 0
            FormulaResultType.Empty => 0.0, // Empty converts to 0
            FormulaResultType.Error => 0.0, // Should not reach here due to error propagation
            _ => 0.0,
        };

        return FormulaResult.FromNumber(result);
    }
}
