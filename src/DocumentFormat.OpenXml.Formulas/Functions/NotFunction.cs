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
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 1)
        {
            return CellValue.Error("#VALUE!");
        }

        if (args[0].IsError)
        {
            return args[0]; // Propagate errors
        }

        // Evaluate as boolean
        var isTrue = args[0].Type switch
        {
            CellValueType.Boolean => args[0].BoolValue,
            CellValueType.Number => args[0].NumericValue != 0,
            CellValueType.Text => !string.IsNullOrEmpty(args[0].StringValue),
            CellValueType.Empty => false,
            _ => false,
        };

        return CellValue.FromBool(!isTrue);
    }
}
