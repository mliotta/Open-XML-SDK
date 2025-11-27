// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the TYPE function.
/// TYPE(value) - returns type code (1=number, 2=text, 4=boolean, 16=error, 64=array).
/// </summary>
public sealed class TypeFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly TypeFunction Instance = new();

    private TypeFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "TYPE";

    /// <inheritdoc/>
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        if (args.Length != 1)
        {
            return FormulaResult.Error("#VALUE!");
        }

        // TYPE does not propagate errors, it returns 16 for error values
        var typeCode = args[0].Type switch
        {
            FormulaResultType.Number => 1,
            FormulaResultType.Text => 2,
            FormulaResultType.Boolean => 4,
            FormulaResultType.Error => 16,
            FormulaResultType.Empty => 1, // Empty cells are treated as numeric 0
            _ => 1, // Default to number
        };

        return FormulaResult.FromNumber(typeCode);
    }
}
