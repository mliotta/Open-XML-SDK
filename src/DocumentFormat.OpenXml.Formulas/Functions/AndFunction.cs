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
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length == 0)
        {
            return CellValue.Error("#VALUE!");
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
                CellValueType.Boolean => arg.BoolValue,
                CellValueType.Number => arg.NumericValue != 0,
                CellValueType.Text => !string.IsNullOrEmpty(arg.StringValue),
                CellValueType.Empty => false,
                _ => false,
            };

            if (!isTrue)
            {
                return CellValue.FromBool(false);
            }
        }

        return CellValue.FromBool(true);
    }
}
