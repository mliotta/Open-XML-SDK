// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the OR function.
/// OR(logical1, [logical2], ...) - TRUE if any argument is true.
/// </summary>
public sealed class OrFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly OrFunction Instance = new();

    private OrFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "OR";

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

            if (isTrue)
            {
                return CellValue.FromBool(true);
            }
        }

        return CellValue.FromBool(false);
    }
}
