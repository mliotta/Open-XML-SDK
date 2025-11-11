// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the IFS function.
/// IFS(condition1, value1, [condition2, value2], ...) - Evaluates multiple conditions and returns the value corresponding to the first TRUE condition.
/// Returns #N/A if no conditions are TRUE.
/// </summary>
public sealed class IfsFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly IfsFunction Instance = new();

    private IfsFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "IFS";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        // IFS requires at least 2 arguments (1 condition and 1 value)
        // Arguments must come in pairs (condition, value)
        if (args.Length < 2 || args.Length % 2 != 0)
        {
            return CellValue.Error("#VALUE!");
        }

        // Evaluate condition/value pairs in order
        for (int i = 0; i < args.Length; i += 2)
        {
            var condition = args[i];
            var value = args[i + 1];

            // Propagate errors from conditions
            if (condition.IsError)
            {
                return condition;
            }

            // Propagate errors from values
            if (value.IsError)
            {
                return value;
            }

            // Evaluate condition as boolean
            var isTrue = condition.Type switch
            {
                CellValueType.Boolean => condition.BoolValue,
                CellValueType.Number => condition.NumericValue != 0,
                CellValueType.Text => !string.IsNullOrEmpty(condition.StringValue),
                CellValueType.Empty => false,
                _ => false,
            };

            // Return the value if condition is TRUE
            if (isTrue)
            {
                return value;
            }
        }

        // No conditions matched, return #N/A error
        return CellValue.Error("#N/A");
    }
}
