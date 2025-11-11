// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the SWITCH function.
/// SWITCH(expression, value1, result1, [value2, result2], ..., [default]) - Evaluates an expression against a list of values and returns the corresponding result.
/// If the argument count is odd, the last argument is treated as the default value.
/// </summary>
public sealed class SwitchFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly SwitchFunction Instance = new();

    private SwitchFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "SWITCH";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        // SWITCH requires at least 3 arguments (expression, value1, result1)
        if (args.Length < 3)
        {
            return CellValue.Error("#VALUE!");
        }

        var expression = args[0];

        // Propagate errors from expression
        if (expression.IsError)
        {
            return expression;
        }

        // Determine if there's a default value (odd argument count)
        bool hasDefault = args.Length % 2 == 0;
        int pairCount = hasDefault ? (args.Length - 2) / 2 : (args.Length - 1) / 2;

        // Evaluate value/result pairs
        for (int i = 0; i < pairCount; i++)
        {
            var value = args[1 + (i * 2)];
            var result = args[2 + (i * 2)];

            // Propagate errors from values
            if (value.IsError)
            {
                return value;
            }

            // Propagate errors from results
            if (result.IsError)
            {
                return result;
            }

            // Compare expression with value
            if (ValuesMatch(expression, value))
            {
                return result;
            }
        }

        // No match found, return default if available
        if (hasDefault)
        {
            var defaultValue = args[args.Length - 1];
            if (defaultValue.IsError)
            {
                return defaultValue;
            }

            return defaultValue;
        }

        // No match and no default, return #N/A
        return CellValue.Error("#N/A");
    }

    private static bool ValuesMatch(CellValue expr, CellValue value)
    {
        // Type must match
        if (expr.Type != value.Type)
        {
            return false;
        }

        return expr.Type switch
        {
            CellValueType.Number => System.Math.Abs(expr.NumericValue - value.NumericValue) < 1e-10,
            CellValueType.Text => string.Equals(expr.StringValue, value.StringValue, StringComparison.OrdinalIgnoreCase),
            CellValueType.Boolean => expr.BoolValue == value.BoolValue,
            CellValueType.Empty => true,
            _ => false,
        };
    }
}
