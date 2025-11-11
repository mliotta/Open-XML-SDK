// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Globalization;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the VALUE function.
/// VALUE(text) - converts text to number.
/// </summary>
public sealed class ValueFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly ValueFunction Instance = new();

    private ValueFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "VALUE";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 1)
        {
            return CellValue.Error("#VALUE!");
        }

        if (args[0].IsError)
        {
            return args[0];
        }

        if (args[0].Type == CellValueType.Number)
        {
            return args[0];
        }

        var text = args[0].StringValue;

        if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var number))
        {
            return CellValue.FromNumber(number);
        }

        return CellValue.Error("#VALUE!");
    }
}
