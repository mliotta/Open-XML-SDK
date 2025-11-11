// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the PRODUCT function.
/// PRODUCT(number1, [number2], ...) - multiplies all numbers.
/// </summary>
public sealed class ProductFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly ProductFunction Instance = new();

    private ProductFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "PRODUCT";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        var product = 1.0;
        var hasValue = false;

        foreach (var arg in args)
        {
            if (arg.IsError)
            {
                return arg; // Propagate errors
            }

            if (arg.Type == CellValueType.Number)
            {
                product *= arg.NumericValue;
                hasValue = true;
            }
        }

        if (!hasValue)
        {
            return CellValue.FromNumber(0);
        }

        return CellValue.FromNumber(product);
    }
}
