// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the ISEVEN function.
/// ISEVEN(number) - TRUE if number is even.
/// </summary>
public sealed class IsEvenFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly IsEvenFunction Instance = new();

    private IsEvenFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "ISEVEN";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 1)
        {
            return CellValue.Error("#VALUE!");
        }

        // Note: Errors are NOT propagated for IS* functions
        if (args[0].Type != CellValueType.Number)
        {
            return CellValue.FromBool(false);
        }

        var number = args[0].NumericValue;
        var truncated = System.Math.Truncate(number);
        var isEven = truncated % 2 == 0;
        return CellValue.FromBool(isEven);
    }
}
