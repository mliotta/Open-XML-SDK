// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the IMPRODUCT function.
/// IMPRODUCT(inumber1, [inumber2], ...) - multiplies complex numbers.
/// </summary>
public sealed class ImProductFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly ImProductFunction Instance = new();

    private ImProductFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "IMPRODUCT";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 1)
        {
            return CellValue.Error("#VALUE!");
        }

        var result = new ComplexNumber(1, 0);
        string? suffix = null;

        foreach (var arg in args)
        {
            if (arg.IsError)
            {
                return arg;
            }

            var inumber = arg.StringValue;
            if (!ComplexNumber.TryParse(inumber, out var complex))
            {
                return CellValue.Error("#NUM!");
            }

            result = ComplexNumber.Multiply(result, complex);

            if (suffix == null)
            {
                suffix = inumber.EndsWith("j") ? "j" : "i";
            }
        }

        return CellValue.FromString(result.ToString(suffix ?? "i"));
    }
}
