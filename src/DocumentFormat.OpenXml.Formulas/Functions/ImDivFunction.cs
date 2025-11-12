// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the IMDIV function.
/// IMDIV(inumber1, inumber2) - divides two complex numbers.
/// </summary>
public sealed class ImDivFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly ImDivFunction Instance = new();

    private ImDivFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "IMDIV";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 2)
        {
            return CellValue.Error("#VALUE!");
        }

        if (args[0].IsError)
        {
            return args[0];
        }

        if (args[1].IsError)
        {
            return args[1];
        }

        var inumber1 = args[0].StringValue;
        var inumber2 = args[1].StringValue;

        if (!ComplexNumber.TryParse(inumber1, out var complex1))
        {
            return CellValue.Error("#NUM!");
        }

        if (!ComplexNumber.TryParse(inumber2, out var complex2))
        {
            return CellValue.Error("#NUM!");
        }

        var result = ComplexNumber.Divide(complex1, complex2);

        if (double.IsNaN(result.Real) || double.IsNaN(result.Imaginary))
        {
            return CellValue.Error("#NUM!");
        }

        var suffix = inumber1.EndsWith("j") ? "j" : "i";
        return CellValue.FromString(result.ToString(suffix));
    }
}
