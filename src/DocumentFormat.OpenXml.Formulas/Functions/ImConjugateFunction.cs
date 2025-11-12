// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the IMCONJUGATE function.
/// IMCONJUGATE(inumber) - returns the complex conjugate of a complex number.
/// </summary>
public sealed class ImConjugateFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly ImConjugateFunction Instance = new();

    private ImConjugateFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "IMCONJUGATE";

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

        var inumber = args[0].StringValue;
        if (!ComplexNumber.TryParse(inumber, out var complex))
        {
            return CellValue.Error("#NUM!");
        }

        var result = complex.Conjugate();
        var suffix = inumber.EndsWith("j") ? "j" : "i";
        return CellValue.FromString(result.ToString(suffix));
    }
}
