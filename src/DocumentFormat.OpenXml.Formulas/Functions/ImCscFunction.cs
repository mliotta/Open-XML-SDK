// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the IMCSC function.
/// IMCSC(inumber) - returns the cosecant of a complex number.
/// </summary>
public sealed class ImCscFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly ImCscFunction Instance = new();

    private ImCscFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "IMCSC";

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

        var result = ComplexNumber.Csc(complex);
        var suffix = inumber.EndsWith("j") ? "j" : "i";
        return CellValue.FromString(result.ToString(suffix));
    }
}
