// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the IMSUB function.
/// IMSUB(inumber1, inumber2) - subtracts two complex numbers.
/// </summary>
public sealed class ImSubFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly ImSubFunction Instance = new();

    private ImSubFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "IMSUB";

    /// <inheritdoc/>
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        if (args.Length != 2)
        {
            return FormulaResult.Error("#VALUE!");
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
            return FormulaResult.Error("#NUM!");
        }

        if (!ComplexNumber.TryParse(inumber2, out var complex2))
        {
            return FormulaResult.Error("#NUM!");
        }

        var result = ComplexNumber.Subtract(complex1!, complex2!);
        var suffix = inumber1.EndsWith("j") ? "j" : "i";
        return FormulaResult.FromString(result.ToString(suffix));
    }
}
