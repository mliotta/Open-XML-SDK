// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the IMLN function.
/// IMLN(inumber) - returns the natural logarithm of a complex number.
/// </summary>
public sealed class ImLnFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly ImLnFunction Instance = new();

    private ImLnFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "IMLN";

    /// <inheritdoc/>
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        if (args.Length != 1)
        {
            return FormulaResult.Error("#VALUE!");
        }

        if (args[0].IsError)
        {
            return args[0];
        }

        var inumber = args[0].StringValue;
        if (!ComplexNumber.TryParse(inumber, out var complex))
        {
            return FormulaResult.Error("#NUM!");
        }

        var result = ComplexNumber.Ln(complex!);
        var suffix = inumber.EndsWith("j") ? "j" : "i";
        return FormulaResult.FromString(result.ToString(suffix));
    }
}
