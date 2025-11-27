// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the IMCOS function.
/// IMCOS(inumber) - returns the cosine of a complex number.
/// </summary>
public sealed class ImCosFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly ImCosFunction Instance = new();

    private ImCosFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "IMCOS";

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

        var result = ComplexNumber.Cos(complex!);
        var suffix = inumber.EndsWith("j") ? "j" : "i";
        return FormulaResult.FromString(result.ToString(suffix));
    }
}
