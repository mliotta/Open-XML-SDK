// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the IMCSCH function.
/// IMCSCH(inumber) - returns the hyperbolic cosecant of a complex number.
/// </summary>
public sealed class ImCschFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly ImCschFunction Instance = new();

    private ImCschFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "IMCSCH";

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

        var result = ComplexNumber.Csch(complex!);

        // Check for division by zero or invalid result
        if (double.IsNaN(result.Real) || double.IsNaN(result.Imaginary) ||
            double.IsInfinity(result.Real) || double.IsInfinity(result.Imaginary))
        {
            return FormulaResult.Error("#NUM!");
        }

        var suffix = inumber.EndsWith("j") ? "j" : "i";
        return FormulaResult.FromString(result.ToString(suffix));
    }
}
