// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the IMARGUMENT function.
/// IMARGUMENT(inumber) - returns the argument (angle) in radians of a complex number.
/// </summary>
public sealed class ImArgumentFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly ImArgumentFunction Instance = new();

    private ImArgumentFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "IMARGUMENT";

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

        return CellValue.FromNumber(complex.Argument());
    }
}
