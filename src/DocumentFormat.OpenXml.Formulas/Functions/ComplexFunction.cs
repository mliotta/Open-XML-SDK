// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the COMPLEX function.
/// COMPLEX(real_num, i_num, [suffix]) - creates complex number from real and imaginary parts.
/// </summary>
public sealed class ComplexFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly ComplexFunction Instance = new();

    private ComplexFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "COMPLEX";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 2 || args.Length > 3)
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

        if (args[0].Type != CellValueType.Number || args[1].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }

        var suffix = "i";
        if (args.Length == 3)
        {
            if (args[2].IsError)
            {
                return args[2];
            }

            suffix = args[2].StringValue?.Trim().ToLowerInvariant();
            if (suffix != "i" && suffix != "j")
            {
                return CellValue.Error("#VALUE!");
            }
        }

        var complex = new ComplexNumber(args[0].NumericValue, args[1].NumericValue);
        return CellValue.FromString(complex.ToString(suffix));
    }
}
