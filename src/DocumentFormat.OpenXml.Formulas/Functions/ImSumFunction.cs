// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the IMSUM function.
/// IMSUM(inumber1, [inumber2], ...) - adds complex numbers.
/// </summary>
public sealed class ImSumFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly ImSumFunction Instance = new();

    private ImSumFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "IMSUM";

    /// <inheritdoc/>
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        if (args.Length < 1)
        {
            return FormulaResult.Error("#VALUE!");
        }

        var result = new ComplexNumber(0, 0);
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
                return FormulaResult.Error("#NUM!");
            }

            result = ComplexNumber.Add(result, complex!);

            if (suffix == null)
            {
                suffix = inumber.EndsWith("j") ? "j" : "i";
            }
        }

        return FormulaResult.FromString(result.ToString(suffix ?? "i"));
    }
}
