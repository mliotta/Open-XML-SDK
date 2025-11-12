// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the IMPOWER function.
/// IMPOWER(inumber, number) - raises a complex number to an integer power.
/// </summary>
public sealed class ImPowerFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly ImPowerFunction Instance = new();

    private ImPowerFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "IMPOWER";

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

        var inumber = args[0].StringValue;
        if (!ComplexNumber.TryParse(inumber, out var complex))
        {
            return CellValue.Error("#NUM!");
        }

        if (args[1].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }

        var power = (int)args[1].NumericValue;
        var result = ComplexNumber.Power(complex, power);

        var suffix = inumber.EndsWith("j") ? "j" : "i";
        return CellValue.FromString(result.ToString(suffix));
    }
}
