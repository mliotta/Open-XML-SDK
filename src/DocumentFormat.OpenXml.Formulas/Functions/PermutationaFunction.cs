// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the PERMUTATIONA function.
/// PERMUTATIONA(number, number_chosen) - returns the number of permutations for a given number of objects (with repetitions).
/// Formula: number^number_chosen
/// </summary>
public sealed class PermutationaFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly PermutationaFunction Instance = new();

    private PermutationaFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "PERMUTATIONA";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 2)
        {
            return CellValue.Error("#VALUE!");
        }

        foreach (var arg in args)
        {
            if (arg.IsError)
            {
                return arg;
            }
        }

        if (args[0].Type != CellValueType.Number || args[1].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }

        int number = (int)args[0].NumericValue;
        int numberChosen = (int)args[1].NumericValue;

        if (number < 0 || numberChosen < 0)
        {
            return CellValue.Error("#NUM!");
        }

        // Special cases
        if (number == 0 && numberChosen > 0)
        {
            return CellValue.FromNumber(0);
        }

        if (numberChosen == 0)
        {
            return CellValue.FromNumber(1);
        }

        // PERMUTATIONA = number^number_chosen
        double result = System.Math.Pow(number, numberChosen);

        // Check for overflow
        if (double.IsInfinity(result))
        {
            return CellValue.Error("#NUM!");
        }

        return CellValue.FromNumber(result);
    }
}
