// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the ARABIC function.
/// ARABIC(text) - converts a Roman numeral to an Arabic numeral (number).
/// </summary>
public sealed class ArabicFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly ArabicFunction Instance = new();

    private static readonly Dictionary<char, int> _romanValues = new()
    {
        { 'I', 1 },
        { 'V', 5 },
        { 'X', 10 },
        { 'L', 50 },
        { 'C', 100 },
        { 'D', 500 },
        { 'M', 1000 }
    };

    private ArabicFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "ARABIC";

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

        string text;
        if (args[0].Type == FormulaResultType.Text)
        {
            text = args[0].StringValue?.Trim().ToUpperInvariant() ?? string.Empty;
        }
        else if (args[0].Type == FormulaResultType.Number)
        {
            // If it's already a number, just return it
            return args[0];
        }
        else
        {
            return FormulaResult.Error("#VALUE!");
        }

        if (string.IsNullOrEmpty(text))
        {
            return FormulaResult.Error("#VALUE!");
        }

        try
        {
            var result = ConvertRomanToArabic(text);
            return FormulaResult.FromNumber(result);
        }
        catch
        {
            return FormulaResult.Error("#VALUE!");
        }
    }

    private static int ConvertRomanToArabic(string roman)
    {
        if (string.IsNullOrEmpty(roman))
        {
            throw new ArgumentException("Invalid Roman numeral");
        }

        int result = 0;
        int previousValue = 0;

        // Process the Roman numeral from right to left
        for (int i = roman.Length - 1; i >= 0; i--)
        {
            char c = roman[i];

            if (!_romanValues.TryGetValue(c, out int currentValue))
            {
                throw new ArgumentException($"Invalid Roman numeral character: {c}");
            }

            // If the current value is less than the previous value, subtract it
            // (e.g., in "IV", I comes before V, so we subtract I)
            if (currentValue < previousValue)
            {
                result -= currentValue;
            }
            else
            {
                result += currentValue;
            }

            previousValue = currentValue;
        }

        // Validate the result is positive
        if (result <= 0)
        {
            throw new ArgumentException("Invalid Roman numeral");
        }

        return result;
    }
}
