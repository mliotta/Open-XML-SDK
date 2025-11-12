// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Globalization;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the NUMBERVALUE function.
/// NUMBERVALUE(text, [decimal_separator], [group_separator]) - converts text to number with custom separators.
/// </summary>
public sealed class NumberValueFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly NumberValueFunction Instance = new();

    private NumberValueFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "NUMBERVALUE";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 1 || args.Length > 3)
        {
            return CellValue.Error("#VALUE!");
        }

        if (args[0].IsError)
        {
            return args[0];
        }

        // Get text
        var text = args[0].StringValue;
        if (string.IsNullOrEmpty(text) || text.Trim().Length == 0)
        {
            return CellValue.Error("#VALUE!");
        }

        // Get decimal separator (default ".")
        string decimalSeparator = ".";
        if (args.Length >= 2)
        {
            if (args[1].IsError)
            {
                return args[1];
            }

            decimalSeparator = args[1].StringValue;
            if (string.IsNullOrEmpty(decimalSeparator))
            {
                decimalSeparator = ".";
            }
        }

        // Get group separator (default ",")
        string groupSeparator = ",";
        if (args.Length >= 3)
        {
            if (args[2].IsError)
            {
                return args[2];
            }

            groupSeparator = args[2].StringValue;
            if (string.IsNullOrEmpty(groupSeparator))
            {
                groupSeparator = ",";
            }
        }

        // Validate that separators are different
        if (decimalSeparator == groupSeparator && !string.IsNullOrEmpty(groupSeparator))
        {
            return CellValue.Error("#VALUE!");
        }

        // Parse the number
        try
        {
            // Remove group separators
            var cleanedText = text.Replace(groupSeparator, string.Empty);

            // Replace decimal separator with standard decimal point
            if (decimalSeparator != ".")
            {
                cleanedText = cleanedText.Replace(decimalSeparator, ".");
            }

            // Remove whitespace
            cleanedText = cleanedText.Trim();

            // Handle percentage
            bool isPercentage = cleanedText.EndsWith("%");
            if (isPercentage)
            {
                cleanedText = cleanedText.Substring(0, cleanedText.Length - 1).Trim();
            }

            // Parse the number
            if (double.TryParse(cleanedText, NumberStyles.Float | NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out var number))
            {
                // If it was a percentage, divide by 100
                if (isPercentage)
                {
                    number /= 100;
                }

                return CellValue.FromNumber(number);
            }

            return CellValue.Error("#VALUE!");
        }
        catch
        {
            return CellValue.Error("#VALUE!");
        }
    }
}
