// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the SUBSTITUTE function.
/// SUBSTITUTE(text, old_text, new_text, [instance_num]) - replaces text.
/// </summary>
public sealed class SubstituteFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly SubstituteFunction Instance = new();

    private SubstituteFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "SUBSTITUTE";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 3 || args.Length > 4)
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

        if (args[2].IsError)
        {
            return args[2];
        }

        var text = args[0].StringValue;
        var oldText = args[1].StringValue;
        var newText = args[2].StringValue;

        if (args.Length == 4)
        {
            if (args[3].IsError)
            {
                return args[3];
            }

            if (args[3].Type != CellValueType.Number)
            {
                return CellValue.Error("#VALUE!");
            }

            var instanceNum = (int)args[3].NumericValue;

            if (instanceNum < 1)
            {
                return CellValue.Error("#VALUE!");
            }

            // Replace only the specified instance
            var count = 0;
            var index = 0;

            while ((index = text.IndexOf(oldText, index)) != -1)
            {
                count++;
                if (count == instanceNum)
                {
                    text = text.Substring(0, index) + newText + text.Substring(index + oldText.Length);
                    break;
                }

                index += oldText.Length;
            }

            return CellValue.FromString(text);
        }

        // Replace all instances
        var result = text.Replace(oldText, newText);
        return CellValue.FromString(result);
    }
}
