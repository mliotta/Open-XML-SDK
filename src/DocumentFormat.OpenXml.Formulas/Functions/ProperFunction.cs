// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Globalization;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the PROPER function.
/// PROPER(text) - capitalizes first letter of each word.
/// </summary>
public sealed class ProperFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly ProperFunction Instance = new();

    private ProperFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "PROPER";

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

        var text = args[0].StringValue;
        var textInfo = CultureInfo.InvariantCulture.TextInfo;
        var result = textInfo.ToTitleCase(text.ToLowerInvariant());

        return CellValue.FromString(result);
    }
}
