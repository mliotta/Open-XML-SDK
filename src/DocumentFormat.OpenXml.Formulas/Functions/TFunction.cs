// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the T function.
/// T(value) - returns text if value is text, otherwise empty string.
/// </summary>
public sealed class TFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly TFunction Instance = new();

    private TFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "T";

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

        // If the value is text, return it; otherwise return empty string
        if (args[0].Type == FormulaResultType.Text)
        {
            return args[0];
        }

        return FormulaResult.FromString(string.Empty);
    }
}
