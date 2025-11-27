// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the ISTEXT function.
/// ISTEXT(value) - TRUE if value is text.
/// </summary>
public sealed class IsTextFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly IsTextFunction Instance = new();

    private IsTextFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "ISTEXT";

    /// <inheritdoc/>
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        if (args.Length != 1)
        {
            return FormulaResult.Error("#VALUE!");
        }

        // Note: Errors are NOT propagated for IS* functions
        var isText = args[0].Type == FormulaResultType.Text;
        return FormulaResult.FromBool(isText);
    }
}
