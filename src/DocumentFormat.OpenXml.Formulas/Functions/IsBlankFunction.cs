// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the ISBLANK function.
/// ISBLANK(value) - TRUE if value is blank/empty.
/// </summary>
public sealed class IsBlankFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly IsBlankFunction Instance = new();

    private IsBlankFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "ISBLANK";

    /// <inheritdoc/>
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        if (args.Length != 1)
        {
            return FormulaResult.Error("#VALUE!");
        }

        // Note: Errors are NOT propagated for IS* functions
        var isBlank = args[0].Type == FormulaResultType.Empty;
        return FormulaResult.FromBool(isBlank);
    }
}
