// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the ISREF function.
/// ISREF(value) - TRUE if value is a reference.
/// </summary>
/// <remarks>
/// In the current implementation, we don't have a distinct Reference type in FormulaResult.
/// This function returns FALSE for all values since references are typically resolved
/// before function execution. A full implementation would require tracking whether
/// a value originated from a cell reference.
/// </remarks>
public sealed class IsRefFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly IsRefFunction Instance = new();

    private IsRefFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "ISREF";

    /// <inheritdoc/>
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        if (args.Length != 1)
        {
            return FormulaResult.Error("#VALUE!");
        }

        // Note: Errors are NOT propagated for IS* functions
        // In Excel, ISREF returns TRUE for cell references, named ranges, etc.
        // Since FormulaResult doesn't currently have a Reference type, we return FALSE
        // This would need to be enhanced if the system tracks reference types
        return FormulaResult.FromBool(false);
    }
}
