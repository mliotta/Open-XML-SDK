// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the ISOMITTED function (Excel 365).
/// ISOMITTED(argument) - Checks whether the value in a LAMBDA is missing and returns TRUE or FALSE.
/// </summary>
/// <remarks>
/// This function is primarily used with LAMBDA functions to check if an optional parameter was omitted.
/// In the current implementation without full LAMBDA support, we check if the value is empty.
/// A complete implementation would require tracking whether a parameter was explicitly omitted
/// versus being provided as an empty value.
/// </remarks>
public sealed class IsOmittedFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly IsOmittedFunction Instance = new();

    private IsOmittedFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "ISOMITTED";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 1)
        {
            return CellValue.Error("#VALUE!");
        }

        // Note: Errors are NOT propagated for IS* functions
        // In a full LAMBDA implementation, this would check if the parameter was omitted
        // For now, we check if the value is empty
        var isOmitted = args[0].Type == CellValueType.Empty;
        return CellValue.FromBool(isOmitted);
    }
}
