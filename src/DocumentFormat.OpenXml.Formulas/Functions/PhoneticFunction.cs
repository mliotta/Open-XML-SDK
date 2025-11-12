// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the PHONETIC function.
/// PHONETIC(reference) - returns phonetic text.
/// Phase 0 implementation: returns the text value (full phonetic guide text requires complex processing).
/// </summary>
public sealed class PhoneticFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly PhoneticFunction Instance = new();

    private PhoneticFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "PHONETIC";

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

        // Phase 0 implementation: return the text value as-is
        // Full implementation would extract phonetic guide text which is stored separately in Excel
        var text = args[0].StringValue;

        return CellValue.FromString(text);
    }
}
