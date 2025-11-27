// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Text;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the ENCODEURL function.
/// ENCODEURL(text) - URL-encodes a string.
/// </summary>
public sealed class EncodeUrlFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly EncodeUrlFunction Instance = new();

    private EncodeUrlFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "ENCODEURL";

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

        if (args[0].Type == FormulaResultType.Empty)
        {
            return FormulaResult.FromString(string.Empty);
        }

        var text = args[0].StringValue;

        // URL encode using Uri.EscapeDataString which is RFC 3986 compliant
        var encoded = Uri.EscapeDataString(text);

        return FormulaResult.FromString(encoded);
    }
}
