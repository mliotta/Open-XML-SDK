// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the MINUTE function.
/// MINUTE(time) - extracts minute (0-59).
/// </summary>
public sealed class MinuteFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly MinuteFunction Instance = new();

    private MinuteFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "MINUTE";

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

        if (args[0].Type != FormulaResultType.Number)
        {
            return FormulaResult.Error("#VALUE!");
        }

        try
        {
            var dateTime = DateTime.FromOADate(args[0].NumericValue);
            return FormulaResult.FromNumber(dateTime.Minute);
        }
        catch
        {
            return FormulaResult.Error("#NUM!");
        }
    }
}
