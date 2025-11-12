// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Globalization;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the DATEVALUE function.
/// DATEVALUE(date_text) - converts text to date serial number.
/// </summary>
public sealed class DateValueFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly DateValueFunction Instance = new();

    private DateValueFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "DATEVALUE";

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

        if (args[0].Type != CellValueType.Text)
        {
            return CellValue.Error("#VALUE!");
        }

        try
        {
            var dateText = args[0].StringValue;

            // Try to parse as DateTime (handles various date formats)
            if (DateTime.TryParse(dateText, CultureInfo.InvariantCulture, DateTimeStyles.None, out var dateTime))
            {
                // Return only the date portion (integer part of OADate)
                var oaDate = dateTime.Date.ToOADate();
                return CellValue.FromNumber(oaDate);
            }

            // If parsing fails, return error
            return CellValue.Error("#VALUE!");
        }
        catch
        {
            return CellValue.Error("#VALUE!");
        }
    }
}
