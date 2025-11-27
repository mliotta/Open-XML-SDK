// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the HYPERLINK function.
/// HYPERLINK(link_location, [friendly_name]) - Creates a clickable hyperlink.
/// </summary>
public sealed class HyperlinkFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly HyperlinkFunction Instance = new();

    private HyperlinkFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "HYPERLINK";

    /// <inheritdoc/>
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        if (args.Length < 1)
        {
            return FormulaResult.Error("#VALUE!");
        }

        var linkLocation = args[0];

        // Check for errors
        if (linkLocation.IsError)
        {
            return linkLocation;
        }

        // Get friendly name if provided, otherwise use the link location
        var friendlyName = args.Length >= 2 ? args[1] : linkLocation;

        // Check for errors in friendly name
        if (friendlyName.IsError)
        {
            return friendlyName;
        }

        // Return the friendly name as the display value
        // Note: The actual hyperlink functionality requires cell formatting/metadata
        // which is beyond the scope of formula evaluation
        if (friendlyName.Type == FormulaResultType.Text)
        {
            return FormulaResult.FromString(friendlyName.StringValue);
        }
        else if (friendlyName.Type == FormulaResultType.Number)
        {
            return FormulaResult.FromNumber(friendlyName.NumericValue);
        }
        else if (linkLocation.Type == FormulaResultType.Text)
        {
            return FormulaResult.FromString(linkLocation.StringValue);
        }

        return FormulaResult.Error("#VALUE!");
    }
}
