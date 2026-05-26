// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

namespace DocumentFormat.OpenXml.Features.PivotTables;

/// <summary>A report-filter (page) field and its optional selected value.</summary>
internal sealed class PivotFilter
{
    /// <summary>Initializes a new instance of the <see cref="PivotFilter"/> class.</summary>
    /// <param name="fieldName">The source field name.</param>
    /// <param name="selectedValue">The selected value, or null for "(All)".</param>
    public PivotFilter(string fieldName, string? selectedValue)
    {
        FieldName = fieldName;
        SelectedValue = selectedValue;
    }

    /// <summary>Gets the source field name.</summary>
    public string FieldName { get; }

    /// <summary>Gets the selected value, or null when all values are included.</summary>
    public string? SelectedValue { get; }
}
