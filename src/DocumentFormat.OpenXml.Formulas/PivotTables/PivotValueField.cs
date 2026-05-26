// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

namespace DocumentFormat.OpenXml.Features.PivotTables;

/// <summary>A configured value (data) field.</summary>
internal sealed class PivotValueField
{
    /// <summary>Initializes a new instance of the <see cref="PivotValueField"/> class.</summary>
    /// <param name="fieldName">The source field name.</param>
    /// <param name="aggregate">The aggregate to apply.</param>
    /// <param name="displayName">The caption.</param>
    /// <param name="showAs">How the value is displayed.</param>
    public PivotValueField(string fieldName, PivotAggregate aggregate, string displayName, PivotShowAs showAs)
    {
        FieldName = fieldName;
        Aggregate = aggregate;
        DisplayName = displayName;
        ShowAs = showAs;
    }

    /// <summary>Gets the source field name.</summary>
    public string FieldName { get; }

    /// <summary>Gets the aggregate to apply.</summary>
    public PivotAggregate Aggregate { get; }

    /// <summary>Gets the caption.</summary>
    public string DisplayName { get; }

    /// <summary>Gets how the value is displayed.</summary>
    public PivotShowAs ShowAs { get; }
}
