// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

namespace DocumentFormat.OpenXml.Features.PivotTables;

/// <summary>
/// Controls how row and column labels are arranged in the generated pivot table.
/// </summary>
public enum PivotLayout
{
    /// <summary>Compact form: nested fields share a single column (Excel default).</summary>
    Compact,

    /// <summary>Outline form: each field gets its own column, labels stacked at the top.</summary>
    Outline,

    /// <summary>Tabular form: each field gets its own column, labels in classic table rows.</summary>
    Tabular,
}
