// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

namespace DocumentFormat.OpenXml.Features.PivotTables;

/// <summary>
/// How a value field's results are displayed. Percent variants render fractions (e.g. 0.5 for 50%)
/// and set the data field's number format to a percentage.
/// </summary>
public enum PivotShowAs
{
    /// <summary>The raw aggregated value.</summary>
    Normal,

    /// <summary>Each value as a fraction of the grand total.</summary>
    PercentOfTotal,

    /// <summary>Each value as a fraction of its column total.</summary>
    PercentOfColumn,

    /// <summary>Each value as a fraction of its row total.</summary>
    PercentOfRow,
}
