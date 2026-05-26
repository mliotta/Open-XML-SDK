// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

namespace DocumentFormat.OpenXml.Features.PivotTables;

/// <summary>One-based inclusive range bounds.</summary>
internal readonly struct RangeBounds
{
    /// <summary>Initializes a new instance of the <see cref="RangeBounds"/> struct.</summary>
    /// <param name="firstColumn">First column.</param>
    /// <param name="firstRow">First row.</param>
    /// <param name="lastColumn">Last column.</param>
    /// <param name="lastRow">Last row.</param>
    public RangeBounds(int firstColumn, int firstRow, int lastColumn, int lastRow)
    {
        FirstColumn = firstColumn;
        FirstRow = firstRow;
        LastColumn = lastColumn;
        LastRow = lastRow;
    }

    /// <summary>Gets the first column.</summary>
    public int FirstColumn { get; }

    /// <summary>Gets the first row.</summary>
    public int FirstRow { get; }

    /// <summary>Gets the last column.</summary>
    public int LastColumn { get; }

    /// <summary>Gets the last row.</summary>
    public int LastRow { get; }
}
