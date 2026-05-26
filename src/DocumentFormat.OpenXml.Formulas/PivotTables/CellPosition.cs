// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

namespace DocumentFormat.OpenXml.Features.PivotTables;

/// <summary>A one-based cell position.</summary>
internal readonly struct CellPosition
{
    /// <summary>Initializes a new instance of the <see cref="CellPosition"/> struct.</summary>
    /// <param name="column">The one-based column.</param>
    /// <param name="row">The one-based row.</param>
    public CellPosition(int column, int row)
    {
        Column = column;
        Row = row;
    }

    /// <summary>Gets the one-based column.</summary>
    public int Column { get; }

    /// <summary>Gets the one-based row.</summary>
    public int Row { get; }
}
