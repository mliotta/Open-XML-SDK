// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml.Features.FormulaEvaluation;

namespace DocumentFormat.OpenXml.Features.PivotTables;

/// <summary>The source data read from a worksheet range: a header row plus typed data rows.</summary>
internal sealed class PivotSourceData
{
    /// <summary>Initializes a new instance of the <see cref="PivotSourceData"/> class.</summary>
    /// <param name="sheetName">The worksheet name (for the cache source).</param>
    /// <param name="reference">The bounding range reference.</param>
    /// <param name="headers">The column headers.</param>
    /// <param name="rows">The data rows, each aligned to <paramref name="headers"/>.</param>
    public PivotSourceData(string sheetName, string reference, string[] headers, List<FormulaResult[]> rows)
    {
        SheetName = sheetName;
        Reference = reference;
        Headers = headers;
        Rows = rows;
    }

    /// <summary>Gets the worksheet name.</summary>
    public string SheetName { get; }

    /// <summary>Gets the bounding range reference.</summary>
    public string Reference { get; }

    /// <summary>Gets the column headers.</summary>
    public string[] Headers { get; }

    /// <summary>Gets the data rows.</summary>
    public List<FormulaResult[]> Rows { get; }

    /// <summary>Returns the zero-based index of a header, or throws when absent.</summary>
    /// <param name="fieldName">The header name.</param>
    /// <returns>The zero-based field index.</returns>
    public int FieldIndex(string fieldName)
    {
        for (var i = 0; i < Headers.Length; i++)
        {
            if (string.Equals(Headers[i], fieldName, StringComparison.OrdinalIgnoreCase))
            {
                return i;
            }
        }

        throw new ArgumentException($"Field '{fieldName}' was not found in the source headers.", nameof(fieldName));
    }

    /// <summary>Gets the column values for a field across all data rows.</summary>
    /// <param name="fieldIndex">The zero-based field index.</param>
    /// <returns>The column values.</returns>
    public IEnumerable<FormulaResult> Column(int fieldIndex) => Rows.Select(r => r[fieldIndex]);
}
