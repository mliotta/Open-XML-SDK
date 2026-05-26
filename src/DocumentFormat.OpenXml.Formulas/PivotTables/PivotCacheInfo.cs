// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Collections.Generic;

using DocumentFormat.OpenXml.Packaging;

namespace DocumentFormat.OpenXml.Features.PivotTables;

/// <summary>The cache artifacts produced for a pivot table.</summary>
internal sealed class PivotCacheInfo
{
    /// <summary>Initializes a new instance of the <see cref="PivotCacheInfo"/> class.</summary>
    /// <param name="cacheDefinitionPart">The created cache definition part.</param>
    /// <param name="cacheId">The workbook-level cache id.</param>
    /// <param name="fieldMembers">Ordered shared members per field (null for non-enumerated fields).</param>
    public PivotCacheInfo(PivotTableCacheDefinitionPart cacheDefinitionPart, uint cacheId, IList<string>?[] fieldMembers)
    {
        CacheDefinitionPart = cacheDefinitionPart;
        CacheId = cacheId;
        FieldMembers = fieldMembers;
    }

    /// <summary>Gets the cache definition part.</summary>
    public PivotTableCacheDefinitionPart CacheDefinitionPart { get; }

    /// <summary>Gets the workbook-level cache id.</summary>
    public uint CacheId { get; }

    /// <summary>Gets the ordered shared members per source field.</summary>
    public IList<string>?[] FieldMembers { get; }
}
