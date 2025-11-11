// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Collections.Generic;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.DependencyGraph;

/// <summary>
/// Represents a dependency graph for formula cells in a worksheet.
/// </summary>
public interface IDependencyGraph
{
    /// <summary>
    /// Adds a cell to the dependency graph.
    /// </summary>
    /// <param name="cellReference">The cell reference (e.g., "A1").</param>
    /// <param name="dependencies">Set of cells this cell depends on.</param>
    void AddCell(string cellReference, HashSet<string> dependencies);

    /// <summary>
    /// Gets the direct dependencies of a cell.
    /// </summary>
    /// <param name="cellReference">The cell reference.</param>
    /// <returns>Set of cells that this cell depends on.</returns>
    HashSet<string> GetDependencies(string cellReference);

    /// <summary>
    /// Gets the direct dependents of a cell.
    /// </summary>
    /// <param name="cellReference">The cell reference.</param>
    /// <returns>Set of cells that depend on this cell.</returns>
    HashSet<string> GetDependents(string cellReference);

    /// <summary>
    /// Gets the evaluation order using topological sort (Kahn's algorithm).
    /// </summary>
    /// <returns>List of cell references in evaluation order.</returns>
    List<string> GetEvaluationOrder();

    /// <summary>
    /// Gets the evaluation order for a subset of cells using topological sort.
    /// </summary>
    /// <param name="cells">The subset of cells to sort.</param>
    /// <returns>List of cell references in evaluation order.</returns>
    List<string> GetEvaluationOrder(IEnumerable<string> cells);

    /// <summary>
    /// Detects circular references in the dependency graph.
    /// </summary>
    /// <returns>List of circular references found.</returns>
    List<CircularReference> DetectCircularReferences();
}

/// <summary>
/// Represents a circular reference chain.
/// </summary>
public class CircularReference
{
    /// <summary>
    /// The chain of cell references forming the cycle.
    /// </summary>
    public List<string> Chain { get; }

    /// <summary>
    /// Initializes a new instance of the CircularReference class.
    /// </summary>
    public CircularReference(List<string> chain)
    {
        Chain = chain;
    }

    /// <summary>
    /// Returns a string representation of the circular reference.
    /// </summary>
    public override string ToString()
    {
        return string.Join(" → ", Chain.ToArray());
    }
}
