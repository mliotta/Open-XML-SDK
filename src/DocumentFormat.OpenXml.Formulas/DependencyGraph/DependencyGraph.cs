// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Linq;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.DependencyGraph;

/// <summary>
/// Implements a dependency graph for formula evaluation order.
/// Uses Kahn's algorithm for topological sorting.
/// </summary>
internal class DependencyGraph : IDependencyGraph
{
    private readonly Dictionary<string, HashSet<string>> _dependencies = new();
    private readonly Dictionary<string, HashSet<string>> _dependents = new();

    public void AddCell(string cellReference, HashSet<string> dependencies)
    {
        _dependencies[cellReference] = dependencies;

        // Build reverse mapping (dependents)
        foreach (var dep in dependencies)
        {
            if (!_dependents.ContainsKey(dep))
            {
                _dependents[dep] = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            }

            _dependents[dep].Add(cellReference);
        }
    }

    public HashSet<string> GetDependencies(string cellReference)
    {
        return _dependencies.TryGetValue(cellReference, out var deps)
            ? deps
            : new HashSet<string>();
    }

    public HashSet<string> GetDependents(string cellReference)
    {
        return _dependents.TryGetValue(cellReference, out var deps)
            ? deps
            : new HashSet<string>();
    }

    public List<string> GetEvaluationOrder()
    {
        // Kahn's algorithm for topological sort
        var inDegree = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        // Initialize in-degrees
        foreach (var cellRef in _dependencies.Keys)
        {
            inDegree[cellRef] = _dependencies[cellRef].Count;
        }

        // Queue cells with no dependencies
        var queue = new Queue<string>(
            inDegree.Where(kvp => kvp.Value == 0).Select(kvp => kvp.Key)
        );

        var result = new List<string>();

        while (queue.Count > 0)
        {
            var current = queue.Dequeue();
            result.Add(current);

            // Reduce in-degree for dependents
            if (_dependents.TryGetValue(current, out var deps))
            {
                foreach (var dep in deps)
                {
                    if (inDegree.ContainsKey(dep))
                    {
                        inDegree[dep]--;
                        if (inDegree[dep] == 0)
                        {
                            queue.Enqueue(dep);
                        }
                    }
                }
            }
        }

        // Check for circular references
        if (result.Count != _dependencies.Count)
        {
            var cycles = DetectCircularReferences();
            if (cycles.Count > 0)
            {
                throw new CircularReferenceException(cycles[0].Chain);
            }
        }

        return result;
    }

    public List<string> GetEvaluationOrder(IEnumerable<string> cells)
    {
        // Kahn's algorithm for topological sort on a subset of cells
        var cellSet = new HashSet<string>(cells, StringComparer.OrdinalIgnoreCase);
        var inDegree = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        // Initialize in-degrees only for cells in the subset
        foreach (var cellRef in cellSet)
        {
            if (_dependencies.ContainsKey(cellRef))
            {
                // Count only dependencies that are also in the subset
                var depsInSubset = _dependencies[cellRef].Count(dep => cellSet.Contains(dep));
                inDegree[cellRef] = depsInSubset;
            }
            else
            {
                // Not a formula cell, treat as having no dependencies
                inDegree[cellRef] = 0;
            }
        }

        // Queue cells with no dependencies within the subset
        var queue = new Queue<string>(
            inDegree.Where(kvp => kvp.Value == 0).Select(kvp => kvp.Key)
        );

        var result = new List<string>();

        while (queue.Count > 0)
        {
            var current = queue.Dequeue();
            result.Add(current);

            // Reduce in-degree for dependents that are in the subset
            if (_dependents.TryGetValue(current, out var deps))
            {
                foreach (var dep in deps)
                {
                    if (inDegree.ContainsKey(dep))
                    {
                        inDegree[dep]--;
                        if (inDegree[dep] == 0)
                        {
                            queue.Enqueue(dep);
                        }
                    }
                }
            }
        }

        // Check for circular references within the subset
        if (result.Count != cellSet.Count)
        {
            var cycles = DetectCircularReferences();
            if (cycles.Count > 0)
            {
                throw new CircularReferenceException(cycles[0].Chain);
            }
        }

        return result;
    }

    public List<CircularReference> DetectCircularReferences()
    {
        var cycles = new List<CircularReference>();
        var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var recursionStack = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var path = new List<string>();

        foreach (var cell in _dependencies.Keys)
        {
            if (!visited.Contains(cell))
            {
                DetectCyclesDFS(cell, visited, recursionStack, path, cycles);
            }
        }

        return cycles;
    }

    private bool DetectCyclesDFS(
        string cell,
        HashSet<string> visited,
        HashSet<string> recursionStack,
        List<string> path,
        List<CircularReference> cycles)
    {
        visited.Add(cell);
        recursionStack.Add(cell);
        path.Add(cell);

        if (_dependencies.TryGetValue(cell, out var deps))
        {
            foreach (var dep in deps)
            {
                if (!visited.Contains(dep))
                {
                    if (DetectCyclesDFS(dep, visited, recursionStack, path, cycles))
                    {
                        return true;
                    }
                }
                else if (recursionStack.Contains(dep))
                {
                    // Found cycle
                    var cycleStart = path.IndexOf(dep);
                    var cycle = path.Skip(cycleStart).ToList();
                    cycle.Add(dep); // Complete the cycle
                    cycles.Add(new CircularReference(cycle));
                    return true;
                }
            }
        }

        path.RemoveAt(path.Count - 1);
        recursionStack.Remove(cell);
        return false;
    }
}
