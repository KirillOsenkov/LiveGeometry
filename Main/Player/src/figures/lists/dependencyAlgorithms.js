// Port of Main/Avalonia/DynamicGeometry/Figures/Lists/DependencyAlgorithms.cs

const DependencyAlgorithms = {
    /** A depth-first post-order over the children: every node after all it leads to */
    topologicalSort(originalSet, childrenSelector) {
        const visitedSet = new Set();
        const finalOrder = [];
        DependencyAlgorithms.addAllDependents(originalSet, childrenSelector, visitedSet, node => finalOrder.push(node));
        return finalOrder;
    },

    addAllDependents(nodes, childrenSelector, visitedSet, resultCollector) {
        for (const node of nodes) {
            if (!visitedSet.has(node)) {
                visitedSet.add(node);
                const children = childrenSelector(node);
                if (children != null && children.length > 0) {
                    DependencyAlgorithms.addAllDependents(children, childrenSelector, visitedSet, resultCollector);
                }

                if (resultCollector != null) {
                    resultCollector(node);
                }
            }
        }
    },

    /** The nodes and all they lead to, sinks first: reversed, a topological order */
    findDescendants(childrenSelector, list) {
        return DependencyAlgorithms.topologicalSort(list, childrenSelector);
    },

    /** The nodes with no children, reached from the list, each once */
    findRoots(childrenSelector, list) {
        const result = [];
        for (const figure of list) {
            DependencyAlgorithms.findRootsOf(childrenSelector, figure, root => result.push(root));
        }

        return [...new Set(result)];
    },

    /** Doesn't use recursion because deep figures could overflow the stack */
    findRootsOf(childrenSelector, figure, collector) {
        const stack = [figure];
        while (stack.length > 0) {
            if (stack.length > 10000) {
                throw new Error("Weird, we hit a cycle in a DAG, need to investigate this bug");
            }

            figure = stack.pop();
            const children = childrenSelector(figure);
            if (children == null || children.length === 0) {
                collector(figure);
            } else {
                for (let i = children.length - 1; i >= 0; i--) {
                    stack.push(children[i]);
                }
            }
        }
    },

    figureCompletelyDependsOnFigures(figure, set) {
        if (set.includes(figure)) {
            return true;
        }

        if (figure.dependencies.length === 0) {
            return false;
        }

        for (const dependency of figure.dependencies) {
            if (!DependencyAlgorithms.figureCompletelyDependsOnFigures(dependency, set)) {
                return false;
            }
        }

        return true;
    },

    /** The figures between the source and the sink that the sink is built on through the source, in dependency order */
    findImpactedDependencyChain(source, sink) {
        const visitedSet = new Set();
        DependencyAlgorithms.addAllDependents([source], f => f.dependents, visitedSet, null);
        const result = [];
        DependencyAlgorithms.addImpactedDependency(sink, visitedSet, result);
        return result;
    },

    addImpactedDependency(sink, candidates, result) {
        if (!candidates.has(sink)) {
            return;
        }

        for (const dependency of sink.dependencies) {
            DependencyAlgorithms.addImpactedDependency(dependency, candidates, result);
        }

        result.push(sink);
        candidates.delete(sink);
    }
};
