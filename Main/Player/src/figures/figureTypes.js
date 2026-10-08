// The kinds of figure a file may hold, by element name: what DrawingDeserializer.FigureTypes
// finds by reflection over the assembly. Each figure's file registers its class under the
// name the C# class has, which is what the file's element is called.

const FigureTypes = {
    types: new Map(),

    register(name, type) {
        type.typeName = name;
        FigureTypes.types.set(name, type);
    },

    find(typeName) {
        return FigureTypes.types.get(typeName) ?? null;
    },

    /** The element names, for the parity check and the dump */
    names() {
        return [...FigureTypes.types.keys()];
    }
};
