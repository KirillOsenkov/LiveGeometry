// Port of Main/Avalonia/LiveGeometry/UI/AppTheme.cs: the two themes' colors that drawings
// take (the Paper group and Text). The chrome's colors and the theme editing are the
// app's, not the player's.

class AppTheme {
    constructor(name, colors) {
        this.name = name;
        Object.assign(this, colors);
    }

    static Light = new AppTheme("Light", {
        paper: Color.fromUInt32(0xFFFFFFFF),
        ink: Color.fromUInt32(0xFF000000),
        line: Color.fromUInt32(0x64000000),
        sliderTrack: Color.fromUInt32(0xFFC0C0C0),
        freePointFill: Color.fromUInt32(0xFFFFFF64),
        pointOnFigureFill: Color.fromUInt32(0xFF7CE38B),
        intersectionPointFill: Color.fromUInt32(0xFF6FD3F7),
        midpointFill: Color.fromUInt32(0xFFFFB45A),
        dependentPointFill: Color.fromUInt32(0xFFD0D0D0),
        shapeFill: Color.fromUInt32(0x64FFFFC8),
        axis: Color.fromUInt32(0xFF8080FF),
        gridMajor: Color.fromUInt32(0xFFD3D3D3),
        gridMinor: Color.fromUInt32(0xFFECECEC),
        text: Color.fromUInt32(0xFF2B3038),
        textEmphasis: Color.fromUInt32(0xFF000000),
        expressionFigure: Color.fromUInt32(0xFF267F99),
        expressionProperty: Color.fromUInt32(0xFFA0522D),
        expressionFunction: Color.fromUInt32(0xFF008000),
        expressionNumber: Color.fromUInt32(0xFF8A2BE2),
        expressionReserved: Color.fromUInt32(0xFF808080)
    });

    static Dark = new AppTheme("Dark", {
        paper: Color.fromUInt32(0xFF2B2B2B),
        ink: Color.fromUInt32(0xFFD0D0D0),
        line: Color.fromUInt32(0xFFD3D3D3),
        sliderTrack: Color.fromUInt32(0xFF606060),
        freePointFill: Color.fromUInt32(0xFFF5C542),
        pointOnFigureFill: Color.fromUInt32(0xFF6BCF7F),
        intersectionPointFill: Color.fromUInt32(0xFF4FC3F7),
        midpointFill: Color.fromUInt32(0xFFF0A050),
        dependentPointFill: Color.fromUInt32(0xFF8E949C),
        shapeFill: Color.fromUInt32(0x46D8CC96),
        axis: Color.fromUInt32(0xFF8C8CFF),
        gridMajor: Color.fromUInt32(0xFF4A4A4A),
        gridMinor: Color.fromUInt32(0xFF383838),
        text: Color.fromUInt32(0xFFD9DEE6),
        textEmphasis: Color.fromUInt32(0xFFFFFFFF),
        expressionFigure: Color.fromUInt32(0xFF4EC9B0),
        expressionProperty: Color.fromUInt32(0xFF9CDCFE),
        expressionFunction: Color.fromUInt32(0xFFDCDCAA),
        expressionNumber: Color.fromUInt32(0xFFB5CEA8),
        expressionReserved: Color.fromUInt32(0xFF569CD6)
    });

    /** The Expressions color of a piece of a formula of this kind (ExpressionClass), as the app's expression boxes have it */
    expressionColor(kind) {
        switch (kind) {
            case ExpressionClass.Figure:
                return this.expressionFigure;
            case ExpressionClass.Property:
                return this.expressionProperty;
            case ExpressionClass.Function:
                return this.expressionFunction;
            case ExpressionClass.Number:
                return this.expressionNumber;
            default:
                return this.expressionReserved;
        }
    }

    static get All() {
        return [AppTheme.Light, AppTheme.Dark];
    }

    /** The theme a style's own values are for */
    static get Base() {
        return AppTheme.Light;
    }

    static current = AppTheme.Light;

    static isBase(theme) {
        return theme === AppTheme.Base;
    }

    static byName(name) {
        return AppTheme.All.find(theme => theme.name === name) ?? null;
    }

    static withAlpha(color, alpha) {
        return color.withAlpha(alpha);
    }

    /** The theme a drawing is shown under: its own (each player on a page chooses), else the page's */
    static of(drawing) {
        return drawing?.theme ?? AppTheme.current;
    }
}
