using Avalonia.Controls;

namespace LiveGeometry;

/// <summary>
/// Picked by <c>[PropertyGridPreferredEditor("SyntaxSample")]</c> on a string: a text of a
/// label in the theme's expression colors, to see them in while they are edited
/// (<see cref="AppTheme.ExpressionSample"/>). It is read against a drawing of its own -
/// points A and B and segment AB - that is never shown, so its names mean the same whatever
/// drawing is open, and what is typed into it is the sample's alone.
/// </summary>
public class SyntaxSampleEditorFactory : BaseValueEditorFactory<SyntaxSampleEditor, string>
{
    public SyntaxSampleEditorFactory()
    {
        // after the plain string editor: only the attribute chooses this one
        LoadOrder = 1;
    }
}

public class SyntaxSampleEditor : StringEditor, IExpressionHost
{
    protected override TextBox CreateTextBoxCore()
    {
        return new ColoredTextBox() { Host = this };
    }

    public Drawing Drawing => SampleDrawing;

    public IFigure Owner => null;

    public InsertContext Context => InsertContext.Text;

    static Drawing sampleDrawing;

    static Drawing SampleDrawing
    {
        get
        {
            if (sampleDrawing == null)
            {
                sampleDrawing = new Drawing(new Canvas());
                new DrawingDeserializer().ReadDrawing(sampleDrawing, """
                    <Drawing Version="1">
                      <Figures>
                        <FreePoint Name="A" X="0" Y="0" />
                        <FreePoint Name="B" X="3" Y="4" />
                        <Segment Name="AB">
                          <Dependency Name="A" />
                          <Dependency Name="B" />
                        </Segment>
                      </Figures>
                    </Drawing>
                    """);
            }

            return sampleDrawing;
        }
    }
}
