namespace DynamicGeometry;

/// <summary>
/// A tool that picks figures by clicking them (<see cref="ShowHideCreator"/>). A click on a
/// row of the Figure List picks the row's figure or lets go of it, as a click on the figure
/// does on the canvas - and reaches a hidden figure, which no click on the canvas can. What
/// is picked is selected, so that the list and the canvas show it (a hidden figure as a
/// ghost); the tool says when that changes (<see cref="Drawing.RaisePicksChanged"/>).
/// </summary>
public interface IFigurePicker
{
    void TogglePick(IFigure figure);
}
