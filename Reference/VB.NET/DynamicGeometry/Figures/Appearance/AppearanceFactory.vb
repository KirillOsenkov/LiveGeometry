Imports System.Collections.Generic
Imports System.Drawing
Imports GuiLabs.Canvas.Renderer

Friend Class AppearanceFactory
    Implements IAppearanceFactory

    Protected Sub New()
        DefaultPointAppearance = CreatePointAppearance(Color.Gray, Color.LightGoldenrodYellow, PointShapeType.Circle, 3)
        Dim SelectedPointAppearance As IPointAppearance = CreatePointAppearance(Color.Gray, Color.LightGreen, PointShapeType.Circle, 3)
        PointStyles = New Dictionary(Of String, IPointAppearance)()
        PointStyles.Add(Points.FreePoint, DefaultPointAppearance)
        PointStyles.Add(Points.SelectedPoint, SelectedPointAppearance)

        DefaultLineAppearance = CreateLineAppearance(Color.Black, 1)
        LineStyles = New Dictionary(Of String, ILineAppearance)()
        LineStyles.Add(Lines.Line, DefaultLineAppearance)
    End Sub

    Public Class Points
        Public Const FreePoint As String = "Free point"
        Public Const SelectedPoint As String = "Selected point"
    End Class

    Public Class Lines
        Public Const Line As String = "Line"
    End Class

    Private mDefaultPointAppearance As IPointAppearance
    Public Property DefaultPointAppearance() As IPointAppearance
        Get
            Return mDefaultPointAppearance
        End Get
        Set(ByVal value As IPointAppearance)
            mDefaultPointAppearance = value
        End Set
    End Property

    Private mDefaultLineAppearance As ILineAppearance
    Public Property DefaultLineAppearance() As ILineAppearance
        Get
            Return mDefaultLineAppearance
        End Get
        Set(ByVal value As ILineAppearance)
            mDefaultLineAppearance = value
        End Set
    End Property

    Private Shared mInstance As AppearanceFactory = New AppearanceFactory()
    Public Shared ReadOnly Property Instance() As AppearanceFactory
        Get
            Return mInstance
        End Get
    End Property

    Public Function CreatePointAppearance(ByVal Color As System.Drawing.Color, ByVal FillColor As System.Drawing.Color, ByVal Shape As PointShapeType, ByVal Width As Integer) As DynamicGeometry.IPointAppearance Implements DynamicGeometry.IAppearanceFactory.CreatePointAppearance
        Dim P As PointAppearance = New PointAppearance()
        P.LineStyle = RendererSingleton.DrawOperations.Factory.ProduceNewLineStyleInfo(Color, 1)
        P.FillStyle = RendererSingleton.DrawOperations.Factory.ProduceNewFillStyleInfo(FillColor)
        P.Shape = Shape
        P.Width = Width
        Return P
    End Function

    Public Function CreateLineAppearance(ByVal Color As System.Drawing.Color, ByVal DrawWidth As Integer) As DynamicGeometry.ILineAppearance Implements DynamicGeometry.IAppearanceFactory.CreateLineAppearance
        Dim L As LineAppearance = New LineAppearance()
        L.LineStyle = RendererSingleton.DrawOperations.Factory.ProduceNewLineStyleInfo(Color, DrawWidth)
        Return L
    End Function

    Private PointStyles As Dictionary(Of String, IPointAppearance)
    Private LineStyles As Dictionary(Of String, ILineAppearance)

    Public Function FindPointAppearance(ByVal StyleName As String) As IPointAppearance Implements IAppearanceFactory.FindPointAppearance
        Dim foundAppearance As IPointAppearance = Nothing

        If PointStyles.TryGetValue(StyleName, foundAppearance) Then
            Return foundAppearance
        End If

        Return DefaultPointAppearance
    End Function

    Public Function FindLineAppearance(ByVal StyleName As String) As ILineAppearance Implements IAppearanceFactory.FindLineAppearance
        Dim foundAppearance As ILineAppearance = Nothing

        If LineStyles.TryGetValue(StyleName, foundAppearance) Then
            Return foundAppearance
        End If

        Return DefaultLineAppearance
    End Function

End Class
