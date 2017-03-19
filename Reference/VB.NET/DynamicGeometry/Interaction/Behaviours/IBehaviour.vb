Imports GuiLabs.Canvas.Events
Imports GuiLabs.Canvas.Renderer

' Describes any object, that can control and modify DGDocument in responce on some user input events
' such as processing mouse and keyboard

Friend Interface IBehaviour
    Inherits IMouseHandler

    ' This class knows about DGDocument, and not only IDGDocument, 
    ' because it needs access to all of the class' methods and maybe even some internal secrets
    ' which are not exposed in IDGDocument
    Property Doc() As DGDocument

    ' you can ask IBehaviour to reset its state
    ' e.g. when the user cancels current operation
    ' and wants to return to the original state
    Sub Reset()

    Sub Draw(ByVal Renderer As IRenderer)

End Interface
