Public Class DGCommand
    Inherits Command

    Private mParentDocument As DGDocument = Nothing
    Friend Property ParentDocument() As DGDocument
        Get
            Return mParentDocument
        End Get
        Set(ByVal Value As DGDocument)
            mParentDocument = Value
        End Set
    End Property

End Class
