Imports System.Collections.Generic
Imports System.Collections.ObjectModel

Public Class CDocumentCollection

	Inherits Collection(Of CDocument)

	Public Event CollectionEmptied()
	Public Event CollectionNotMoreEmpty()

	Protected Overrides Sub InsertItem(ByVal index As Integer, ByVal item As CDocument)
		MyBase.InsertItem(index, item)
		AddHandler item.Closed, AddressOf OnDocumentClosed
		If Count = 1 Then RaiseEvent CollectionNotMoreEmpty()
	End Sub

	Protected Overrides Sub RemoveItem(ByVal index As Integer)
		RemoveHandler MyBase.Item(index).Closed, AddressOf OnDocumentClosed
		MyBase.RemoveItem(index)
		If Count = 0 Then RaiseEvent CollectionEmptied()
	End Sub

	Protected Overrides Sub SetItem(ByVal index As Integer, ByVal item As CDocument)
		RemoveHandler MyBase.Item(index).Closed, AddressOf OnDocumentClosed
		MyBase.SetItem(index, item)
		AddHandler item.Closed, AddressOf OnDocumentClosed
	End Sub

	Protected Sub OnDocumentClosed(ByVal Doc As CDocument)
		Remove(Doc)
	End Sub

End Class
