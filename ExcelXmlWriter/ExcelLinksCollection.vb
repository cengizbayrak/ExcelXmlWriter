Imports System.CodeDom
Imports System.Collections
Imports System.Globalization
Imports System.Reflection
Imports System.Xml

Namespace ExcelXmlWriter
	Public NotInheritable Class ExcelLinksCollection
		Inherits CollectionBase
		Implements IWriter
		Implements ICodeWriter
		Public Default ReadOnly Property Item(index As Integer) As SupBook
			Get
				Return DirectCast(MyBase.InnerList(index), SupBook)
			End Get
		End Property

		Friend Sub New()
		End Sub

		Public Function Add(link As SupBook) As Integer
			Return MyBase.InnerList.Add(link)
		End Function

		Public Function Add() As SupBook
			Dim supBook As New SupBook()
			Me.Add(supBook)
			Return supBook
		End Function

        Private Sub ExcelXmlWriter_ICodeWriter_WriteTo(type As CodeTypeDeclaration, method As CodeMemberMethod, targetObject As CodeExpression) Implements ExcelXmlWriter.ICodeWriter.WriteTo
            For i As Integer = 0 To MyBase.InnerList.Count - 1
                Dim ıtem As SupBook = Me(i)
                Dim str As String = String.Concat("SupBook", i.ToString(CultureInfo.InvariantCulture))
                Dim codeVariableDeclarationStatement As New CodeVariableDeclarationStatement(GetType(SupBook), str, New CodeMethodInvokeExpression(New CodePropertyReferenceExpression(targetObject, "Links"), "Add", New CodeExpression(-1) {}))
                method.Statements.Add(codeVariableDeclarationStatement)
                DirectCast(ıtem, ICodeWriter).WriteTo(type, method, New CodeVariableReferenceExpression(str))
            Next
        End Sub

        Private Sub ExcelXmlWriter_IWriter_WriteXml(writer As XmlWriter) Implements ExcelXmlWriter.IWriter.WriteXml
            For i As Integer = 0 To MyBase.InnerList.Count - 1
                DirectCast(MyBase.InnerList(i), IWriter).WriteXml(writer)
            Next
        End Sub

		Public Function Contains(link As SupBook) As Boolean
			Return MyBase.InnerList.Contains(link)
		End Function

		Public Sub CopyTo(array As SupBook(), index As Integer)
			MyBase.InnerList.CopyTo(array, index)
		End Sub

		Public Function IndexOf(item As SupBook) As Integer
			Return MyBase.InnerList.IndexOf(item)
		End Function

		Public Sub Insert(index As Integer, item As SupBook)
			MyBase.InnerList.Insert(index, item)
		End Sub

		Public Sub Remove(item As String)
			MyBase.InnerList.Remove(item)
		End Sub
	End Class
End Namespace
