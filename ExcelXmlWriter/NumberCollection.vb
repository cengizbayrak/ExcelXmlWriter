Imports System.CodeDom
Imports System.Collections
Imports System.Reflection
Imports System.Xml

Namespace ExcelXmlWriter
	Public NotInheritable Class NumberCollection
		Inherits CollectionBase
		Implements IWriter
		Implements ICodeWriter
		Public Default ReadOnly Property Item(index As Integer) As String
			Get
				Return DirectCast(MyBase.InnerList(index), String)
			End Get
		End Property

		Friend Sub New()
		End Sub

		Public Function Add(item As String) As Integer
			Return MyBase.InnerList.Add(item)
		End Function

        Private Sub ExcelXmlWriter_ICodeWriter_WriteTo(type As CodeTypeDeclaration, method As CodeMemberMethod, targetObject As CodeExpression) Implements ExcelXmlWriter.ICodeWriter.WriteTo
            For i As Integer = 0 To MyBase.InnerList.Count - 1
                Dim ıtem As String = Me(i)
                Dim codePropertyReferenceExpression As New CodePropertyReferenceExpression(targetObject, "Numbers")
                Dim codePrimitiveExpression As CodeExpression() = New CodeExpression() {New CodePrimitiveExpression(ıtem)}
                Dim codeMethodInvokeExpression As New CodeMethodInvokeExpression(codePropertyReferenceExpression, "Add", codePrimitiveExpression)
                method.Statements.Add(codeMethodInvokeExpression)
            Next
        End Sub

        Private Sub ExcelXmlWriter_IWriter_WriteXml(writer As XmlWriter) Implements ExcelXmlWriter.IWriter.WriteXml
            For i As Integer = 0 To MyBase.InnerList.Count - 1
                writer.WriteElementString("Number", "urn:schemas-microsoft-com:office:excel", Me(i))
            Next
        End Sub

		Public Function Contains(link As String) As Boolean
			Return MyBase.InnerList.Contains(link)
		End Function

		Public Sub CopyTo(array As String(), index As Integer)
			MyBase.InnerList.CopyTo(array, index)
		End Sub

		Public Function IndexOf(item As String) As Integer
			Return MyBase.InnerList.IndexOf(item)
		End Function

		Public Sub Insert(index As Integer, item As String)
			MyBase.InnerList.Insert(index, item)
		End Sub

		Public Sub Remove(item As String)
			MyBase.InnerList.Remove(item)
		End Sub
	End Class
End Namespace
