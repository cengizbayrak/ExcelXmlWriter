Imports System.CodeDom
Imports System.Collections
Imports System.Globalization
Imports System.Reflection
Imports System.Xml

Namespace ExcelXmlWriter
	Public NotInheritable Class CrnCollection
		Inherits CollectionBase
		Implements IWriter
		Implements ICodeWriter
		Friend Shared GlobalCounter As Integer

		Public Default ReadOnly Property Item(index As Integer) As Crn
			Get
				Return DirectCast(MyBase.InnerList(index), Crn)
			End Get
		End Property

		Friend Sub New()
		End Sub

		Public Function Add(item As Crn) As Integer
			Return MyBase.InnerList.Add(item)
		End Function

		Public Function Add() As Crn
			Dim crn As New Crn()
			Me.Add(crn)
			Return crn
		End Function

        Private Sub ExcelXmlWriter_ICodeWriter_WriteTo(type As CodeTypeDeclaration, method As CodeMemberMethod, targetObject As CodeExpression) Implements ExcelXmlWriter.ICodeWriter.WriteTo
            For i As Integer = 0 To MyBase.InnerList.Count - 1
                Dim ıtem As Crn = Me(i)
                Dim globalCounter As Integer = CrnCollection.GlobalCounter
                CrnCollection.GlobalCounter = globalCounter + 1
                Dim ınt32 As Integer = globalCounter
                Dim str As String = String.Concat("Crn", ınt32.ToString(CultureInfo.InvariantCulture))
                Dim codeVariableDeclarationStatement As New CodeVariableDeclarationStatement(GetType(Crn), str, New CodeMethodInvokeExpression(New CodePropertyReferenceExpression(targetObject, "Operands"), "Add", New CodeExpression(-1) {}))
                method.Statements.Add(codeVariableDeclarationStatement)
                DirectCast(ıtem, ICodeWriter).WriteTo(type, method, New CodeVariableReferenceExpression(str))
            Next
        End Sub

        Private Sub ExcelXmlWriter_IWriter_WriteXml(writer As XmlWriter) Implements ExcelXmlWriter.IWriter.WriteXml
            For i As Integer = 0 To MyBase.InnerList.Count - 1
                DirectCast(MyBase.InnerList(i), IWriter).WriteXml(writer)
            Next
        End Sub

		Public Function Contains(link As Crn) As Boolean
			Return MyBase.InnerList.Contains(link)
		End Function

		Public Sub CopyTo(array As Crn(), index As Integer)
			MyBase.InnerList.CopyTo(array, index)
		End Sub

		Public Function IndexOf(item As Crn) As Integer
			Return MyBase.InnerList.IndexOf(item)
		End Function

		Public Sub Insert(index As Integer, item As Crn)
			MyBase.InnerList.Insert(index, item)
		End Sub

		Public Sub Remove(item As String)
			MyBase.InnerList.Remove(item)
		End Sub
	End Class
End Namespace
