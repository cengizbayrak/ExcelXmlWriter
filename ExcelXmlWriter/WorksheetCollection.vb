Imports System.CodeDom
Imports System.Collections
Imports System.Globalization
Imports System.Reflection
Imports System.Xml

Namespace ExcelXmlWriter
	Public NotInheritable Class WorksheetCollection
		Inherits CollectionBase
		Implements IWriter
		Implements ICodeWriter
		Public Default Property Item(index As Integer) As Worksheet
			Get
				Return DirectCast(MyBase.InnerList(index), Worksheet)
			End Get
			Set
				MyBase.InnerList(index) = value
			End Set
		End Property

		Public Default Property Item(name As String) As Worksheet
			Get
				Dim ınt32 As Integer = Me.IndexOf(name)
				If ınt32 = -1 Then
					Throw New ArgumentException(String.Concat("The specified worksheet ", name, " does not exists in the collection"))
				End If
				Return DirectCast(MyBase.InnerList(ınt32), Worksheet)
			End Get
			Set
				Dim ınt32 As Integer = Me.IndexOf(name)
				If ınt32 = -1 Then
					Throw New ArgumentException(String.Concat("The specified worksheet ", name, " does not exists in the collection"))
				End If
				MyBase.InnerList(ınt32) = value
			End Set
		End Property

		Friend Sub New()
		End Sub

		Public Function Add(sheet As Worksheet) As Integer
			Return MyBase.InnerList.Add(sheet)
		End Function

		Public Function Add(name As String) As Worksheet
			Dim worksheet As New Worksheet(name)
			Me.Add(worksheet)
			Return worksheet
		End Function

        Private Sub ExcelXmlWriter_ICodeWriter_WriteTo(type As CodeTypeDeclaration, method As CodeMemberMethod, targetObject As CodeExpression) Implements ExcelXmlWriter.ICodeWriter.WriteTo
            For i As Integer = 0 To MyBase.InnerList.Count - 1
                Dim item As Worksheet = Me(i)
                Dim str As String = Util.CreateSafeName(item.Name, "Sheet")
                Dim codeMemberMethod As New CodeMemberMethod()
                codeMemberMethod.Name = String.Concat("GenerateWorksheet", str)
                codeMemberMethod.Parameters.Add(New CodeParameterDeclarationExpression(GetType(WorksheetCollection), "sheets"))
                Worksheet.cellDeclaration = Nothing
                Util.AddComment(method, String.Concat("Generate ", item.Name, " Worksheet"))
                Dim statements As CodeStatementCollection = method.Statements
                Dim codeThisReferenceExpression As New CodeThisReferenceExpression()
                Dim name As String = codeMemberMethod.Name
                Dim codePrimitiveExpression As CodeExpression() = New CodeExpression() {targetObject}
                statements.Add(New CodeMethodInvokeExpression(codeThisReferenceExpression, name, codePrimitiveExpression))
                type.Members.Add(codeMemberMethod)
                Dim type1 As Type = GetType(Worksheet)
                Dim codeVariableReferenceExpression As New CodeVariableReferenceExpression("sheets")
                codePrimitiveExpression = New CodeExpression() {New CodePrimitiveExpression(str)}
                Dim codeVariableDeclarationStatement As New CodeVariableDeclarationStatement(type1, "sheet", New CodeMethodInvokeExpression(codeVariableReferenceExpression, "Add", codePrimitiveExpression))
                codeMemberMethod.Statements.Add(codeVariableDeclarationStatement)
                DirectCast(item, ICodeWriter).WriteTo(type, codeMemberMethod, New CodeVariableReferenceExpression("sheet"))
            Next
        End Sub

        Private Sub ExcelXmlWriter_IWriter_WriteXml(writer As XmlWriter) Implements ExcelXmlWriter.IWriter.WriteXml
            For i As Integer = 0 To MyBase.InnerList.Count - 1
                DirectCast(MyBase.InnerList(i), IWriter).WriteXml(writer)
            Next
        End Sub

		Public Function Contains(item As Worksheet) As Boolean
			Return MyBase.InnerList.Contains(item)
		End Function

		Public Sub CopyTo(array As Worksheet(), index As Integer)
			MyBase.InnerList.CopyTo(array, index)
		End Sub

		Public Function IndexOf(item As Worksheet) As Integer
			Return MyBase.InnerList.IndexOf(item)
		End Function

		Public Function IndexOf(name As String) As Integer
			For i As Integer = 0 To MyBase.InnerList.Count - 1
				If String.Compare(DirectCast(MyBase.InnerList(i), Worksheet).Name, name, True, CultureInfo.InvariantCulture) = 0 Then
					Return i
				End If
			Next
			Return -1
		End Function

		Public Sub Insert(index As Integer, sheet As Worksheet)
			MyBase.InnerList.Insert(index, sheet)
		End Sub

		Public Sub Remove(item As Worksheet)
			MyBase.InnerList.Remove(item)
		End Sub

		Public Function ToArray() As Object()
			Return MyBase.InnerList.ToArray()
		End Function
	End Class
End Namespace
