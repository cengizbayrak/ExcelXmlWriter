Imports System.CodeDom
Imports System.Collections
Imports System.Globalization
Imports System.Reflection
Imports System.Xml

Namespace ExcelXmlWriter
    Public NotInheritable Class WorksheetStyleCollection
        Inherits CollectionBase
        Implements IWriter
        Implements IReader
        Implements ICodeWriter
        Friend _book As Workbook

        Default Public Property Item(index As Integer) As WorksheetStyle
            Get
                Return DirectCast(MyBase.InnerList(index), WorksheetStyle)
            End Get
            Set(value As WorksheetStyle)
                MyBase.InnerList(index) = value
            End Set
        End Property

        Default Public ReadOnly Property Item(id As String) As WorksheetStyle
            Get
                Dim ınt32 As Integer = Me.IndexOf(id)
                If ınt32 = -1 Then
                    Throw New ArgumentException(String.Concat("The specified style ", id, " does not exists in the collection"))
                End If
                Return DirectCast(MyBase.InnerList(ınt32), WorksheetStyle)
            End Get
        End Property

        Friend Sub New(book As Workbook)
            If book Is Nothing Then
                Throw New ArgumentNullException("book")
            End If
            Me._book = book
        End Sub

        Public Function Add(style As WorksheetStyle) As Integer
            If style Is Nothing Then
                Throw New ArgumentNullException("style")
            End If
            style._book = Me._book
            Return MyBase.InnerList.Add(style)
        End Function

        Public Function Add(id As String) As WorksheetStyle
            Dim worksheetStyle As New WorksheetStyle(id)
            Me.Add(worksheetStyle)
            Return worksheetStyle
        End Function

        Private Sub ExcelXmlWriter_ICodeWriter_WriteTo(type As CodeTypeDeclaration, method As CodeMemberMethod, targetObject As CodeExpression) Implements ExcelXmlWriter.ICodeWriter.WriteTo
            Dim codeMemberMethod As New CodeMemberMethod()
            codeMemberMethod.Name = "GenerateStyles"
            codeMemberMethod.Parameters.Add(New CodeParameterDeclarationExpression(GetType(WorksheetStyleCollection), "styles"))
            Util.AddComment(method, "Generate Styles")
            Dim statements As CodeStatementCollection = method.Statements
            Dim codeThisReferenceExpression As New CodeThisReferenceExpression()
            Dim name As String = codeMemberMethod.Name
            Dim codePrimitiveExpression As CodeExpression() = New CodeExpression() {targetObject}
            statements.Add(New CodeMethodInvokeExpression(codeThisReferenceExpression, name, codePrimitiveExpression))
            type.Members.Add(codeMemberMethod)
            For i As Integer = 0 To MyBase.InnerList.Count - 1
                Dim ıtem As WorksheetStyle = Me(i)
                Dim str As String = Util.CreateSafeName(ıtem.ID, "style")
                Dim type1 As Type = GetType(WorksheetStyle)
                Dim codeVariableReferenceExpression As New CodeVariableReferenceExpression("styles")
                codePrimitiveExpression = New CodeExpression() {New CodePrimitiveExpression(ıtem.ID)}
                Dim codeVariableDeclarationStatement As New CodeVariableDeclarationStatement(type1, str, New CodeMethodInvokeExpression(codeVariableReferenceExpression, "Add", codePrimitiveExpression))
                Util.AddComment(codeMemberMethod, ıtem.ID)
                codeMemberMethod.Statements.Add(codeVariableDeclarationStatement)
                DirectCast(ıtem, ICodeWriter).WriteTo(type, codeMemberMethod, New CodeVariableReferenceExpression(str))
            Next
        End Sub

        Private Sub ExcelXmlWriter_IReader_ReadXml(element As XmlElement) Implements ExcelXmlWriter.IReader.ReadXml
            If Not WorksheetStyleCollection.IsElement(element) Then
                Throw New ArgumentException("Invalid element", "element")
            End If
            For Each childNode As XmlNode In element.ChildNodes
                Dim xmlElement As XmlElement = TryCast(childNode, XmlElement)
                If xmlElement Is Nothing OrElse Not WorksheetStyle.IsElement(xmlElement) Then
                    Continue For
                End If
                Dim worksheetStyle__1 As New WorksheetStyle(Nothing)
                DirectCast(worksheetStyle__1, IReader).ReadXml(xmlElement)
                Me.Add(worksheetStyle__1)
            Next
        End Sub

        Private Sub ExcelXmlWriter_IWriter_WriteXml(writer As XmlWriter) Implements ExcelXmlWriter.IWriter.WriteXml
            writer.WriteStartElement("Styles", "urn:schemas-microsoft-com:office:spreadsheet")
            For i As Integer = 0 To MyBase.InnerList.Count - 1
                DirectCast(MyBase.InnerList(i), IWriter).WriteXml(writer)
            Next
            writer.WriteEndElement()
        End Sub

        Public Function Contains(item As WorksheetStyle) As Boolean
            Return MyBase.InnerList.Contains(item)
        End Function

        Public Sub CopyTo(array As WorksheetStyle(), index As Integer)
            MyBase.InnerList.CopyTo(array, index)
        End Sub

        Public Function IndexOf(item As WorksheetStyle) As Integer
            Return MyBase.InnerList.IndexOf(item)
        End Function

        Public Function IndexOf(id As String) As Integer
            For i As Integer = 0 To MyBase.InnerList.Count - 1
                If String.Compare(DirectCast(MyBase.InnerList(i), WorksheetStyle).ID, id, True, CultureInfo.InvariantCulture) = 0 Then
                    Return i
                End If
            Next
            Return -1
        End Function

        Public Sub Insert(index As Integer, item As WorksheetStyle)
            MyBase.InnerList.Insert(index, item)
        End Sub

        Friend Shared Function IsElement(element As XmlElement) As Boolean
            Return Util.IsElement(element, "Styles", "urn:schemas-microsoft-com:office:spreadsheet")
        End Function

        Public Sub Remove(item As WorksheetStyle)
            MyBase.InnerList.Remove(item)
        End Sub
    End Class
End Namespace
