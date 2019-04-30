Imports System.CodeDom
Imports System.Collections
Imports System.Reflection
Imports System.Xml

Namespace ExcelXmlWriter
    Public NotInheritable Class WorksheetNamedRangeCollection
        Inherits CollectionBase
        Implements IWriter
        Implements ICodeWriter
        Implements IReader
        Default Public ReadOnly Property Item(index As Integer) As WorksheetNamedRange
            Get
                Return DirectCast(MyBase.InnerList(index), WorksheetNamedRange)
            End Get
        End Property

        Friend Sub New()
        End Sub

        Public Function Add(namedRange As WorksheetNamedRange) As Integer
            If namedRange Is Nothing Then
                Throw New ArgumentNullException("namedRange")
            End If
            Return MyBase.InnerList.Add(namedRange)
        End Function

        Private Sub ExcelXmlWriter_ICodeWriter_WriteTo(type As CodeTypeDeclaration, method As CodeMemberMethod, targetObject As CodeExpression) Implements ExcelXmlWriter.ICodeWriter.WriteTo
            For i As Integer = 0 To MyBase.InnerList.Count - 1
                Dim ıtem As WorksheetNamedRange = Me(i)
                Dim statements As CodeStatementCollection = method.Statements
                Dim codeObjectCreateExpression As CodeExpression() = New CodeExpression(0) {}
                Dim codeTypeReference As New CodeTypeReference(GetType(WorksheetNamedRange))
                Dim codePrimitiveExpression As CodeExpression() = New CodeExpression() {New CodePrimitiveExpression(ıtem.Name), New CodePrimitiveExpression(ıtem.RefersTo), New CodePrimitiveExpression(DirectCast(ıtem.Hidden, Object))}
                codeObjectCreateExpression(0) = New CodeObjectCreateExpression(codeTypeReference, codePrimitiveExpression)
                statements.Add(New CodeMethodInvokeExpression(targetObject, "Add", codeObjectCreateExpression))
            Next
        End Sub

        Private Sub ExcelXmlWriter_IReader_ReadXml(element As XmlElement) Implements ExcelXmlWriter.IReader.ReadXml
            If Not WorksheetNamedRangeCollection.IsElement(element) Then
                Throw New ArgumentException("Invalid element", "element")
            End If
            For Each childNode As XmlNode In element.ChildNodes
                Dim xmlElement As XmlElement = TryCast(childNode, XmlElement)
                If xmlElement Is Nothing OrElse Not WorksheetNamedRange.IsElement(xmlElement) Then
                    Continue For
                End If
                Dim worksheetNamedRange__1 As New WorksheetNamedRange()
                DirectCast(worksheetNamedRange__1, IReader).ReadXml(xmlElement)
                Me.Add(worksheetNamedRange__1)
            Next
        End Sub

        Private Sub ExcelXmlWriter_IWriter_WriteXml(writer As XmlWriter) Implements ExcelXmlWriter.IWriter.WriteXml
            writer.WriteStartElement("s", "Names", "urn:schemas-microsoft-com:office:spreadsheet")
            For i As Integer = 0 To MyBase.InnerList.Count - 1
                DirectCast(MyBase.InnerList(i), IWriter).WriteXml(writer)
            Next
            writer.WriteEndElement()
        End Sub

        Public Function Contains(item As WorksheetNamedRange) As Boolean
            Return MyBase.InnerList.Contains(item)
        End Function

        Public Sub CopyTo(array As WorksheetNamedRange(), index As Integer)
            MyBase.InnerList.CopyTo(array, index)
        End Sub

        Public Function IndexOf(item As WorksheetNamedRange) As Integer
            Return MyBase.InnerList.IndexOf(item)
        End Function

        Public Sub Insert(index As Integer, item As WorksheetNamedRange)
            MyBase.InnerList.Insert(index, item)
        End Sub

        Friend Shared Function IsElement(element As XmlElement) As Boolean
            Return Util.IsElement(element, "Names", "urn:schemas-microsoft-com:office:spreadsheet")
        End Function

        Public Sub Remove(item As WorksheetNamedRange)
            MyBase.InnerList.Remove(item)
        End Sub
    End Class
End Namespace
