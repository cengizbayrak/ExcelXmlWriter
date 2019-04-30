Imports System.CodeDom
Imports System.Collections
Imports System.Globalization
Imports System.Xml

Namespace ExcelXmlWriter
    Public NotInheritable Class WorksheetRow
        Implements IWriter
        Implements IReader
        Implements ICodeWriter
        Private _index As Integer

        Private _height As Integer = -2147483648

        Private _autoFitHeight As Boolean = True

        Private _cells As WorksheetCellCollection

        Friend _table As WorksheetTable

        Public Property AutoFitHeight() As Boolean
            Get
                Return Me._autoFitHeight
            End Get
            Set(value As Boolean)
                Me._autoFitHeight = value
            End Set
        End Property

        Public ReadOnly Property Cells() As WorksheetCellCollection
            Get
                If Me._cells Is Nothing Then
                    Me._cells = New WorksheetCellCollection(Me)
                End If
                Return Me._cells
            End Get
        End Property

        Public Property Height() As Integer
            Get
                If Me._height <> -2147483648 OrElse Me._table Is Nothing Then
                    Return Me._height
                End If
                Return CInt(Math.Truncate(Me._table.DefaultRowHeight))
            End Get
            Set(value As Integer)
                Me._height = value
            End Set
        End Property

        Public Property Index() As Integer
            Get
                Return Me._index
            End Get
            Set(value As Integer)
                Me._index = value
            End Set
        End Property

        Public ReadOnly Property Table() As WorksheetTable
            Get
                Return Me._table
            End Get
        End Property

        Public Sub New()
        End Sub

        Private Sub ExcelXmlWriter_ICodeWriter_WriteTo(type As CodeTypeDeclaration, method As CodeMemberMethod, targetObject As CodeExpression) Implements ExcelXmlWriter.ICodeWriter.WriteTo
            If Me._index > 0 Then
                Util.AddAssignment(method, targetObject, "Index", Me._index)
            End If
            If Me._height <> -2147483648 Then
                Util.AddAssignment(method, targetObject, "Height", Me._height)
            End If
            If Not Me._autoFitHeight Then
                Util.AddAssignment(method, targetObject, "AutoFitHeight", Me._autoFitHeight)
            End If
            If Me._cells IsNot Nothing Then
                Util.Traverse(type, Me._cells, method, targetObject, "Cells")
            End If
        End Sub

        Private Sub ExcelXmlWriter_IReader_ReadXml(element As XmlElement) Implements ExcelXmlWriter.IReader.ReadXml
            If Not WorksheetRow.IsElement(element) Then
                Throw New ArgumentException("Invalid element", "element")
            End If
            Me._index = Util.GetAttribute(element, "Index", "urn:schemas-microsoft-com:office:spreadsheet", 0)
            Me._height = Util.GetAttribute(element, "Height", "urn:schemas-microsoft-com:office:spreadsheet", Int32.MinValue)
            Me._autoFitHeight = Util.GetAttribute(element, "AutoFitHeight", "urn:schemas-microsoft-com:office:spreadsheet", True)
            For Each childNode As XmlNode In element.ChildNodes
                Dim xmlElement As XmlElement = TryCast(childNode, XmlElement)
                If xmlElement Is Nothing OrElse Not WorksheetCell.IsElement(xmlElement) Then
                    Continue For
                End If
                Dim worksheetCell__1 As New WorksheetCell()
                DirectCast(worksheetCell__1, IReader).ReadXml(xmlElement)
                Me.Cells.Add(worksheetCell__1)
            Next
        End Sub

        Private Sub ExcelXmlWriter_IWriter_WriteXml(writer As XmlWriter) Implements ExcelXmlWriter.IWriter.WriteXml
            writer.WriteStartElement("s", "Row", "urn:schemas-microsoft-com:office:spreadsheet")
            If Me._index > 0 Then
                writer.WriteAttributeString("s", "Index", "urn:schemas-microsoft-com:office:spreadsheet", Me._index.ToString(CultureInfo.InvariantCulture))
            End If
            If Me._height <> -2147483648 Then
                writer.WriteAttributeString("s", "Height", "urn:schemas-microsoft-com:office:spreadsheet", Me._height.ToString(CultureInfo.InvariantCulture))
            End If
            If Not Me._autoFitHeight Then
                writer.WriteAttributeString("s", "AutoFitHeight", "urn:schemas-microsoft-com:office:spreadsheet", "0")
            End If
            If Me._cells IsNot Nothing Then
                DirectCast(Me._cells, IWriter).WriteXml(writer)
            End If
            writer.WriteEndElement()
        End Sub

        Friend Shared Function IsElement(element As XmlElement) As Boolean
            Return Util.IsElement(element, "Row", "urn:schemas-microsoft-com:office:spreadsheet")
        End Function
    End Class
End Namespace
