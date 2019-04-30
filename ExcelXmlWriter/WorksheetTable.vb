Imports System.CodeDom
Imports System.Collections
Imports System.ComponentModel
Imports System.Globalization
Imports System.Xml

Namespace ExcelXmlWriter
    Public NotInheritable Class WorksheetTable
        Implements IWriter
        Implements IReader
        Implements ICodeWriter
        Private _expandedColumnCount As Integer = -2147483648

        Private _expandedRowCount As Integer = -2147483648

        Private _fullColumns As Integer = -2147483648

        Private _fullRows As Integer = -2147483648

        Private _columns As WorksheetColumnCollection

        Private _rows As WorksheetRowCollection

        Private _defaultRowHeight As Single = 12.75F

        Private _defaultColumnWidth As Single = 48.0F

        Private _styleID As String

        Public ReadOnly Property Columns() As WorksheetColumnCollection
            Get
                If Me._columns Is Nothing Then
                    Me._columns = New WorksheetColumnCollection(Me)
                End If
                Return Me._columns
            End Get
        End Property

        Public Property DefaultColumnWidth() As Single
            Get
                Return Me._defaultColumnWidth
            End Get
            Set(value As Single)
                Me._defaultColumnWidth = value
            End Set
        End Property

        Public Property DefaultRowHeight() As Single
            Get
                Return Me._defaultRowHeight
            End Get
            Set(value As Single)
                Me._defaultRowHeight = value
            End Set
        End Property

        Public Property ExpandedColumnCount() As Integer
            Get
                Return Me._expandedColumnCount
            End Get
            Set(value As Integer)
                Me._expandedColumnCount = value
            End Set
        End Property

        Public Property ExpandedRowCount() As Integer
            Get
                Return Me._expandedRowCount
            End Get
            Set(value As Integer)
                Me._expandedRowCount = value
            End Set
        End Property

        '[EditorBrowsable(,)]    // JustDecompile was unable to locate the assembly where attribute parameters types are defined. Generating parameters values is impossible.
        Public Property FullColumns() As Integer
            Get
                Return Me._fullColumns
            End Get
            Set(value As Integer)
                Me._fullColumns = value
            End Set
        End Property

        '[EditorBrowsable(,)]    // JustDecompile was unable to locate the assembly where attribute parameters types are defined. Generating parameters values is impossible.
        Public Property FullRows() As Integer
            Get
                Return Me._fullRows
            End Get
            Set(value As Integer)
                Me._fullRows = value
            End Set
        End Property

        Public ReadOnly Property Rows() As WorksheetRowCollection
            Get
                If Me._rows Is Nothing Then
                    Me._rows = New WorksheetRowCollection(Me)
                End If
                Return Me._rows
            End Get
        End Property

        Public Property StyleID() As String
            Get
                Return Me._styleID
            End Get
            Set(value As String)
                Me._styleID = value
            End Set
        End Property

        Friend Sub New()
        End Sub

        Private Sub ExcelXmlWriter_ICodeWriter_WriteTo(type As CodeTypeDeclaration, method As CodeMemberMethod, targetObject As CodeExpression) Implements ExcelXmlWriter.ICodeWriter.WriteTo
            If Me._defaultRowHeight <> 12.75F Then
                Util.AddAssignment(method, targetObject, "DefaultRowHeight", Me._defaultRowHeight)
            End If
            If Me._defaultColumnWidth <> 48.0F Then
                Util.AddAssignment(method, targetObject, "DefaultColumnWidth", Me._defaultColumnWidth)
            End If
            If Me._expandedColumnCount <> -2147483648 Then
                Util.AddAssignment(method, targetObject, "ExpandedColumnCount", Me._expandedColumnCount)
            End If
            If Me._expandedRowCount <> -2147483648 Then
                Util.AddAssignment(method, targetObject, "ExpandedRowCount", Me._expandedRowCount)
            End If
            If Me._fullColumns <> -2147483648 Then
                Util.AddAssignment(method, targetObject, "FullColumns", Me._fullColumns)
            End If
            If Me._fullRows <> -2147483648 Then
                Util.AddAssignment(method, targetObject, "FullRows", Me._fullRows)
            End If
            If Me._styleID IsNot Nothing Then
                Util.AddAssignment(method, targetObject, "StyleID", Me._styleID)
            End If
            If Me._columns IsNot Nothing Then
                Util.Traverse(type, Me._columns, method, targetObject, "Columns")
            End If
            If Me._rows IsNot Nothing Then
                Util.Traverse(type, Me._rows, method, targetObject, "Rows")
            End If
        End Sub

        Private Sub ExcelXmlWriter_IReader_ReadXml(element As XmlElement) Implements ExcelXmlWriter.IReader.ReadXml
            If Not WorksheetTable.IsElement(element) Then
                Throw New ArgumentException("Invalid element", "element")
            End If
            Me._defaultRowHeight = Util.GetAttribute(element, "DefaultRowHeight", "urn:schemas-microsoft-com:office:spreadsheet", 12.75F)
            Me._defaultColumnWidth = Util.GetAttribute(element, "DefaultColumnWidth", "urn:schemas-microsoft-com:office:spreadsheet", 48.0F)
            Me._expandedColumnCount = Util.GetAttribute(element, "ExpandedColumnCount", "urn:schemas-microsoft-com:office:spreadsheet", Int32.MinValue)
            Me._expandedRowCount = Util.GetAttribute(element, "ExpandedRowCount", "urn:schemas-microsoft-com:office:spreadsheet", Int32.MinValue)
            Me._fullColumns = Util.GetAttribute(element, "FullColumns", "urn:schemas-microsoft-com:office:excel", Int32.MinValue)
            Me._fullRows = Util.GetAttribute(element, "FullRows", "urn:schemas-microsoft-com:office:excel", Int32.MinValue)
            Me._fullRows = Util.GetAttribute(element, "FullRows", "urn:schemas-microsoft-com:office:excel", Int32.MinValue)
            Me._styleID = Util.GetAttribute(element, "StyleID", "urn:schemas-microsoft-com:office:spreadsheet")
            For Each childNode As XmlNode In element.ChildNodes
                Dim xmlElement As XmlElement = TryCast(childNode, XmlElement)
                If xmlElement Is Nothing Then
                    Continue For
                End If
                If Not WorksheetColumn.IsElement(xmlElement) Then
                    If Not WorksheetRow.IsElement(xmlElement) Then
                        Continue For
                    End If
                    Dim worksheetRow__1 As New WorksheetRow()
                    DirectCast(worksheetRow__1, IReader).ReadXml(xmlElement)
                    Me.Rows.Add(worksheetRow__1)
                Else
                    Dim worksheetColumn__2 As New WorksheetColumn()
                    DirectCast(worksheetColumn__2, IReader).ReadXml(xmlElement)
                    Me.Columns.Add(worksheetColumn__2)
                End If
            Next
        End Sub

        Private Sub ExcelXmlWriter_IWriter_WriteXml(writer As XmlWriter) Implements ExcelXmlWriter.IWriter.WriteXml
            writer.WriteStartElement("s", "Table", "urn:schemas-microsoft-com:office:spreadsheet")
            If Me._defaultRowHeight <> 12.75F Then
                writer.WriteAttributeString("s", "DefaultRowHeight", "urn:schemas-microsoft-com:office:spreadsheet", Me._defaultRowHeight.ToString(CultureInfo.InvariantCulture))
            End If
            If Me._defaultColumnWidth <> 48.0F Then
                writer.WriteAttributeString("s", "DefaultColumnWidth", "urn:schemas-microsoft-com:office:spreadsheet", Me._defaultColumnWidth.ToString(CultureInfo.InvariantCulture))
            End If
            If Me._expandedColumnCount <> -2147483648 Then
                writer.WriteAttributeString("s", "ExpandedColumnCount", "urn:schemas-microsoft-com:office:spreadsheet", Me._expandedColumnCount.ToString(CultureInfo.InvariantCulture))
            End If
            If Me._expandedRowCount <> -2147483648 Then
                writer.WriteAttributeString("s", "ExpandedRowCount", "urn:schemas-microsoft-com:office:spreadsheet", Me._expandedRowCount.ToString(CultureInfo.InvariantCulture))
            End If
            If Me._fullColumns <> -2147483648 Then
                writer.WriteAttributeString("x", "FullColumns", "urn:schemas-microsoft-com:office:excel", Me._fullColumns.ToString(CultureInfo.InvariantCulture))
            End If
            If Me._fullRows <> -2147483648 Then
                writer.WriteAttributeString("x", "FullRows", "urn:schemas-microsoft-com:office:excel", Me._fullRows.ToString(CultureInfo.InvariantCulture))
            End If
            If Me._styleID IsNot Nothing Then
                writer.WriteAttributeString("s", "StyleID", "urn:schemas-microsoft-com:office:spreadsheet", Me._styleID)
            End If
            If Me._columns IsNot Nothing Then
                DirectCast(Me._columns, IWriter).WriteXml(writer)
            End If
            If Me._rows IsNot Nothing Then
                DirectCast(Me._rows, IWriter).WriteXml(writer)
            End If
            writer.WriteEndElement()
        End Sub

        Friend Shared Function IsElement(element As XmlElement) As Boolean
            Return Util.IsElement(element, "Table", "urn:schemas-microsoft-com:office:spreadsheet")
        End Function
    End Class
End Namespace
