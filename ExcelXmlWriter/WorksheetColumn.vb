Imports System.CodeDom
Imports System.Globalization
Imports System.Xml

Namespace ExcelXmlWriter
    Public NotInheritable Class WorksheetColumn
        Implements IWriter
        Implements IReader
        Implements ICodeWriter
        Private _width As Integer = -2147483648

        Private _hidden As Boolean

        Private _styleID As String

        Friend _table As WorksheetTable

        Private _index As Integer

        Private _span As Integer = -2147483648

        Private _autoFitWidth As Boolean

        Public Property AutoFitWidth() As Boolean
            Get
                Return Me._autoFitWidth
            End Get
            Set(value As Boolean)
                Me._autoFitWidth = Value
            End Set
        End Property

        Public Property Hidden() As Boolean
            Get
                Return Me._hidden
            End Get
            Set(value As Boolean)
                Me._hidden = Value
            End Set
        End Property

        Public Property Index() As Integer
            Get
                Return Me._index
            End Get
            Set(value As Integer)
                Me._index = Value
            End Set
        End Property

        Friend ReadOnly Property IsSimple() As Boolean
            Get
                If Me._width <> -2147483648 AndAlso Me._index = 0 AndAlso Not Me._hidden AndAlso Me._styleID Is Nothing AndAlso Me._span = -2147483648 Then
                    Return True
                End If
                Return False
            End Get
        End Property

        Public Property Span() As Integer
            Get
                If Me._span = -2147483648 Then
                    Return 0
                End If
                Return Me._span
            End Get
            Set(value As Integer)
                Me._span = Value
            End Set
        End Property

        Public Property StyleID() As String
            Get
                Return Me._styleID
            End Get
            Set(value As String)
                Me._styleID = Value
            End Set
        End Property

        Public ReadOnly Property Table() As WorksheetTable
            Get
                Return Me._table
            End Get
        End Property

        Public Property Width() As Integer
            Get
                If Me._width <> -2147483648 OrElse Me._table Is Nothing Then
                    Return Me._width
                End If
                Return CInt(Math.Truncate(Me._table.DefaultColumnWidth))
            End Get
            Set(value As Integer)
                Me._width = Value
            End Set
        End Property

        Public Sub New()
        End Sub

        Public Sub New(width As Integer)
            Me._width = width
        End Sub

        Private Sub ExcelXmlWriter_ICodeWriter_WriteTo(type As CodeTypeDeclaration, method As CodeMemberMethod, targetObject As CodeExpression) Implements ExcelXmlWriter.ICodeWriter.WriteTo
            If Me._index <> 0 Then
                Util.AddAssignment(method, targetObject, "Index", Me._index)
            End If
            If Me._width <> -2147483648 Then
                Util.AddAssignment(method, targetObject, "Width", Me._width)
            End If
            If Me._hidden Then
                Util.AddAssignment(method, targetObject, "Hidden", Me._hidden)
            End If
            If Me._autoFitWidth Then
                Util.AddAssignment(method, targetObject, "AutoFitWidth", Me._autoFitWidth)
            End If
            If Me._styleID IsNot Nothing Then
                Util.AddAssignment(method, targetObject, "StyleID", Me._styleID)
            End If
            If Me._span <> -2147483648 Then
                Util.AddAssignment(method, targetObject, "Span", Me._span)
            End If
        End Sub

        Private Sub ExcelXmlWriter_IReader_ReadXml(element As XmlElement) Implements ExcelXmlWriter.IReader.ReadXml
            If Not WorksheetColumn.IsElement(element) Then
                Throw New ArgumentException("Invalid element", "element")
            End If
            Me._index = Util.GetAttribute(element, "Index", "urn:schemas-microsoft-com:office:spreadsheet", 0)
            Me._span = Util.GetAttribute(element, "Span", "urn:schemas-microsoft-com:office:spreadsheet", Int32.MinValue)
            Me._width = Util.GetAttribute(element, "Width", "urn:schemas-microsoft-com:office:spreadsheet", Int32.MinValue)
            Me._hidden = Util.GetAttribute(element, "Hidden", "urn:schemas-microsoft-com:office:spreadsheet", False)
            Me._autoFitWidth = Util.GetAttribute(element, "AutoFitWidth", "urn:schemas-microsoft-com:office:spreadsheet", False)
            Me._styleID = Util.GetAttribute(element, "StyleID", "urn:schemas-microsoft-com:office:spreadsheet")
        End Sub

        Private Sub ExcelXmlWriter_IWriter_WriteXml(writer As XmlWriter) Implements ExcelXmlWriter.IWriter.WriteXml
            writer.WriteStartElement("s", "Column", "urn:schemas-microsoft-com:office:spreadsheet")
            If Me._index <> 0 Then
                writer.WriteAttributeString("s", "Index", "urn:schemas-microsoft-com:office:spreadsheet", Me._index.ToString(CultureInfo.InvariantCulture))
            End If
            If Me._width <> -2147483648 Then
                writer.WriteAttributeString("s", "Width", "urn:schemas-microsoft-com:office:spreadsheet", Me._width.ToString(CultureInfo.InvariantCulture))
            End If
            If Me._hidden Then
                writer.WriteAttributeString("s", "Hidden", "urn:schemas-microsoft-com:office:spreadsheet", "1")
            End If
            If Me._autoFitWidth Then
                writer.WriteAttributeString("s", "AutoFitWidth", "urn:schemas-microsoft-com:office:spreadsheet", "1")
            End If
            If Me._styleID IsNot Nothing Then
                writer.WriteAttributeString("s", "StyleID", "urn:schemas-microsoft-com:office:spreadsheet", Me._styleID)
            End If
            If Me._span <> -2147483648 Then
                writer.WriteAttributeString("s", "Span", "urn:schemas-microsoft-com:office:spreadsheet", Me._span.ToString(CultureInfo.InvariantCulture))
            End If
            writer.WriteEndElement()
        End Sub

        Friend Shared Function IsElement(element As XmlElement) As Boolean
            Return Util.IsElement(element, "Column", "urn:schemas-microsoft-com:office:spreadsheet")
        End Function
    End Class
End Namespace
