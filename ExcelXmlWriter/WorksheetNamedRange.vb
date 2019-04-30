Imports System.Xml

Namespace ExcelXmlWriter
    Public NotInheritable Class WorksheetNamedRange
        Implements IWriter
        Implements IReader
        Private _name As String

        Private _refersTo As String

        Private _hidden As Boolean

        Public Property Hidden() As Boolean
            Get
                Return Me._hidden
            End Get
            Set(value As Boolean)
                Me._hidden = value
            End Set
        End Property

        Public ReadOnly Property Name() As String
            Get
                Return Me._name
            End Get
        End Property

        Public Property RefersTo() As String
            Get
                Return Me._refersTo
            End Get
            Set(value As String)
                Me._refersTo = value
            End Set
        End Property

        Friend Sub New()
        End Sub

        Public Sub New(name As String)
            If name Is Nothing Then
                Throw New ArgumentNullException("name")
            End If
            Me._name = name
        End Sub

        Public Sub New(name As String, refersTo As String, hidden As Boolean)
            Me.New(name)
            Me._refersTo = refersTo
            Me._hidden = hidden
        End Sub

        Private Sub ExcelXmlWriter_IReader_ReadXml(element As XmlElement) Implements ExcelXmlWriter.IReader.ReadXml
            If Not WorksheetNamedRange.IsElement(element) Then
                Throw New ArgumentException("Invalid element", "element")
            End If
            Me._name = Util.GetAttribute(element, "Name", "urn:schemas-microsoft-com:office:spreadsheet")
            Me._refersTo = Util.GetAttribute(element, "RefersTo", "urn:schemas-microsoft-com:office:spreadsheet")
            Me._hidden = Util.GetAttribute(element, "Hidden", "urn:schemas-microsoft-com:office:spreadsheet", False)
        End Sub

        Private Sub ExcelXmlWriter_IWriter_WriteXml(writer As XmlWriter) Implements ExcelXmlWriter.IWriter.WriteXml
            writer.WriteStartElement("s", "NamedRange", "urn:schemas-microsoft-com:office:spreadsheet")
            If Me._name IsNot Nothing Then
                writer.WriteAttributeString("s", "Name", "urn:schemas-microsoft-com:office:spreadsheet", Me._name)
            End If
            If Me._refersTo IsNot Nothing Then
                writer.WriteAttributeString("s", "RefersTo", "urn:schemas-microsoft-com:office:spreadsheet", Me._refersTo)
            End If
            If Me._hidden Then
                writer.WriteAttributeString("s", "Hidden", "urn:schemas-microsoft-com:office:spreadsheet", "1")
            End If
            writer.WriteEndElement()
        End Sub

        Friend Shared Function IsElement(element As XmlElement) As Boolean
            Return Util.IsElement(element, "NamedRange", "urn:schemas-microsoft-com:office:spreadsheet")
        End Function
    End Class
End Namespace
