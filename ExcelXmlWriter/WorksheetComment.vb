Imports System.CodeDom
Imports System.Collections
Imports System.Globalization
Imports System.Xml

Namespace ExcelXmlWriter
    Public NotInheritable Class WorksheetComment
        Implements IWriter
        Implements IReader
        Implements ICodeWriter
        Private _author As String

        Private _showAlways As Boolean

        Private _data As WorksheetCellData

        Public Property Author() As String
            Get
                Return Me._author
            End Get
            Set(value As String)
                Me._author = value
            End Set
        End Property

        Public ReadOnly Property Data() As WorksheetCellData
            Get
                If Me._data Is Nothing Then
                    Me._data = New WorksheetCellData(Me)
                End If
                Return Me._data
            End Get
        End Property

        Public Property ShowAlways() As Boolean
            Get
                Return Me._showAlways
            End Get
            Set(value As Boolean)
                Me._showAlways = value
            End Set
        End Property

        Friend Sub New()
        End Sub

        Private Sub ExcelXmlWriter_ICodeWriter_WriteTo(type As CodeTypeDeclaration, method As CodeMemberMethod, targetObject As CodeExpression) Implements ExcelXmlWriter.ICodeWriter.WriteTo
            If Me._showAlways Then
                Util.AddAssignment(method, targetObject, "ShowAlways", Me._showAlways)
            End If
            If Me._author IsNot Nothing Then
                Util.AddAssignment(method, targetObject, "Author", Me._author)
            End If
            If Me._data IsNot Nothing Then
                Util.Traverse(type, Me._data, method, targetObject, "Data")
            End If
        End Sub

        Private Sub ExcelXmlWriter_IReader_ReadXml(element As XmlElement) Implements ExcelXmlWriter.IReader.ReadXml
            If Not WorksheetComment.IsElement(element) Then
                Throw New ArgumentException("Invalid element", "element")
            End If
            Me._showAlways = Util.GetAttribute(element, "ShowAlways", "urn:schemas-microsoft-com:office:spreadsheet", False)
            Me._author = Util.GetAttribute(element, "Author", "urn:schemas-microsoft-com:office:spreadsheet")
            For Each childNode As XmlNode In element.ChildNodes
                Dim xmlElement As XmlElement = TryCast(childNode, XmlElement)
                If xmlElement Is Nothing OrElse Not WorksheetCellData.IsElement(xmlElement) Then
                    Continue For
                End If
                DirectCast(Me.Data, IReader).ReadXml(xmlElement)
            Next
        End Sub

        Private Sub ExcelXmlWriter_IWriter_WriteXml(writer As XmlWriter) Implements ExcelXmlWriter.IWriter.WriteXml
            writer.WriteStartElement("s", "Comment", "urn:schemas-microsoft-com:office:spreadsheet")
            If Me._showAlways Then
                writer.WriteAttributeString("s", "ShowAlways", "urn:schemas-microsoft-com:office:spreadsheet", Me._showAlways.ToString(CultureInfo.InvariantCulture))
            End If
            If Me._author IsNot Nothing Then
                writer.WriteAttributeString("s", "Author", "urn:schemas-microsoft-com:office:spreadsheet", Me._author)
            End If
            If Me._data IsNot Nothing Then
                DirectCast(Me._data, IWriter).WriteXml(writer)
            End If
            writer.WriteEndElement()
        End Sub

        Friend Shared Function IsElement(element As XmlElement) As Boolean
            Return Util.IsElement(element, "Comment", "urn:schemas-microsoft-com:office:spreadsheet")
        End Function
    End Class
End Namespace
