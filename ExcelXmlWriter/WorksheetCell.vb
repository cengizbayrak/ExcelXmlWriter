Imports System.CodeDom
Imports System.Collections
Imports System.Globalization
Imports System.Xml

Namespace ExcelXmlWriter
    Public NotInheritable Class WorksheetCell
        Implements IWriter
        Implements IReader
        Implements ICodeWriter
        Private _mergeAcross As Integer = -1

        Private _mergeDown As Integer = -1

        Private _styleID As String

        Private _formula As String

        Private _href As String

        Friend _row As WorksheetRow

        Private _index As Integer

        Private _comment As WorksheetComment

        Private _data As WorksheetCellData

        Private _namedCell As WorksheetNamedCellCollection

        Public ReadOnly Property Comment() As WorksheetComment
            Get
                If Me._comment Is Nothing Then
                    Me._comment = New WorksheetComment()
                End If
                Return Me._comment
            End Get
        End Property

        Public ReadOnly Property Data() As WorksheetCellData
            Get
                If Me._data Is Nothing Then
                    Me._data = New WorksheetCellData(Me)
                End If
                Return Me._data
            End Get
        End Property

        Public Property Formula() As String
            Get
                Return Me._formula
            End Get
            Set(value As String)
                Me._formula = value
            End Set
        End Property

        Public Property HRef() As String
            Get
                Return Me._href
            End Get
            Set(value As String)
                Me._href = value
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

        Friend ReadOnly Property IsSimple() As Boolean
            Get
                If Me._styleID IsNot Nothing AndAlso Me._data IsNot Nothing AndAlso Me._data.IsSimple AndAlso Me._index = 0 AndAlso Me._mergeAcross <= 0 AndAlso Me._mergeDown < 0 AndAlso Me._formula Is Nothing AndAlso Me._href Is Nothing AndAlso Me._comment Is Nothing AndAlso Me._namedCell Is Nothing Then
                    Return True
                End If
                Return False
            End Get
        End Property

        Public Property MergeAcross() As Integer
            Get
                If Me._mergeAcross = -1 Then
                    Return 0
                End If
                Return Me._mergeAcross
            End Get
            Set(value As Integer)
                If value < 0 Then
                    Throw New ArgumentOutOfRangeException("value")
                End If
                Me._mergeAcross = value
            End Set
        End Property

        Public Property MergeDown() As Integer
            Get
                If Me._mergeDown = -1 Then
                    Return 0
                End If
                Return Me._mergeDown
            End Get
            Set(value As Integer)
                If value < 0 Then
                    Throw New ArgumentOutOfRangeException("value")
                End If
                Me._mergeDown = value
            End Set
        End Property

        Public ReadOnly Property NamedCell() As WorksheetNamedCellCollection
            Get
                If Me._namedCell Is Nothing Then
                    Me._namedCell = New WorksheetNamedCellCollection()
                End If
                Return Me._namedCell
            End Get
        End Property

        Public ReadOnly Property Row() As WorksheetRow
            Get
                Return Me._row
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

        Public Sub New()
        End Sub

        Public Sub New(text As String)
            Me.Data.Text = text
            Me.Data.Type = DataType.[String]
        End Sub

        Public Sub New(text As String, styleID As String)
            Me.Data.Text = text
            Me._styleID = styleID
            Me.Data.Type = DataType.[String]
        End Sub

        Public Sub New(text As String, type As DataType)
            Me.Data.Text = text
            Me.Data.Type = type
        End Sub

        Public Sub New(text As String, type As DataType, styleID As String)
            Me.Data.Text = text
            Me.Data.Type = type
            Me._styleID = styleID
        End Sub

        Private Sub ExcelXmlWriter_ICodeWriter_WriteTo(type As CodeTypeDeclaration, method As CodeMemberMethod, targetObject As CodeExpression) Implements ExcelXmlWriter.ICodeWriter.WriteTo
            If Me._styleID IsNot Nothing Then
                Util.AddAssignment(method, targetObject, "StyleID", Me._styleID)
            End If
            If Me._data IsNot Nothing Then
                Util.Traverse(type, Me._data, method, targetObject, "Data")
            End If
            If Me._index <> 0 Then
                Util.AddAssignment(method, targetObject, "Index", Me._index)
            End If
            If Me._mergeAcross > 0 Then
                Util.AddAssignment(method, targetObject, "MergeAcross", Me._mergeAcross)
            End If
            If Me._mergeDown >= 0 Then
                Util.AddAssignment(method, targetObject, "MergeDown", Me._mergeDown)
            End If
            If Me._formula IsNot Nothing Then
                Util.AddAssignment(method, targetObject, "Formula", Me._formula)
            End If
            If Me._href IsNot Nothing Then
                Util.AddAssignment(method, targetObject, "HRef", Me._href)
            End If
            If Me._comment IsNot Nothing Then
                Util.Traverse(type, Me._comment, method, targetObject, "Comment")
            End If
            If Me._namedCell IsNot Nothing Then
                For Each str As String In Me._namedCell
                    Dim statements As CodeStatementCollection = method.Statements
                    Dim codePropertyReferenceExpression As New CodePropertyReferenceExpression(targetObject, "NamedCell")
                    Dim codePrimitiveExpression As CodeExpression() = New CodeExpression() {New CodePrimitiveExpression(str)}
                    statements.Add(New CodeMethodInvokeExpression(codePropertyReferenceExpression, "Add", codePrimitiveExpression))
                Next
            End If
        End Sub

        Private Sub ExcelXmlWriter_IReader_ReadXml(element As XmlElement) Implements ExcelXmlWriter.IReader.ReadXml
            If Not WorksheetCell.IsElement(element) Then
                Throw New ArgumentException("Invalid element", "element")
            End If
            Me._index = Util.GetAttribute(element, "Index", "urn:schemas-microsoft-com:office:spreadsheet", 0)
            Me._mergeAcross = Util.GetAttribute(element, "MergeAcross", "urn:schemas-microsoft-com:office:spreadsheet", -1)
            Me._mergeDown = Util.GetAttribute(element, "MergeDown", "urn:schemas-microsoft-com:office:spreadsheet", -1)
            Me._styleID = Util.GetAttribute(element, "StyleID", "urn:schemas-microsoft-com:office:spreadsheet")
            Me._formula = Util.GetAttribute(element, "Formula", "urn:schemas-microsoft-com:office:spreadsheet")
            Me._href = Util.GetAttribute(element, "HRef", "urn:schemas-microsoft-com:office:spreadsheet")
            For Each childNode As XmlNode In element.ChildNodes
                Dim xmlElement As XmlElement = TryCast(childNode, XmlElement)
                If xmlElement Is Nothing Then
                    Continue For
                End If
                If WorksheetCellData.IsElement(xmlElement) Then
                    DirectCast(Me.Data, IReader).ReadXml(xmlElement)
                ElseIf Not WorksheetComment.IsElement(xmlElement) Then
                    If xmlElement.LocalName <> "NamedCell" Then
                        Continue For
                    End If
                    Dim attribute As String = Util.GetAttribute(xmlElement, "Name", "urn:schemas-microsoft-com:office:spreadsheet")
                    If attribute Is Nothing OrElse attribute.Length <= 0 Then
                        Continue For
                    End If
                    Me.NamedCell.Add(attribute)
                Else
                    DirectCast(Me.Comment, IReader).ReadXml(xmlElement)
                End If
            Next
        End Sub

        Private Sub ExcelXmlWriter_IWriter_WriteXml(writer As XmlWriter) Implements ExcelXmlWriter.IWriter.WriteXml
            writer.WriteStartElement("s", "Cell", "urn:schemas-microsoft-com:office:spreadsheet")
            If Me._index <> 0 Then
                writer.WriteAttributeString("s", "Index", "urn:schemas-microsoft-com:office:spreadsheet", Me._index.ToString(CultureInfo.InvariantCulture))
            End If
            If Me._mergeAcross > 0 Then
                writer.WriteAttributeString("s", "MergeAcross", "urn:schemas-microsoft-com:office:spreadsheet", Me._mergeAcross.ToString(CultureInfo.InvariantCulture))
            End If
            If Me._mergeDown >= 0 Then
                writer.WriteAttributeString("s", "MergeDown", "urn:schemas-microsoft-com:office:spreadsheet", Me._mergeDown.ToString(CultureInfo.InvariantCulture))
            End If
            If Me._styleID IsNot Nothing Then
                writer.WriteAttributeString("s", "StyleID", "urn:schemas-microsoft-com:office:spreadsheet", Me._styleID)
            End If
            If Me._formula IsNot Nothing Then
                writer.WriteAttributeString("s", "Formula", "urn:schemas-microsoft-com:office:spreadsheet", Me._formula)
            End If
            If Me._href IsNot Nothing Then
                writer.WriteAttributeString("s", "HRef", "urn:schemas-microsoft-com:office:spreadsheet", Me._href)
            End If
            If Me._comment IsNot Nothing Then
                DirectCast(Me._comment, IWriter).WriteXml(writer)
            End If
            If Me._data IsNot Nothing Then
                DirectCast(Me._data, IWriter).WriteXml(writer)
            End If
            If Me._namedCell IsNot Nothing Then
                For Each str As String In Me._namedCell
                    writer.WriteStartElement("s", "NamedCell", "urn:schemas-microsoft-com:office:spreadsheet")
                    writer.WriteAttributeString("s", "Name", "urn:schemas-microsoft-com:office:spreadsheet", str)
                    writer.WriteEndElement()
                Next
            End If
            writer.WriteEndElement()
        End Sub

        Friend Shared Function IsElement(element As XmlElement) As Boolean
            Return Util.IsElement(element, "Cell", "urn:schemas-microsoft-com:office:spreadsheet")
        End Function
    End Class
End Namespace
