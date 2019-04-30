Imports System.CodeDom
Imports System.Collections
Imports System.Collections.Specialized
Imports System.Xml

Namespace ExcelXmlWriter
    Public NotInheritable Class Worksheet
        Implements IWriter
        Implements IReader
        Implements ICodeWriter
        Private _name As String

        Private _table As WorksheetTable

        Private _pivotTable As ExcelXmlWriter.PivotTable

        Private _options As WorksheetOptions

        Private _names As WorksheetNamedRangeCollection

        Private _autoFilter As WorksheetAutoFilter

        Private _sorting As StringCollection

        Private _protected As Boolean

        Friend Shared cellDeclaration As CodeVariableDeclarationStatement

        Public ReadOnly Property AutoFilter() As WorksheetAutoFilter
            Get
                If Me._autoFilter Is Nothing Then
                    Me._autoFilter = New WorksheetAutoFilter()
                End If
                Return Me._autoFilter
            End Get
        End Property

        Public ReadOnly Property Name() As String
            Get
                Return Me._name
            End Get
        End Property

        Public ReadOnly Property Names() As WorksheetNamedRangeCollection
            Get
                If Me._names Is Nothing Then
                    Me._names = New WorksheetNamedRangeCollection()
                End If
                Return Me._names
            End Get
        End Property

        Public ReadOnly Property Options() As WorksheetOptions
            Get
                If Me._options Is Nothing Then
                    Me._options = New WorksheetOptions()
                End If
                Return Me._options
            End Get
        End Property

        Public ReadOnly Property PivotTable() As ExcelXmlWriter.PivotTable
            Get
                If Me._pivotTable Is Nothing Then
                    Me._pivotTable = New ExcelXmlWriter.PivotTable()
                End If
                Return Me._pivotTable
            End Get
        End Property

        Public Property [Protected]() As Boolean
            Get
                Return Me._protected
            End Get
            Set(value As Boolean)
                Me._protected = value
            End Set
        End Property

        Public ReadOnly Property Sorting() As StringCollection
            Get
                If Me._sorting Is Nothing Then
                    Me._sorting = New StringCollection()
                End If
                Return Me._sorting
            End Get
        End Property

        Public ReadOnly Property Table() As WorksheetTable
            Get
                If Me._table Is Nothing Then
                    Me._table = New WorksheetTable()
                End If
                Return Me._table
            End Get
        End Property

        Friend Sub New(name As String)
            Me._name = name
        End Sub

        Private Sub ExcelXmlWriter_ICodeWriter_WriteTo(type As CodeTypeDeclaration, method As CodeMemberMethod, targetObject As CodeExpression) Implements ExcelXmlWriter.ICodeWriter.WriteTo
            If Me._protected Then
                Util.AddAssignment(method, targetObject, "Protected", True)
            End If
            If Me._names IsNot Nothing Then
                Util.Traverse(type, Me._names, method, targetObject, "Names")
            End If
            If Me._table IsNot Nothing Then
                Util.Traverse(type, Me._table, method, targetObject, "Table")
            End If
            If Me._options IsNot Nothing Then
                Util.Traverse(type, Me._options, method, targetObject, "Options")
            End If
            If Me._autoFilter IsNot Nothing Then
                Util.Traverse(type, Me._autoFilter, method, targetObject, "AutoFilter")
            End If
            If Me._sorting IsNot Nothing Then
                For Each str As String In Me._sorting
                    Dim statements As CodeStatementCollection = method.Statements
                    Dim codePropertyReferenceExpression As New CodePropertyReferenceExpression(targetObject, "Sorting")
                    Dim codePrimitiveExpression As CodeExpression() = New CodeExpression() {New CodePrimitiveExpression(str)}
                    statements.Add(New CodeMethodInvokeExpression(codePropertyReferenceExpression, "Add", codePrimitiveExpression))
                Next
            End If
        End Sub

        Private Sub ExcelXmlWriter_IReader_ReadXml(element As XmlElement) Implements ExcelXmlWriter.IReader.ReadXml
            If Not Worksheet.IsElement(element) Then
                Throw New ArgumentException("Invalid element", "element")
            End If
            Me._name = Util.GetAttribute(element, "Name", "urn:schemas-microsoft-com:office:spreadsheet")
            Me._protected = Util.GetAttribute(element, "Protected", "urn:schemas-microsoft-com:office:spreadsheet", False)
            For Each childNode As XmlNode In element.ChildNodes
                Dim xmlElement As XmlElement = TryCast(childNode, XmlElement)
                If xmlElement Is Nothing Then
                    Continue For
                End If
                If WorksheetTable.IsElement(xmlElement) Then
                    DirectCast(Me.Table, IReader).ReadXml(xmlElement)
                ElseIf WorksheetNamedRangeCollection.IsElement(xmlElement) Then
                    DirectCast(Me.Names, IReader).ReadXml(xmlElement)
                ElseIf WorksheetAutoFilter.IsElement(xmlElement) Then
                    DirectCast(Me.AutoFilter, IReader).ReadXml(xmlElement)
                ElseIf Not WorksheetOptions.IsElement(xmlElement) Then
                    If Not Util.IsElement(xmlElement, "Sorting", "urn:schemas-microsoft-com:office:excel") Then
                        Continue For
                    End If
                    For Each childNode1 As XmlElement In xmlElement.ChildNodes
                        If Not Util.IsElement(childNode1, "Sort", "urn:schemas-microsoft-com:office:excel") Then
                            Continue For
                        End If
                        Me.Sorting.Add(childNode1.InnerText)
                    Next
                Else
                    DirectCast(Me.Options, IReader).ReadXml(xmlElement)
                End If
            Next
        End Sub

        Private Sub ExcelXmlWriter_IWriter_WriteXml(writer As XmlWriter) Implements ExcelXmlWriter.IWriter.WriteXml
            writer.WriteStartElement("s", "Worksheet", "urn:schemas-microsoft-com:office:spreadsheet")
            If Me._name IsNot Nothing Then
                writer.WriteAttributeString("Name", "urn:schemas-microsoft-com:office:spreadsheet", Me._name)
            End If
            If Me._protected Then
                writer.WriteAttributeString("Protected", "urn:schemas-microsoft-com:office:spreadsheet", "1")
            End If
            If Me._names IsNot Nothing Then
                DirectCast(Me._names, IWriter).WriteXml(writer)
            End If
            If Me._table IsNot Nothing Then
                DirectCast(Me._table, IWriter).WriteXml(writer)
            End If
            If Me._options IsNot Nothing Then
                DirectCast(Me._options, IWriter).WriteXml(writer)
            End If
            If Me._autoFilter IsNot Nothing Then
                DirectCast(Me._autoFilter, IWriter).WriteXml(writer)
            End If
            If Me._pivotTable IsNot Nothing Then
                DirectCast(Me._pivotTable, IWriter).WriteXml(writer)
            End If
            If Me._sorting IsNot Nothing Then
                writer.WriteStartElement("x", "Sorting", "urn:schemas-microsoft-com:office:excel")
                For Each str As String In Me._sorting
                    writer.WriteElementString("Sort", "urn:schemas-microsoft-com:office:excel", str)
                Next
                writer.WriteEndElement()
            End If
            writer.WriteEndElement()
        End Sub

        Friend Shared Function IsElement(element As XmlElement) As Boolean
            Return Util.IsElement(element, "Worksheet", "urn:schemas-microsoft-com:office:spreadsheet")
        End Function
    End Class
End Namespace
