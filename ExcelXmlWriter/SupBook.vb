Imports System.CodeDom
Imports System.Collections
Imports System.Collections.Specialized
Imports System.Xml

Namespace ExcelXmlWriter
	Public NotInheritable Class SupBook
		Implements IWriter
		Implements IReader
		Implements ICodeWriter
		Private _path As String

		Private _sheetNames As StringCollection

		Private _externNames As StringCollection

		Private _references As XctCollection

		Public ReadOnly Property ExternNames() As StringCollection
			Get
				If Me._externNames Is Nothing Then
					Me._externNames = New StringCollection()
				End If
				Return Me._externNames
			End Get
		End Property

		Public Property Path() As String
			Get
				Return Me._path
			End Get
			Set
				Me._path = value
			End Set
		End Property

		Public ReadOnly Property References() As XctCollection
			Get
				If Me._references Is Nothing Then
					Me._references = New XctCollection()
				End If
				Return Me._references
			End Get
		End Property

		Public ReadOnly Property SheetNames() As StringCollection
			Get
				If Me._sheetNames Is Nothing Then
					Me._sheetNames = New StringCollection()
				End If
				Return Me._sheetNames
			End Get
		End Property

		Public Sub New()
		End Sub

        Private Sub ExcelXmlWriter_ICodeWriter_WriteTo(type As CodeTypeDeclaration, method As CodeMemberMethod, targetObject As CodeExpression) Implements ExcelXmlWriter.ICodeWriter.WriteTo
            Dim codePrimitiveExpression As CodeExpression()
            If Me._path IsNot Nothing Then
                Util.AddAssignment(method, targetObject, "Path", Me._path)
            End If
            If Me._sheetNames IsNot Nothing Then
                For Each _sheetName As String In Me._sheetNames
                    Dim statements As CodeStatementCollection = method.Statements
                    Dim codePropertyReferenceExpression As New CodePropertyReferenceExpression(targetObject, "SheetNames")
                    codePrimitiveExpression = New CodeExpression() {New CodePrimitiveExpression(_sheetName)}
                    statements.Add(New CodeMethodInvokeExpression(codePropertyReferenceExpression, "Add", codePrimitiveExpression))
                Next
            End If
            If Me._externNames IsNot Nothing Then
                For Each _externName As String In Me._externNames
                    Dim codeStatementCollection As CodeStatementCollection = method.Statements
                    Dim codePropertyReferenceExpression1 As New CodePropertyReferenceExpression(targetObject, "ExternNames")
                    codePrimitiveExpression = New CodeExpression() {New CodePrimitiveExpression(_externName)}
                    codeStatementCollection.Add(New CodeMethodInvokeExpression(codePropertyReferenceExpression1, "Add", codePrimitiveExpression))
                Next
            End If
            If Me._references IsNot Nothing Then
                DirectCast(Me._references, ICodeWriter).WriteTo(type, method, targetObject)
            End If
        End Sub

        Private Sub ExcelXmlWriter_IReader_ReadXml(element As XmlElement) Implements ExcelXmlWriter.IReader.ReadXml
            If Not SupBook.IsElement(element) Then
                Throw New ArgumentException("Invalid element", "element")
            End If
            For Each childNode As XmlNode In element.ChildNodes
                Dim xmlElement As XmlElement = TryCast(childNode, XmlElement)
                If xmlElement Is Nothing Then
                    Continue For
                End If
                If Util.IsElement(xmlElement, "Path", "urn:schemas-microsoft-com:office:excel") Then
                    Me._path = xmlElement.InnerText
                ElseIf Util.IsElement(xmlElement, "SheetName", "urn:schemas-microsoft-com:office:excel") Then
                    Me.SheetNames.Add(xmlElement.InnerText)
                ElseIf Not Util.IsElement(xmlElement, "ExternName", "urn:schemas-microsoft-com:office:excel") Then
                    If Not Xct.IsElement(xmlElement) Then
                        Continue For
                    End If
                    Dim xct__1 As New Xct()
                    DirectCast(xct__1, IReader).ReadXml(xmlElement)
                    Me.References.Add(xct__1)
                Else
                    Me.ExternNames.Add(xmlElement.InnerText)
                End If
            Next
        End Sub

        Private Sub ExcelXmlWriter_IWriter_WriteXml(writer As XmlWriter) Implements ExcelXmlWriter.IWriter.WriteXml
            writer.WriteStartElement("x", "SupBook", "urn:schemas-microsoft-com:office:excel")
            If Me._path IsNot Nothing Then
                writer.WriteElementString("Path", "urn:schemas-microsoft-com:office:excel", Me._path)
            End If
            If Me._sheetNames IsNot Nothing Then
                For Each _sheetName As String In Me._sheetNames
                    writer.WriteElementString("SheetName", "urn:schemas-microsoft-com:office:excel", _sheetName)
                Next
            End If
            If Me._externNames IsNot Nothing Then
                For Each _externName As String In Me._externNames
                    writer.WriteStartElement("ExternName", "urn:schemas-microsoft-com:office:excel")
                    writer.WriteElementString("Name", "urn:schemas-microsoft-com:office:excel", _externName)
                    writer.WriteEndElement()
                Next
            End If
            If Me._references IsNot Nothing Then
                DirectCast(Me._references, IWriter).WriteXml(writer)
            End If
            writer.WriteEndElement()
        End Sub

		Friend Shared Function IsElement(element As XmlElement) As Boolean
			Return Util.IsElement(element, "SupBook", "urn:schemas-microsoft-com:office:excel")
		End Function
	End Class
End Namespace
