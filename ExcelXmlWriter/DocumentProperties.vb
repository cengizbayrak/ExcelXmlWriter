Imports System.CodeDom
Imports System.Collections
Imports System.Globalization
Imports System.Xml

Namespace ExcelXmlWriter
	Public NotInheritable Class DocumentProperties
		Implements IWriter
		Implements IReader
		Implements ICodeWriter
		Private _author As String

		Private _company As String

		Private _created As DateTime = DateTime.MinValue

		Private _lastAuthor As String

		Private _lastSaved As DateTime = DateTime.MinValue

		Private _manager As String

		Private _subject As String

		Private _title As String

		Private _version As String

		Public Property Author() As String
			Get
				Return Me._author
			End Get
			Set
				Me._author = value
			End Set
		End Property

		Public Property Company() As String
			Get
				Return Me._company
			End Get
			Set
				Me._company = value
			End Set
		End Property

		Public Property Created() As DateTime
			Get
				Return Me._created
			End Get
			Set
				Me._created = value
			End Set
		End Property

		Public Property LastAuthor() As String
			Get
				Return Me._lastAuthor
			End Get
			Set
				Me._lastAuthor = value
			End Set
		End Property

		Public Property LastSaved() As DateTime
			Get
				Return Me._lastSaved
			End Get
			Set
				Me._lastSaved = value
			End Set
		End Property

		Public Property Manager() As String
			Get
				Return Me._manager
			End Get
			Set
				Me._manager = value
			End Set
		End Property

		Public Property Subject() As String
			Get
				Return Me._subject
			End Get
			Set
				Me._subject = value
			End Set
		End Property

		Public Property Title() As String
			Get
				Return Me._title
			End Get
			Set
				Me._title = value
			End Set
		End Property

		Public Property Version() As String
			Get
				Return Me._version
			End Get
			Set
				Me._version = value
			End Set
		End Property

		Friend Sub New()
		End Sub

        Private Sub ExcelXmlWriter_ICodeWriter_WriteTo(type As CodeTypeDeclaration, method As CodeMemberMethod, targetObject As CodeExpression) Implements ExcelXmlWriter.ICodeWriter.WriteTo
            Util.AddComment(method, "Properties")
            If Me._title IsNot Nothing Then
                Util.AddAssignment(method, targetObject, "Title", Me._title)
            End If
            If Me._subject IsNot Nothing Then
                Util.AddAssignment(method, targetObject, "Subject", Me._subject)
            End If
            If Me._author IsNot Nothing Then
                Util.AddAssignment(method, targetObject, "Author", Me._author)
            End If
            If Me._lastAuthor IsNot Nothing Then
                Util.AddAssignment(method, targetObject, "LastAuthor", Me._lastAuthor)
            End If
            If Me._created <> DateTime.MinValue Then
                Util.AddAssignment(method, targetObject, "Created", Me._created)
            End If
            If Me._lastSaved <> DateTime.MinValue Then
                Util.AddAssignment(method, targetObject, "LastSaved", Me._lastSaved)
            End If
            If Me._manager IsNot Nothing Then
                Util.AddAssignment(method, targetObject, "Manager", Me._manager)
            End If
            If Me._company IsNot Nothing Then
                Util.AddAssignment(method, targetObject, "Company", Me._company)
            End If
            If Me._version IsNot Nothing Then
                Util.AddAssignment(method, targetObject, "Version", Me._version)
            End If
        End Sub

        Private Sub ExcelXmlWriter_IReader_ReadXml(element As XmlElement) Implements ExcelXmlWriter.IReader.ReadXml
            If Not DocumentProperties.IsElement(element) Then
                Throw New ArgumentException("Invalid element", "element")
            End If
            For Each childNode As XmlNode In element.ChildNodes
                Dim xmlElement As XmlElement = TryCast(childNode, XmlElement)
                If xmlElement Is Nothing Then
                    Continue For
                End If
                If Util.IsElement(xmlElement, "Title", "urn:schemas-microsoft-com:office:office") Then
                    Me._title = xmlElement.InnerText
                ElseIf Util.IsElement(xmlElement, "Subject", "urn:schemas-microsoft-com:office:office") Then
                    Me._subject = xmlElement.InnerText
                ElseIf Util.IsElement(xmlElement, "Author", "urn:schemas-microsoft-com:office:office") Then
                    Me._author = xmlElement.InnerText
                ElseIf Util.IsElement(xmlElement, "LastAuthor", "urn:schemas-microsoft-com:office:office") Then
                    Me._lastAuthor = xmlElement.InnerText
                ElseIf Util.IsElement(xmlElement, "Created", "urn:schemas-microsoft-com:office:office") Then
                    Me._created = DateTime.Parse(xmlElement.InnerText, CultureInfo.InvariantCulture)
                ElseIf Util.IsElement(xmlElement, "LastSaved", "urn:schemas-microsoft-com:office:office") Then
                    Me._lastSaved = DateTime.Parse(xmlElement.InnerText, CultureInfo.InvariantCulture)
                ElseIf Util.IsElement(xmlElement, "Manager", "urn:schemas-microsoft-com:office:office") Then
                    Me._manager = xmlElement.InnerText
                ElseIf Not Util.IsElement(xmlElement, "Company", "urn:schemas-microsoft-com:office:office") Then
                    If Not Util.IsElement(xmlElement, "Version", "urn:schemas-microsoft-com:office:office") Then
                        Continue For
                    End If
                    Me._version = xmlElement.InnerText
                Else
                    Me._company = xmlElement.InnerText
                End If
            Next
        End Sub

        Private Sub ExcelXmlWriter_IWriter_WriteXml(writer As XmlWriter) Implements ExcelXmlWriter.IWriter.WriteXml
            writer.WriteStartElement("o", "DocumentProperties", "urn:schemas-microsoft-com:office:office")
            If Me._title IsNot Nothing Then
                writer.WriteElementString("Title", "urn:schemas-microsoft-com:office:office", Me._title)
            End If
            If Me._subject IsNot Nothing Then
                writer.WriteElementString("Subject", "urn:schemas-microsoft-com:office:office", Me._subject)
            End If
            If Me._author IsNot Nothing Then
                writer.WriteElementString("Author", "urn:schemas-microsoft-com:office:office", Me._author)
            End If
            If Me._lastAuthor IsNot Nothing Then
                writer.WriteElementString("LastAuthor", "urn:schemas-microsoft-com:office:office", Me._lastAuthor)
            End If
            If Me._created <> DateTime.MinValue Then
                writer.WriteElementString("Created", "urn:schemas-microsoft-com:office:office", Me._created.ToString("s", CultureInfo.InvariantCulture))
            End If
            If Me._lastSaved <> DateTime.MinValue Then
                writer.WriteElementString("LastSaved", "urn:schemas-microsoft-com:office:office", Me._lastSaved.ToString("s", CultureInfo.InvariantCulture))
            End If
            If Me._manager IsNot Nothing Then
                writer.WriteElementString("Manager", "urn:schemas-microsoft-com:office:office", Me._manager)
            End If
            If Me._company IsNot Nothing Then
                writer.WriteElementString("Company", "urn:schemas-microsoft-com:office:office", Me._company)
            End If
            If Me._version IsNot Nothing Then
                writer.WriteElementString("Version", "urn:schemas-microsoft-com:office:office", Me._version)
            End If
            writer.WriteEndElement()
        End Sub

		Friend Shared Function IsElement(element As XmlElement) As Boolean
			Return Util.IsElement(element, "DocumentProperties", "urn:schemas-microsoft-com:office:office")
		End Function
	End Class
End Namespace
