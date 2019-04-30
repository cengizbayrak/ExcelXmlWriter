Imports System.Xml

Namespace ExcelXmlWriter
	Public NotInheritable Class PivotTable
		Implements IWriter
		Private _name As String

		Private _location As String

		Private _defaultVersion As String

		Private _versionLastUpdate As String

		Private _pivotFields As PivotFieldCollection

		Private _pTLinesItems As PTLineItemCollection

		Private _pTSource As PTSource

		Public Property DefaultVersion() As String
			Get
				Return Me._defaultVersion
			End Get
			Set
				Me._defaultVersion = value
			End Set
		End Property

		Public ReadOnly Property LineItems() As PTLineItemCollection
			Get
				If Me._pTLinesItems Is Nothing Then
					Me._pTLinesItems = New PTLineItemCollection()
				End If
				Return Me._pTLinesItems
			End Get
		End Property

		Public Property Location() As String
			Get
				Return Me._location
			End Get
			Set
				Me._location = value
			End Set
		End Property

		Public Property Name() As String
			Get
				Return Me._name
			End Get
			Set
				Me._name = value
			End Set
		End Property

		Public ReadOnly Property PivotFields() As PivotFieldCollection
			Get
				If Me._pivotFields Is Nothing Then
					Me._pivotFields = New PivotFieldCollection()
				End If
				Return Me._pivotFields
			End Get
		End Property

		Public ReadOnly Property Source() As PTSource
			Get
				If Me._pTSource Is Nothing Then
					Me._pTSource = New PTSource()
				End If
				Return Me._pTSource
			End Get
		End Property

		Public Property VersionLastUpdate() As String
			Get
				Return Me._versionLastUpdate
			End Get
			Set
				Me._versionLastUpdate = value
			End Set
		End Property

		Friend Sub New()
		End Sub

        Private Sub ExcelXmlWriter_IWriter_WriteXml(writer As XmlWriter) Implements ExcelXmlWriter.IWriter.WriteXml
            writer.WriteStartElement("x", "PivotTable", "urn:schemas-microsoft-com:office:excel")
            If Me._name IsNot Nothing Then
                writer.WriteElementString("Name", "urn:schemas-microsoft-com:office:excel", Me._name)
            End If
            writer.WriteStartElement("x", "ImmediateItemsOnDrop", "urn:schemas-microsoft-com:office:excel")
            writer.WriteEndElement()
            writer.WriteStartElement("x", "ShowPageMultipleItemLabel", "urn:schemas-microsoft-com:office:excel")
            writer.WriteEndElement()
            If Me._location IsNot Nothing Then
                writer.WriteElementString("Location", "urn:schemas-microsoft-com:office:excel", Me._location)
            End If
            If Me._defaultVersion IsNot Nothing Then
                writer.WriteElementString("DefaultVersion", "urn:schemas-microsoft-com:office:excel", Me._defaultVersion)
            End If
            If Me._versionLastUpdate IsNot Nothing Then
                writer.WriteElementString("VersionLastUpdate", "urn:schemas-microsoft-com:office:excel", Me._versionLastUpdate)
            End If
            If Me._pivotFields IsNot Nothing Then
                DirectCast(Me._pivotFields, IWriter).WriteXml(writer)
            End If
            If Me._pTLinesItems IsNot Nothing Then
                DirectCast(Me._pTLinesItems, IWriter).WriteXml(writer)
            End If
            If Me._pTSource IsNot Nothing Then
                DirectCast(Me._pTSource, IWriter).WriteXml(writer)
            End If
            writer.WriteEndElement()
        End Sub
	End Class
End Namespace
