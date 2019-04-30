Imports System.Globalization
Imports System.Xml

Namespace ExcelXmlWriter
	Public NotInheritable Class PTSource
		Implements IWriter
		Private _cacheIndex As Integer = -2147483648

		Private _versionLastRefresh As Integer = -2147483648

		Private _refreshName As String

		Private _refreshDate As DateTime = DateTime.MinValue

		Private _refreshDateCopy As DateTime = DateTime.MinValue

		Private _consolidationReference As PTConsolidationReference

		Public Property CacheIndex() As Integer
			Get
				Return Me._cacheIndex
			End Get
			Set
				Me._cacheIndex = value
			End Set
		End Property

		Public Property ConsolidationReference() As PTConsolidationReference
			Get
				If Me._consolidationReference Is Nothing Then
					Me._consolidationReference = New PTConsolidationReference()
				End If
				Return Me._consolidationReference
			End Get
			Set
				Me._consolidationReference = value
			End Set
		End Property

		Public Property RefreshDate() As DateTime
			Get
				Return Me._refreshDate
			End Get
			Set
				Me._refreshDate = value
			End Set
		End Property

		Public Property RefreshDateCopy() As DateTime
			Get
				Return Me._refreshDateCopy
			End Get
			Set
				Me._refreshDateCopy = value
			End Set
		End Property

		Public Property RefreshName() As String
			Get
				Return Me._refreshName
			End Get
			Set
				Me._refreshName = value
			End Set
		End Property

		Public Property VersionLastRefresh() As Integer
			Get
				Return Me._versionLastRefresh
			End Get
			Set
				Me._versionLastRefresh = value
			End Set
		End Property

		Friend Sub New()
		End Sub

        Private Sub ExcelXmlWriter_IWriter_WriteXml(writer As XmlWriter) Implements ExcelXmlWriter.IWriter.WriteXml
            writer.WriteStartElement("x", "PTSource", "urn:schemas-microsoft-com:office:excel")
            If Me._cacheIndex <> -2147483648 Then
                writer.WriteElementString("CacheIndex", "urn:schemas-microsoft-com:office:excel", Me._cacheIndex.ToString(CultureInfo.InvariantCulture))
            End If
            If Me._versionLastRefresh <> -2147483648 Then
                writer.WriteElementString("VersionLastRefresh", "urn:schemas-microsoft-com:office:excel", Me._versionLastRefresh.ToString(CultureInfo.InvariantCulture))
            End If
            If Me._refreshName IsNot Nothing Then
                writer.WriteElementString("RefreshName", "urn:schemas-microsoft-com:office:excel", Me._refreshName)
            End If
            If Me._refreshDate <> DateTime.MinValue Then
                writer.WriteElementString("RefreshDate", "urn:schemas-microsoft-com:office:excel", Me._refreshDate.ToString("s", CultureInfo.InvariantCulture))
            End If
            If Me._refreshDateCopy <> DateTime.MinValue Then
                writer.WriteElementString("RefreshDateCopy", "urn:schemas-microsoft-com:office:excel", Me._refreshDateCopy.ToString("s", CultureInfo.InvariantCulture))
            End If
            If Me._consolidationReference IsNot Nothing Then
                DirectCast(Me._consolidationReference, IWriter).WriteXml(writer)
            End If
            writer.WriteEndElement()
        End Sub
	End Class
End Namespace
