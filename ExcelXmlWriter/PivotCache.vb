Imports Bc.ExcelXmlWriter.Schemas
Imports System.Globalization
Imports System.Xml

Namespace ExcelXmlWriter
	Public NotInheritable Class PivotCache
		Implements IWriter
		Private _cacheIndex As Integer = -2147483648

        Private _schema As ExcelXmlWriter.Schemas.Schema

		Private _rowsetData As RowsetData

		Public Property CacheIndex() As Integer
			Get
				Return Me._cacheIndex
			End Get
			Set
				Me._cacheIndex = value
			End Set
		End Property

		Public ReadOnly Property Data() As RowsetData
			Get
				If Me._rowsetData Is Nothing Then
					Me._rowsetData = New RowsetData()
				End If
				Return Me._rowsetData
			End Get
		End Property

        Public ReadOnly Property Schema() As ExcelXmlWriter.Schemas.Schema
            Get
                If Me._schema Is Nothing Then
                    Me._schema = New ExcelXmlWriter.Schemas.Schema()
                End If
                Return Me._schema
            End Get
        End Property

		Friend Sub New()
		End Sub

        Private Sub ExcelXmlWriter_IWriter_WriteXml(writer As XmlWriter) Implements ExcelXmlWriter.IWriter.WriteXml
            writer.WriteStartElement("x", "PivotCache", "urn:schemas-microsoft-com:office:excel")
            If Me._cacheIndex <> -2147483648 Then
                writer.WriteElementString("CacheIndex", "urn:schemas-microsoft-com:office:excel", Me._cacheIndex.ToString(CultureInfo.InvariantCulture))
            End If
            If Me._schema IsNot Nothing Then
                DirectCast(Me._schema, IWriter).WriteXml(writer)
            End If
            If Me._rowsetData IsNot Nothing Then
                DirectCast(Me._rowsetData, IWriter).WriteXml(writer)
            End If
            writer.WriteEndElement()
        End Sub
	End Class
End Namespace
