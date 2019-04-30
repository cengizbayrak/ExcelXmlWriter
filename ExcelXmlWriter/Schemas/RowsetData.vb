Imports ExcelXmlWriter
Imports System.Xml

Namespace ExcelXmlWriter.Schemas
	Public NotInheritable Class RowsetData
		Implements IWriter
		Private _rows As RowsetRowCollection

		Public ReadOnly Property Rows() As RowsetRowCollection
			Get
				If Me._rows Is Nothing Then
					Me._rows = New RowsetRowCollection()
				End If
				Return Me._rows
			End Get
		End Property

		Public Sub New()
		End Sub

        Private Sub ExcelXmlWriter_IWriter_WriteXml(writer As XmlWriter) Implements ExcelXmlWriter.IWriter.WriteXml
            writer.WriteStartElement("rs", "data", "urn:schemas-microsoft-com:rowset")
            If Me._rows IsNot Nothing Then
                DirectCast(Me._rows, IWriter).WriteXml(writer)
            End If
            writer.WriteEndElement()
        End Sub
	End Class
End Namespace
