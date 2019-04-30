Imports System.Globalization
Imports System.Xml

Namespace ExcelXmlWriter
	Public NotInheritable Class PTLineItem
		Implements IWriter
		Private _item As String

        Private _itemType As ExcelXmlWriter.ItemType

		Public Property Item() As String
			Get
				Return Me._item
			End Get
			Set
				Me._item = value
			End Set
		End Property

        Public Property ItemType() As ExcelXmlWriter.ItemType
            Get
                Return Me._itemType
            End Get
            Set(value As ExcelXmlWriter.ItemType)
                Me._itemType = Value
            End Set
        End Property

		Public Sub New()
		End Sub

		Public Sub New(item As String)
			Me._item = item
		End Sub

        Public Sub New(item As String, itemType As ExcelXmlWriter.ItemType)
            Me._item = item
            Me._itemType = itemType
        End Sub

        Private Sub ExcelXmlWriter_IWriter_WriteXml(writer As XmlWriter) Implements ExcelXmlWriter.IWriter.WriteXml
            writer.WriteStartElement("x", "PTLineItem", "urn:schemas-microsoft-com:office:excel")
            writer.WriteElementString("Item", "urn:schemas-microsoft-com:office:excel", Me._item)
            If Me._itemType <> ExcelXmlWriter.ItemType.NotSet Then
                writer.WriteElementString("ItemType", "urn:schemas-microsoft-com:office:excel", Me._itemType.ToString(CultureInfo.InvariantCulture))
            End If
            writer.WriteEndElement()
        End Sub
	End Class
End Namespace
