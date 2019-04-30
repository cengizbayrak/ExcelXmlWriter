Imports ExcelXmlWriter
Imports System.Xml

Namespace ExcelXmlWriter.Schemas
	Public NotInheritable Class RowsetColumn
		Implements IWriter
		Private _name As String

		Private _value As String

		Public Property Name() As String
			Get
				Return Me._name
			End Get
			Set
				Me._name = value
			End Set
		End Property

		Public Property Value() As String
			Get
				Return Me._value
			End Get
			Set
				Me._value = value
			End Set
		End Property

		Public Sub New(name As String, value As String)
			Me._name = name
			Me._value = value
		End Sub

        Private Sub ExcelXmlWriter_IWriter_WriteXml(writer As XmlWriter) Implements ExcelXmlWriter.IWriter.WriteXml
            writer.WriteAttributeString("z", Me._name, "#RowsetSchema", Me._value)
        End Sub
	End Class
End Namespace
