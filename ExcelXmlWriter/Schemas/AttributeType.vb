Imports Bc
Imports System.Xml

Namespace ExcelXmlWriter.Schemas
    Public NotInheritable Class AttributeType
        Inherits SchemaType
        Implements IWriter
        Private _name As String

        Private _rsname As String

        Private _dataType As String

        Public Property DataType() As String
            Get
                Return Me._dataType
            End Get
            Set(value As String)
                Me._dataType = value
            End Set
        End Property

        Public Property Name() As String
            Get
                Return Me._name
            End Get
            Set(value As String)
                Me._name = value
            End Set
        End Property

        Public Property RowsetName() As String
            Get
                Return Me._rsname
            End Get
            Set(value As String)
                Me._rsname = value
            End Set
        End Property

        Public Sub New()
        End Sub

        Public Sub New(name As String, rowsetName As String, dataType As String)
            Me._name = name
            Me._rsname = rowsetName
            Me._dataType = dataType
        End Sub

        Private Sub ExcelXmlWriter_IWriter_WriteXml(writer As XmlWriter) Implements ExcelXmlWriter.IWriter.WriteXml
            writer.WriteStartElement("s", "AttributeType", "uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882")
            If Me._name IsNot Nothing Then
                writer.WriteAttributeString("s", "name", "uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882", Me._name)
            End If
            If Me._rsname IsNot Nothing Then
                writer.WriteAttributeString("rs", "name", "urn:schemas-microsoft-com:rowset", Me._rsname)
            End If
            If Me._dataType IsNot Nothing Then
                writer.WriteStartElement("dt", "datatype", "uuid:C2F41010-65B3-11d1-A29F-00AA00C14882")
                writer.WriteAttributeString("dt", "type", "uuid:C2F41010-65B3-11d1-A29F-00AA00C14882", Me._dataType)
                writer.WriteEndElement()
            End If
            writer.WriteEndElement()
        End Sub
    End Class
End Namespace
