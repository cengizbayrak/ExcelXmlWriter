Imports System.CodeDom
Imports System.Collections
Imports System.Globalization
Imports System.Reflection
Imports System.Xml

Namespace ExcelXmlWriter
    Public NotInheritable Class WorksheetRowCollection
        Inherits CollectionBase
        Implements IWriter
        Implements ICodeWriter
        Private _table As WorksheetTable

        Default Public ReadOnly Property Item(index As Integer) As WorksheetRow
            Get
                Return DirectCast(MyBase.InnerList(index), WorksheetRow)
            End Get
        End Property

        Friend Sub New()
        End Sub

        Friend Sub New(table As WorksheetTable)
            If table Is Nothing Then
                Throw New ArgumentNullException("table")
            End If
            Me._table = table
        End Sub

        Public Function Add(row As WorksheetRow) As Integer
            If row Is Nothing Then
                Throw New ArgumentNullException("row")
            End If
            row._table = Me._table
            Return MyBase.InnerList.Add(row)
        End Function

        Public Function Add() As WorksheetRow
            Dim worksheetRow As New WorksheetRow()
            Me.Add(worksheetRow)
            Return worksheetRow
        End Function

        Private Sub ExcelXmlWriter_ICodeWriter_WriteTo(type As CodeTypeDeclaration, method As CodeMemberMethod, targetObject As CodeExpression) Implements ExcelXmlWriter.ICodeWriter.WriteTo
            For i As Integer = 0 To MyBase.InnerList.Count - 1
                Util.AddCommentSeparator(method)
                Dim ıtem As WorksheetRow = Me(i)
                Dim str As String = String.Concat("Row", i.ToString(CultureInfo.InvariantCulture))
                Dim codeVariableDeclarationStatement As New CodeVariableDeclarationStatement(GetType(WorksheetRow), str, New CodeMethodInvokeExpression(targetObject, "Add", New CodeExpression(-1) {}))
                method.Statements.Add(codeVariableDeclarationStatement)
                DirectCast(ıtem, ICodeWriter).WriteTo(type, method, New CodeVariableReferenceExpression(str))
            Next
        End Sub

        Private Sub ExcelXmlWriter_IWriter_WriteXml(writer As XmlWriter) Implements ExcelXmlWriter.IWriter.WriteXml
            For i As Integer = 0 To MyBase.InnerList.Count - 1
                DirectCast(MyBase.InnerList(i), IWriter).WriteXml(writer)
            Next
        End Sub

        Public Function Contains(item As WorksheetRow) As Boolean
            Return MyBase.InnerList.Contains(item)
        End Function

        Public Sub CopyTo(array As WorksheetRow(), index As Integer)
            MyBase.InnerList.CopyTo(array, index)
        End Sub

        Public Function IndexOf(item As WorksheetRow) As Integer
            Return MyBase.InnerList.IndexOf(item)
        End Function

        Public Sub Insert(index As Integer, row As WorksheetRow)
            If row Is Nothing Then
                Throw New ArgumentNullException("row")
            End If
            row._table = Me._table
            MyBase.InnerList.Insert(index, row)
        End Sub

        Public Sub Remove(row As WorksheetRow)
            row._table = Nothing
            MyBase.InnerList.Remove(row)
        End Sub
    End Class
End Namespace
