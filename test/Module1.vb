Imports Bc.ExcelXmlWriter

Module Module1

    Sub Main()

        Dim book As New Workbook()
        Dim sheet As Worksheet = book.Worksheets.Add("sample")
        sheet.Options.FitToPage = True

        For i As Integer = 0 To 20
            Dim row As WorksheetRow = sheet.Table.Rows.Add()
            row.AutoFitHeight = True
            For j As Integer = 0 To 10
                Dim cell As WorksheetCell = row.Cells.Add("Hello world " & i & "x" & j)
            Next
        Next

        book.Save(AppDomain.CurrentDomain.BaseDirectory & "\test.xls")
        book.Save(AppDomain.CurrentDomain.BaseDirectory & "\test.xlsx")

    End Sub

End Module
