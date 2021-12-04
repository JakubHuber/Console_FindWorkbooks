Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices
Module Module1

    Sub Main()
        Dim ExcelApp = New Excel.Application
        ExcelApp = Marshal.GetActiveObject("Excel.Application")

        Console.WriteLine("Active workbooks count: " & ExcelApp.Workbooks.Count.ToString)

        Dim oWorkbooks As Excel.Workbooks = ExcelApp.Workbooks

        Dim oWorkbook As Excel.Workbook
        For Each oWorkbook In oWorkbooks
            Console.WriteLine("Workbooks name: " & oWorkbook.Name)
        Next

        ReleaseAll(oWorkbooks)
        ReleaseAll(ExcelApp)

        Console.ReadLine()
    End Sub

    'ReleaseAll(xlApp)
    'ReleaseAll(xlWorkBook)
    Private Sub ReleaseAll(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub
End Module
