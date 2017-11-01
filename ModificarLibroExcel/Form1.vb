Imports System.ComponentModel
Imports Microsoft.Office.Interop.Excel

Public Class Form1
    Dim ExcelApp = New Microsoft.Office.Interop.Excel.Application
    Dim Libro = ExcelApp.Workbooks.Open("C:\Users\USER-XPS\Desktop\Prueba1.xlsx")
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim UltimaFila As Integer
        'Ubicar ultima fila

        UltimaFila = ExcelApp.Cells(1, 1).End(XlDirection.xlDown).Row
        UltimaFila = UltimaFila + 1

        Libro.Worksheets("Hoja1").Cells(UltimaFila, 1) = "Texto de Prueba"

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Libro.Save()
        MsgBox("Los cambios se han guardado en " & Libro.Name)
        ExcelApp.Quit()
        Libro = Nothing
        ExcelApp = Nothing
        End

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Form1_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        If Libro.saved() = False Then
            Dim Respuesta As MsgBoxResult = MsgBox("Desea guardar los cambios en el libro " & Libro.Name & vbExclamation + vbYesNo, "Microsoft Excel")

            Select Case Respuesta
                Case MsgBoxResult.Yes
                    Libro.Save()
                    MsgBox("Los cambios se han guardado en " & Libro.Name)
                    ExcelApp.Quit()
                    Libro = Nothing
                    ExcelApp = Nothing
                Case MsgBoxResult.No
                    Libro.saved() = True
                    ExcelApp.Quit()
                    Libro = Nothing
                    ExcelApp = Nothing
            End Select

        Else
            ExcelApp.Quit()
            Libro = Nothing
            ExcelApp = Nothing
        End If
    End Sub
End Class
