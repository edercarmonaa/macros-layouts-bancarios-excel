Attribute VB_Name = "Módulo4"
Sub Limpiar()
    Worksheets(3).Range("A3:C200").ClearContents
    Worksheets(3).Range("E3:G200").ClearContents
    Range("A12:G10000").ClearContents
    Range("A12:G10000").Interior.Color = RGB(255, 255, 255)
    Range("A12:G10000").Font.Color = RGB(0, 0, 0)
End Sub
