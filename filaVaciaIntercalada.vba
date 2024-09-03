Sub Macro1()
'
' Macro1 Macro
'
' Acceso directo: CTRL+x
'

 

Do While Not IsEmpty(ActiveCell)
ActiveCell.EntireRow.Insert
ActiveCell.Offset(2, 0).Select
Loop

End Sub
