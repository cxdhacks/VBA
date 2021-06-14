# Setting up font

[[How to setup font and size for entire workbook instead of sheet by sheet?](https://stackoverflow.com/questions/34039067/how-to-setup-font-and-size-for-entire-workbook-instead-of-sheet-by-sheet)](https://stackoverflow.com/questions/34039067/how-to-setup-font-and-size-for-entire-workbook-instead-of-sheet-by-sheet)

> If that suits you, you can work on `Styles`. Changing the default style of the workbook is very quick, but may have side effects. Try it.

```vb
With ActiveWorkbook.Styles("Normal").Font
    .Name = "Aharoni"
    .Size = 11
End With
```

> This code should loop through every sheet in your workbook and change the properties.

```vb
Sub SetFormat()
Dim ws as Worksheet
    For Each ws in Worksheets
  	'For and Each 都是关键词
         With ws
            .Cells.Font.Name = "Segoe UI"
            .Cells.Font.Size = 10
            .Cells.VerticalAlignment = xlCenter
         End With
    Next ws
End Sub
```

> The problem with the method is that the workbook default is unchanged and new worksheets added will use the default font. – [ChrisB](https://stackoverflow.com/users/5640342/chrisb) [Jun 15 '17 at 22:48](https://stackoverflow.com/questions/34039067/how-to-setup-font-and-size-for-entire-workbook-instead-of-sheet-by-sheet#comment76145145_34039284)

