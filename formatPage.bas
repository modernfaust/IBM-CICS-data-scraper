Attribute VB_Name = "Module3"
Sub format()
    
    For Each Sheet In Worksheets
        With Sheet.Range("D4:E5")
            .Name = "Calibri"
            .NumberFormat = "General"
            .Font.Size = 12
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
            .Font.Bold = True
        End With
        With Sheet.Range("D11:D14")
            .Name = "Calibri"
            .NumberFormat = "m/d/yyyy"
            .Font.Size = 12
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        With Sheet.Range("B20:D28")
            .Name = "Calibri"
            .Font.Size = 12
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
        End With
        With Sheet.Range("G12:G17")
            .Name = "Calibri"
            .Font.Size = 12
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
    Next Sheet
End Sub

