Attribute VB_Name = "Fractal03"
Option Explicit
Option Base 1

Type img
    x As Double
    y As Double
End Type
Private Function cmult(v1 As img, v2 As img) As img
    Dim res As img
        res.x = v1.x * v2.x - v1.y * v2.y
        res.y = v1.x * v2.y + v1.y * v2.x
        cmult = res

End Function
Private Function cmod(v1 As img) As Double
    cmod = v1.x * v1.x + v1.y * v1.y

End Function

Private Function cadd(v1 As img, v2 As img) As img
    Dim res As img
        res.x = v1.x + v2.x
        res.y = v1.y + v2.y
        cadd = res

End Function
Private Function cexp(v1 As img, exp As Integer) As img
    Dim res As img
    Dim n As Integer
    res = v1
    For n = 0 To exp - 2
        res = cmult(res, v1)
    Next n
    cexp = res
    
End Function
Private Sub DrawFractal(SheetName As String, exp As Integer)

Dim ws1 As Worksheet

Const HEIGHT = 200
Const WIDTH = 200

Dim x As Integer
Dim y As Integer
Dim i As Integer
Dim xd As Double
Dim yd As Double
Dim rg1 As Range

Set ws1 = Worksheets(SheetName)
ActiveWindow.Zoom = 10
ws1.Columns.ColumnWidth = 2
ws1.Rows.RowHeight = 15

Dim Data As Variant
Dim z As img
Dim c As img

Set rg1 = Range(ws1.Cells(1, 1), ws1.Cells(HEIGHT, WIDTH))
rg1.Cells.Clear
Data = rg1.Value
FormatSheet rg1

Const IMAX = 12

For y = 0 To HEIGHT - 1
    For x = 0 To WIDTH - 1
    
        xd = x
        yd = y
        z.x = 0
        z.y = 0
    
        c.x = (xd - WIDTH / 2) * 4 / WIDTH
        c.y = (yd - HEIGHT / 2) * 4 / WIDTH
        i = 0
        While (cmod(z) < 8 And (i < IMAX))
            z = cadd(cexp(z, exp), c)
            i = i + 1
        Wend
        
        If i < IMAX Then
            Data(x + 1, y + 1) = i
        Else
            Data(x + 1, y + 1) = 0
        End If
    
    Next x
Next y

    rg1.Value = Data


End Sub

Public Sub Calc1()
Dim ws1 As Worksheet
Dim n As Integer

For n = 2 To 8
    
    Set ws1 = Worksheets.add(after:=Sheets(Sheets.Count))
    ws1.Name = "Fractal" + CStr(n)
    DrawFractal ws1.Name, n

Next n

End Sub

Private Sub FormatSheet(rg1 As Range)

    rg1.FormatConditions.Delete
    rg1.FormatConditions.AddColorScale ColorScaleType:=2
    With rg1.FormatConditions(1)
        .ColorScaleCriteria(1).Type = xlConditionValueLowestValue
        .ColorScaleCriteria(1).FormatColor.Color = vbGreen
        .ColorScaleCriteria(2).Type = xlConditionValueHighestValue
'        .ColorScaleCriteria(2).Type = xlConditionValueNumber
'        .ColorScaleCriteria(2).Value = 3
        .ColorScaleCriteria(2).FormatColor.Color = vbBlue
    End With

End Sub


