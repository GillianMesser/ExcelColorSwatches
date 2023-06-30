Attribute VB_Name = "ColorCodes"
' Note: in the AddColorCodes sub, change ws_name to whatever you have
'       named the worksheet you will use for changing the background
'       color manually in order to have the code automatically generated.
'
' Add the following, commented out code into the specific sheet you
'       will use for changing the background color manually to
'       have the code automatically generated.
'-------------------------------------------------------------------------
'Private Sub worksheet_change(ByVal Target As Range)
'On Error Resume Next
'
'Dim rng As Range
'
'' for the selected/changed cell
'For Each rng In Target
'    Call AddColor(rng)
'Next rng
'
'End Sub
'-------------------------------------------------------------------------


Public Sub AddColor(rng As Range)

Dim code As String, split_code() As String
Dim r As Long, g As Long, b As Long
code = rng.Value
    
' if the length is six then it is a hex code
If Len(code) = 6 Then
    ' put the hex code in as the cell background color
    rng.Interior.color = rgb(Application.Hex2Dec(Left(code, 2)), Application.Hex2Dec(Mid(code, 3, 2)), Application.Hex2Dec(Right(code, 2)))
Else
    ' split the code into its component parts
    split_code = Split(code, ", ")
    r = CLng(split_code(0))
    g = CLng(split_code(1))
    b = CLng(split_code(2))
    ' put the rgb code in as the cell background color
    rng.Interior.color = rgb(r, g, b)
End If

End Sub
'

Public Sub AddColorCodes()
Application.ScreenUpdating = False

Dim clr As Long
Dim r As Long
Dim b As Long
Dim g As Long
Dim rgb_code As String
Dim hex_code As String
Dim ws_name As String

' CHANGE THIS NAME to be your worksheet name
ws_name = "Color to Code"

Worksheets(ws_name).Activate

' CHANGE THESE VALUES for what c (column) and w (row) iterate over
' to increase the range that the code will look at.  The default is
' columns 2 to 9 (B to I) and rows 5 to 20.

For c = 2 To 9          ' all columns B-I
    For w = 5 To 20     ' all rows 5-20
        ' pull the interior color, calculate the r/g/b values
        clr = Cells(w, c).Interior.color
        r = clr Mod 256
        g = clr \ 256 Mod 256
        b = clr \ 65536 Mod 256
        rgb_code = CStr(r) & ", " & CStr(g) & ", " & CStr(b)
        
        ' convert to hex code
        hex_code = hex(rgb(b, g, r))
        
        ' print in cell, ignoring white cells
        If Not hex_code = "FFFFFF" Then
            Cells(w, c).Value = hex_code & vbCrLf & rgb_code
        End If
    Next w
Next c

Application.ScreenUpdating = True
End Sub
'
