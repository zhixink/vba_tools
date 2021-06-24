Attribute VB_Name = "bintohex_module"
Function binToHex(binStr As String) As String
    Dim i As Integer, h As String, hexStr As String
    For i = 1 To Len(binStr) Step 4
        h = Application.bin2hex(Mid(binStr, i, 4), 1)
        hexStr = hexStr & h
    Next i
    binToHex = hexStr
End Function
