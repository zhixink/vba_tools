Attribute VB_Name = "hextobin_module"
Function hexToBin(hexStr As String) As String
    Dim i As Integer, b As String, binStr As String
    For i = 1 To Len(hexStr)
        b = Application.hex2bin(Mid(hexStr, i, 1), 4)
        binStr = binStr & b
    Next i
    hexToBin = binStr
End Function
