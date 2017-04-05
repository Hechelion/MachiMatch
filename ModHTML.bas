Attribute VB_Name = "ModHTML"
Public Function decodingXML(nLinea As String) As String
Dim auxLinea As String
Dim i As Integer

auxLinea = nLinea
For i = 0 To UBound(lstCharSearch)
    auxLinea = Replace(auxLinea, lstCharSearch(i), lstCharReplace(i))
Next

decodingXML = auxLinea
End Function

Public Function GetAscii(nValor As String) As String
Dim auxValor As String
If InStr(nValor, "ascii_") > 0 Then
    auxValor = Replace(nValor, "ascii_", "")
    If IsNumeric(auxValor) Then
        auxValor = Chr(CLng(auxValor))
    Else
        MsgBox "ASCII ERROR", "ERROR"
        auxValor = ""
    End If
Else
    auxValor = nValor
End If

GetAscii = auxValor
End Function
