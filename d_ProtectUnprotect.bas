
Function Hash(text As String) As String
    Dim md As Object
    Set md = CreateObject("System.Security.Cryptography.SHA256Managed")
    Dim bytes() As Byte
    bytes = StrConv(text, vbFromUnicode)
    Dim hashResult() As Byte
    hashResult = md.ComputeHash_2((bytes))
    Hash = ConvToHexString(hashResult)
End Function

Function ConvToHexString(ByRef arr() As Byte) As String
    Dim hexStr As String
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        hexStr = hexStr & Hex(arr(i))
    Next i
    ConvToHexString = hexStr
End Function

Function Base64Decode(base64String As String) As Byte()
    Dim xmlDoc As Object
    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
    Dim elem As Object
    Set elem = xmlDoc.createElement("base64")
    elem.DataType = "bin.base64"
    elem.text = base64String
    Base64Decode = elem.nodeTypedValue
End Function

Sub UnprotectSheet()
    If ActiveSheet.ProtectContents = False Then
        MsgJaDesprotegida = MsgBox("Os dados já estão desprotegidos.", vbInformation, "Desproteger dados")
        Exit Sub
    End If
    UserForm_Password.Show
End Sub

Sub ProtectSheet()
    If ActiveSheet.ProtectContents = True Then
        MsgJaProtegida = MsgBox("Os dados já estão protegidos.", vbInformation, "Proteger dados")
        Exit Sub
    End If
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True, Password:=StrConv(Base64Decode("UGFkaWxoYUgyTSo="), vbUnicode) 'Alterado na v1.8
    MsgProtegida = MsgBox("Dados protegidos!", vbInformation, "Proteger dados")
End Sub
