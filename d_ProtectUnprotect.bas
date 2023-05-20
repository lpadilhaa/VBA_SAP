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
