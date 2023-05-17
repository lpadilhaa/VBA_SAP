Sub Atualizar_SAP() '//NUNCA ALTERAR O NOME DA SUB
    

        newCode1 = GetGitHubFileContent("lpadilhaa", "VBA_SAP", "main", "a_PreencherDados.bas")
                    ThisWorkbook.VBProject.VBComponents("a_PreecherDados").CodeModule.DeleteLines 1, ThisWorkbook.VBProject.VBComponents("a_PreecherDados").CodeModule.CountOfLines
                    ThisWorkbook.VBProject.VBComponents("a_PreecherDados").CodeModule.InsertLines 1, newCode1
    
If Range("Label_NomeLT").Locked = True Then
        
    Sheets("zeq_cadeia_isol").Unprotect (StrConv(Base64Decode("UGFkaWxoYUgyTSo="), vbUnicode))
    On Error Resume Next
        Range("Tab_zeq_cadeia_isol[DESENHO DO ISOLADOR]").Replace What:="0", Replacement:=vbNullString, LookAt:=xlWhole 'v1.5
    On Error GoTo -1
    On Error GoTo 0
    Sheets("zeq_cadeia_isol").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True, Password:=StrConv(Base64Decode("UGFkaWxoYUgyTSo="), vbUnicode)
    
    Sheets("zeq_servidao").Unprotect (StrConv(Base64Decode("UGFkaWxoYUgyTSo="), vbUnicode))
    On Error Resume Next
        Range("Tab_zeq_servidao[OBSERVAÇÃO]").Replace What:="0", Replacement:="-", LookAt:=xlWhole 'v1.5
    On Error GoTo -1
    On Error GoTo 0
    Sheets("zeq_servidao").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True, Password:=StrConv(Base64Decode("UGFkaWxoYUgyTSo="), vbUnicode)

    Sheets("zeq_pararaio").Unprotect (StrConv(Base64Decode("UGFkaWxoYUgyTSo="), vbUnicode))
    On Error Resume Next
        Range("Tab_zeq_pararaio[DESENHO DO ARRANJO]").Replace What:="0", Replacement:=vbNullString, LookAt:=xlWhole 'v1.6
    On Error GoTo -1
    On Error GoTo 0
    Sheets("zeq_pararaio").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True, Password:=StrConv(Base64Decode("UGFkaWxoYUgyTSo="), vbUnicode)

    Sheets("zeq_estru_geral").Unprotect (StrConv(Base64Decode("UGFkaWxoYUgyTSo="), vbUnicode)) 'v1.6
    Sheets("zeq_estru_geral").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True, Password:=StrConv(Base64Decode("UGFkaWxoYUgyTSo="), vbUnicode) 'v1.6

    Sheets("zeq_estru_autop&estai").Unprotect (StrConv(Base64Decode("UGFkaWxoYUgyTSo="), vbUnicode)) 'v1.6
    Sheets("zeq_estru_autop&estai").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True, Password:=StrConv(Base64Decode("UGFkaWxoYUgyTSo="), vbUnicode) 'v1.6

    Sheets("zeq_aterramento").Unprotect (StrConv(Base64Decode("UGFkaWxoYUgyTSo="), vbUnicode)) 'v1.6
    Sheets("zeq_aterramento").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True, Password:=StrConv(Base64Decode("UGFkaWxoYUgyTSo="), vbUnicode) 'v1.6

    Sheets("zeq_acessos").Unprotect (StrConv(Base64Decode("UGFkaWxoYUgyTSo="), vbUnicode)) 'v1.6
    Sheets("zeq_acessos").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True, Password:=StrConv(Base64Decode("UGFkaWxoYUgyTSo="), vbUnicode) 'v1.6
    
    Sheets("zeq_condutor").Unprotect (StrConv(Base64Decode("UGFkaWxoYUgyTSo="), vbUnicode)) 'v1.6
    Sheets("zeq_condutor").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True, Password:=StrConv(Base64Decode("UGFkaWxoYUgyTSo="), vbUnicode) 'v1.6

    Sheets("zeq_opgw").Unprotect (StrConv(Base64Decode("UGFkaWxoYUgyTSo="), vbUnicode)) 'v1.6
    Sheets("zeq_opgw").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True, Password:=StrConv(Base64Decode("UGFkaWxoYUgyTSo="), vbUnicode) 'v1.6

End If

End Sub
