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

    Sheets("zeq_cadeia_isol").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=False, Password:=StrConv(Base64Decode("UGFkaWxoYUgyTSo="), vbUnicode)
    
    
    Sheets("zeq_servidao").Unprotect (StrConv(Base64Decode("UGFkaWxoYUgyTSo="), vbUnicode))
    
    On Error Resume Next
        Range("Tab_zeq_servidao[OBSERVAÇÃO]").Replace What:="0", Replacement:="-", LookAt:=xlWhole 'v1.5
    On Error GoTo -1
    On Error GoTo 0

    Sheets("zeq_servidao").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=False, Password:=StrConv(Base64Decode("UGFkaWxoYUgyTSo="), vbUnicode)

    End If

End Sub
