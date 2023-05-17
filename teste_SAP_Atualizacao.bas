Sub Atualizar_SAP() '//NUNCA ALTERAR O NOME DA SUB
	
  'MsgBox "isso é uma mensagem de teste"
  
	'Exit sub

		
        newCode1 = GetGitHubFileContent("lpadilhaa", "VBA_SAP", "main", "a_PreecherDados.bas")
                    ThisWorkbook.VBProject.VBComponents("a_PreecherDados").CodeModule.DeleteLines 1, ThisWorkbook.VBProject.VBComponents("a_PreecherDados").CodeModule.CountOfLines
                    ThisWorkbook.VBProject.VBComponents("a_PreecherDados").CodeModule.InsertLines 1, newCode1
	
	If Range("Label_NomeLT").locked = true Then
		
	Sheets("zeq_cadeia_isol").Unprotect (StrConv(Base64Decode("UGFkaWxoYUgyTSo="), vbUnicode))
	
	On Error Resume Next
		Range("Tab_zeq_cadeia_isol[DESENHO DO ISOLADOR]").Replace What:="0", Replacement:=vbNullString, LookAt:=xlWhole 'v1.4
	On Error GoTo -1
	On Error GoTo 0

	Sheets("zeq_cadeia_isol").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=False, Password:=StrConv(Base64Decode("UGFkaWxoYUgyTSo="), vbUnicode)
	
	End if

End Sub
