Sub Atualizar_SAP() '//NUNCA ALTERAR O NOME DA SUB
	
  MsgBox "isso é uma mensagem de teste"
  
	Exit sub

		
        newCode1 = GetGitHubFileContent("lpadilhaa", "VBA_SAP", "main", "teste_SAP_Atualizacao.bas", "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")
                    ThisWorkbook.VBProject.VBComponents("a_PreecherDados").CodeModule.DeleteLines 1, ThisWorkbook.VBProject.VBComponents("a_PreecherDados").CodeModule.CountOfLines
                    ThisWorkbook.VBProject.VBComponents("a_PreecherDados").CodeModule.InsertLines 1, newCode1
		
End Sub
