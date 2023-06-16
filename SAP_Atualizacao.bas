Sub Atualizar_SAP() '//NUNCA ALTERAR O NOME DA SUB
    

        newCode1 = GetGitHubFileContent("lpadilhaa", "VBA_SAP", "main", "a_PreencherDados.bas")
                    ThisWorkbook.VBProject.VBComponents("a_PreecherDados").CodeModule.DeleteLines 1, ThisWorkbook.VBProject.VBComponents("a_PreecherDados").CodeModule.CountOfLines
                    ThisWorkbook.VBProject.VBComponents("a_PreecherDados").CodeModule.InsertLines 1, newCode1
        Set newCode1 = Nothing 'v1.6
    
        newCode2 = GetGitHubFileContent("lpadilhaa", "VBA_SAP", "main", "b_EnviosAPIs.bas") 'v1.6
                    ThisWorkbook.VBProject.VBComponents("b_EnviosAPIs").CodeModule.DeleteLines 1, ThisWorkbook.VBProject.VBComponents("b_EnviosAPIs").CodeModule.CountOfLines 'v1.6
                    ThisWorkbook.VBProject.VBComponents("b_EnviosAPIs").CodeModule.InsertLines 1, newCode2 'v1.6
        Set newCode2 = Nothing 'v1.6

        AtivarFiltro = Replace(ThisWorkbook.VBProject.VBComponents("ProtectUnprotect").CodeModule.Lines(40, 1), "AllowFiltering:=False", "AllowFiltering:=True") 'v1.8
        ThisWorkbook.VBProject.VBComponents("ProtectUnprotect").CodeModule.ReplaceLine 40, AtivarFiltro 'v1.8
    
        ActiveWindow.DisplayWorkbookTabs = False 'v1.6


    On Error Resume Next
        ActiveWorkbook.Queries.Add Name:="Param_APIToken", Formula:= _
            """X"" meta [IsParameterQuery=true, Type=""Any"", IsParameterQueryRequired=true]" 'v1.9
        ActiveWorkbook.Queries.Item("Param_APIToken").Formula = _
            """X"" meta [IsParameterQuery=true, Type=""Any"", IsParameterQueryRequired=True]" 'v1.9
    On Error GoTo -1
    On Error GoTo 0
    
    Dim consultas As Variant 'v1.9
    consultas = Array("Query_TorresAtivasVinculadas", "Query_Dominio_FundacaoPe", "Query_Dominio_FundacaoMastro", "Query_Dominio_FundacaoEstai", "BASE_BD_TorresLT", _
                      "Query_ID_zeq_cadeia_isol_lt_fase_2", "Query_ID_zeq_cadeia_isol_lt_fase_3", "Query_ID_zeq_condutor_fase2", "Query_ID_zeq_condutor_fase3", _
                      "Query_ID_zeq_pararaio_direito", "Query_ID_zeq_opgw_direito", "BASE_BD_OPGWLT", "BASE_BD_SerieEstrutura", "BASE_BD_Aterramento", "BASE_BD_ParaRaiosLT", "BASE_BD_VaosLT", _
                      "Query_ID_zlis", "Query_ID_zeq_estru_geral", "Query_ID_zeq_estru_autop", "Query_ID_zeq_estru_estai", "Query_ID_zeq_cadeia_isol", "Query_ID_zeq_aterramento", _
                      "Query_ID_zeq_condutor", "Query_ID_zeq_pararaio", "Query_ID_zeq_opgw", "Query_ID_zeq_servidao")  'v1.9
    
    Dim i As Long 'v1.9
    Dim oldFormula As String 'v1.9
    Dim newFormula As String 'v1.9
    
    For i = LBound(consultas) To UBound(consultas) 'v1.9
        oldFormula = ActiveWorkbook.Queries.Item(consultas(i)).Formula 'v1.9
        If InStr(oldFormula, "Param_APIToken]]") = 0 Then 'v1.9
            newFormula = Replace(oldFormula, ")),", ", [Headers=[Authorization=Param_APIToken]])),") 'v1.9
            ActiveWorkbook.Queries.Item(consultas(i)).Formula = newFormula 'v1.9
        End If 'v1.9
    Next i 'v1.9

'Atualização da consulta "BASE_BD_ProjetosLT" (necessário que fosse realizado à parte)
    oldFormula = ActiveWorkbook.Queries.Item("BASE_BD_ProjetosLT").Formula
    If InStr(oldFormula, "Param_APIToken]]") = 0 Then
    newFormula = Replace(oldFormula, "api/projeto_lt/listar"")),", "api/projeto_lt/listar"", [Headers=[Authorization=Param_APIToken]])),")
    ActiveWorkbook.Queries.Item("BASE_BD_ProjetosLT").Formula = newFormula
    End If

'Atualização de SAP já gerados:

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
    
    Sheets("zeq_estru_autop&estai").Unprotect (StrConv(Base64Decode("UGFkaWxoYUgyTSo="), vbUnicode))
    On Error Resume Next
        Range("Tab_zeq_estru_autop_estai[DESENHO FUNDAÇÃO PÉ]").Replace What:="0", Replacement:=vbNullString, LookAt:=xlWhole 'v1.7
    On Error GoTo -1
    On Error GoTo 0
    Sheets("zeq_estru_autop&estai").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True, Password:=StrConv(Base64Decode("UGFkaWxoYUgyTSo="), vbUnicode)
    
    Sheets("zeq_estru_geral").Unprotect (StrConv(Base64Decode("UGFkaWxoYUgyTSo="), vbUnicode)) 'v1.6
        Range("Tab_zeq_estru_geral[VÃO DE PESO (m)]").FormulaR1C1 = _
            "=IF([@SILHUETA]=""-"",""-"",IF(OR([@ALTITUDE]="""",OFFSET([@ALTITUDE],-1,0)="""",OFFSET([@ALTITUDE],1,0)=""""),"""",[@[VÃO DE VENTO (m)]]-(IFERROR((VLOOKUP(INDEX(BASE_BD_VaosLT[NomeCabo],MATCH(OFFSET([@[NÚMERO DE OPERAÇÃO]],-1,0),BASE_BD_VaosLT[torre_numero_torre_1],0)),BASE_CabosWithOPGW,5,0))*(((IFERROR(VALUE(OFFSET([@[ALTURA MISULA (m)]],-1,0)),0)+IFERROR(VALUE(OFFSET([@ALTITUDE],-1,0)),0))-(IFERROR(VALUE([@[ALTURA MISULA (m)]]),0)+IFERROR(VALUE([@ALTITUDE]),0)))/(OFFSET([@[C" & _
            "OMPRIMENTO DO VÃO (m)]],-1,0))),0)+IFERROR((VLOOKUP(INDEX(BASE_BD_VaosLT[NomeCabo],MATCH(OFFSET([@[NÚMERO DE OPERAÇÃO]],1,0),BASE_BD_VaosLT[torre_numero_torre_1],0)),BASE_CabosWithOPGW,5,0))*(((IFERROR(VALUE(OFFSET([@[ALTURA MISULA (m)]],1,0)),0)+IFERROR(VALUE(OFFSET([@ALTITUDE],1,0)),0))-(IFERROR(VALUE([@[ALTURA MISULA (m)]]),0)+IFERROR(VALUE([@ALTITUDE]),0)))/([@[" & _
            "COMPRIMENTO DO VÃO (m)]])),0))))" '\Corrigido na v1.7
        Range("Tab_zeq_estru_geral[VÃO DE PESO (m)]").Value = Range("Tab_zeq_estru_geral[VÃO DE PESO (m)]").Value 'v1.6
    Sheets("zeq_estru_geral").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True, Password:=StrConv(Base64Decode("UGFkaWxoYUgyTSo="), vbUnicode) 'v1.6

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
