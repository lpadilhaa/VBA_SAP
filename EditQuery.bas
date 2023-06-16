Sub EditarConsultas()


On Error Resume Next
    ActiveWorkbook.Queries.Add Name:="Param_APIToken", Formula:= _
        """X"" meta [IsParameterQuery=true, Type=""Any"", IsParameterQueryRequired=true]"
    ActiveWorkbook.Queries.Item("Param_APIToken").Formula = _
        """X"" meta [IsParameterQuery=true, Type=""Any"", IsParameterQueryRequired=True]"
On Error GoTo -1
On Error GoTo 0

oldFormula = ActiveWorkbook.Queries.Item("Query_TorresAtivasVinculadas").Formula
If InStr(oldFormula, "Param_APIToken]]") = 0 Then
newFormula = Replace(oldFormula, ")),", ", [Headers=[Authorization=Param_APIToken]])),")
ActiveWorkbook.Queries.Item("Query_TorresAtivasVinculadas").Formula = newFormula
End If

oldFormula = ActiveWorkbook.Queries.Item("Query_Dominio_FundacaoPe").Formula
If InStr(oldFormula, "Param_APIToken]]") = 0 Then
newFormula = Replace(oldFormula, ")),", ", [Headers=[Authorization=Param_APIToken]])),")
ActiveWorkbook.Queries.Item("Query_Dominio_FundacaoPe").Formula = newFormula
End If

oldFormula = ActiveWorkbook.Queries.Item("Query_Dominio_FundacaoMastro").Formula
If InStr(oldFormula, "Param_APIToken]]") = 0 Then
newFormula = Replace(oldFormula, ")),", ", [Headers=[Authorization=Param_APIToken]])),")
ActiveWorkbook.Queries.Item("Query_Dominio_FundacaoMastro").Formula = newFormula
End If

oldFormula = ActiveWorkbook.Queries.Item("Query_Dominio_FundacaoEstai").Formula
If InStr(oldFormula, "Param_APIToken]]") = 0 Then
newFormula = Replace(oldFormula, ")),", ", [Headers=[Authorization=Param_APIToken]])),")
ActiveWorkbook.Queries.Item("Query_Dominio_FundacaoEstai").Formula = newFormula
End If

oldFormula = ActiveWorkbook.Queries.Item("BASE_BD_TorresLT").Formula
If InStr(oldFormula, "Param_APIToken]]") = 0 Then
newFormula = Replace(oldFormula, ")),", ", [Headers=[Authorization=Param_APIToken]])),")
ActiveWorkbook.Queries.Item("BASE_BD_TorresLT").Formula = newFormula
End If

oldFormula = ActiveWorkbook.Queries.Item("Query_ID_zeq_cadeia_isol_lt_fase_2").Formula
If InStr(oldFormula, "Param_APIToken]]") = 0 Then
newFormula = Replace(oldFormula, ")),", ", [Headers=[Authorization=Param_APIToken]])),")
ActiveWorkbook.Queries.Item("Query_ID_zeq_cadeia_isol_lt_fase_2").Formula = newFormula
End If

oldFormula = ActiveWorkbook.Queries.Item("Query_ID_zeq_cadeia_isol_lt_fase_3").Formula
If InStr(oldFormula, "Param_APIToken]]") = 0 Then
newFormula = Replace(oldFormula, ")),", ", [Headers=[Authorization=Param_APIToken]])),")
ActiveWorkbook.Queries.Item("Query_ID_zeq_cadeia_isol_lt_fase_3").Formula = newFormula
End If

oldFormula = ActiveWorkbook.Queries.Item("Query_ID_zeq_condutor_fase2").Formula
If InStr(oldFormula, "Param_APIToken]]") = 0 Then
newFormula = Replace(oldFormula, ")),", ", [Headers=[Authorization=Param_APIToken]])),")
ActiveWorkbook.Queries.Item("Query_ID_zeq_condutor_fase2").Formula = newFormula
End If

oldFormula = ActiveWorkbook.Queries.Item("Query_ID_zeq_condutor_fase3").Formula
If InStr(oldFormula, "Param_APIToken]]") = 0 Then
newFormula = Replace(oldFormula, ")),", ", [Headers=[Authorization=Param_APIToken]])),")
ActiveWorkbook.Queries.Item("Query_ID_zeq_condutor_fase3").Formula = newFormula
End If

oldFormula = ActiveWorkbook.Queries.Item("Query_ID_zeq_pararaio_direito").Formula
If InStr(oldFormula, "Param_APIToken]]") = 0 Then
newFormula = Replace(oldFormula, ")),", ", [Headers=[Authorization=Param_APIToken]])),")
ActiveWorkbook.Queries.Item("Query_ID_zeq_pararaio_direito").Formula = newFormula
End If

oldFormula = ActiveWorkbook.Queries.Item("Query_ID_zeq_opgw_direito").Formula
If InStr(oldFormula, "Param_APIToken]]") = 0 Then
newFormula = Replace(oldFormula, ")),", ", [Headers=[Authorization=Param_APIToken]])),")
ActiveWorkbook.Queries.Item("Query_ID_zeq_opgw_direito").Formula = newFormula
End If

oldFormula = ActiveWorkbook.Queries.Item("BASE_BD_OPGWLT").Formula
If InStr(oldFormula, "Param_APIToken]]") = 0 Then
newFormula = Replace(oldFormula, ")),", ", [Headers=[Authorization=Param_APIToken]])),")
ActiveWorkbook.Queries.Item("BASE_BD_OPGWLT").Formula = newFormula
End If

oldFormula = ActiveWorkbook.Queries.Item("BASE_BD_SerieEstrutura").Formula
If InStr(oldFormula, "Param_APIToken]]") = 0 Then
newFormula = Replace(oldFormula, ")),", ", [Headers=[Authorization=Param_APIToken]])),")
ActiveWorkbook.Queries.Item("BASE_BD_SerieEstrutura").Formula = newFormula
End If

oldFormula = ActiveWorkbook.Queries.Item("BASE_BD_Aterramento").Formula
If InStr(oldFormula, "Param_APIToken]]") = 0 Then
newFormula = Replace(oldFormula, ")),", ", [Headers=[Authorization=Param_APIToken]])),")
ActiveWorkbook.Queries.Item("BASE_BD_Aterramento").Formula = newFormula
End If

oldFormula = ActiveWorkbook.Queries.Item("BASE_BD_ParaRaiosLT").Formula
If InStr(oldFormula, "Param_APIToken]]") = 0 Then
newFormula = Replace(oldFormula, ")),", ", [Headers=[Authorization=Param_APIToken]])),")
ActiveWorkbook.Queries.Item("BASE_BD_ParaRaiosLT").Formula = newFormula
End If

oldFormula = ActiveWorkbook.Queries.Item("BASE_BD_ProjetosLT").Formula
If InStr(oldFormula, "Param_APIToken]]") = 0 Then
newFormula = Replace(oldFormula, ")),", ", [Headers=[Authorization=Param_APIToken]])),")
ActiveWorkbook.Queries.Item("BASE_BD_ProjetosLT").Formula = newFormula
End If

oldFormula = ActiveWorkbook.Queries.Item("BASE_BD_VaosLT").Formula
If InStr(oldFormula, "Param_APIToken]]") = 0 Then
newFormula = Replace(oldFormula, ")),", ", [Headers=[Authorization=Param_APIToken]])),")
ActiveWorkbook.Queries.Item("BASE_BD_VaosLT").Formula = newFormula
End If

oldFormula = ActiveWorkbook.Queries.Item("Query_ID_zlis").Formula
If InStr(oldFormula, "Param_APIToken]]") = 0 Then
newFormula = Replace(oldFormula, ")),", ", [Headers=[Authorization=Param_APIToken]])),")
ActiveWorkbook.Queries.Item("Query_ID_zlis").Formula = newFormula
End If

oldFormula = ActiveWorkbook.Queries.Item("Query_ID_zeq_estru_geral").Formula
If InStr(oldFormula, "Param_APIToken]]") = 0 Then
newFormula = Replace(oldFormula, ")),", ", [Headers=[Authorization=Param_APIToken]])),")
ActiveWorkbook.Queries.Item("Query_ID_zeq_estru_geral").Formula = newFormula
End If

oldFormula = ActiveWorkbook.Queries.Item("Query_ID_zeq_estru_autop").Formula
If InStr(oldFormula, "Param_APIToken]]") = 0 Then
newFormula = Replace(oldFormula, ")),", ", [Headers=[Authorization=Param_APIToken]])),")
ActiveWorkbook.Queries.Item("Query_ID_zeq_estru_autop").Formula = newFormula
End If

oldFormula = ActiveWorkbook.Queries.Item("Query_ID_zeq_estru_estai").Formula
If InStr(oldFormula, "Param_APIToken]]") = 0 Then
newFormula = Replace(oldFormula, ")),", ", [Headers=[Authorization=Param_APIToken]])),")
ActiveWorkbook.Queries.Item("Query_ID_zeq_estru_estai").Formula = newFormula
End If

oldFormula = ActiveWorkbook.Queries.Item("Query_ID_zeq_cadeia_isol").Formula
If InStr(oldFormula, "Param_APIToken]]") = 0 Then
newFormula = Replace(oldFormula, ")),", ", [Headers=[Authorization=Param_APIToken]])),")
ActiveWorkbook.Queries.Item("Query_ID_zeq_cadeia_isol").Formula = newFormula
End If

oldFormula = ActiveWorkbook.Queries.Item("Query_ID_zeq_aterramento").Formula
If InStr(oldFormula, "Param_APIToken]]") = 0 Then
newFormula = Replace(oldFormula, ")),", ", [Headers=[Authorization=Param_APIToken]])),")
ActiveWorkbook.Queries.Item("Query_ID_zeq_aterramento").Formula = newFormula
End If

oldFormula = ActiveWorkbook.Queries.Item("Query_ID_zeq_condutor").Formula
If InStr(oldFormula, "Param_APIToken]]") = 0 Then
newFormula = Replace(oldFormula, ")),", ", [Headers=[Authorization=Param_APIToken]])),")
ActiveWorkbook.Queries.Item("Query_ID_zeq_condutor").Formula = newFormula
End If

oldFormula = ActiveWorkbook.Queries.Item("Query_ID_zeq_pararaio").Formula
If InStr(oldFormula, "Param_APIToken]]") = 0 Then
newFormula = Replace(oldFormula, ")),", ", [Headers=[Authorization=Param_APIToken]])),")
ActiveWorkbook.Queries.Item("Query_ID_zeq_pararaio").Formula = newFormula
End If

oldFormula = ActiveWorkbook.Queries.Item("Query_ID_zeq_opgw").Formula
If InStr(oldFormula, "Param_APIToken]]") = 0 Then
newFormula = Replace(oldFormula, ")),", ", [Headers=[Authorization=Param_APIToken]])),")
ActiveWorkbook.Queries.Item("Query_ID_zeq_opgw").Formula = newFormula
End If

oldFormula = ActiveWorkbook.Queries.Item("Query_ID_zeq_servidao").Formula
If InStr(oldFormula, "Param_APIToken]]") = 0 Then
newFormula = Replace(oldFormula, ")),", ", [Headers=[Authorization=Param_APIToken]])),")
ActiveWorkbook.Queries.Item("Query_ID_zeq_servidao").Formula = newFormula
End If


End Sub

