
Public LoadAll As String
Public CodLT As String
Public TempoInicioAll As Date
Public TempoFimAll As Date


Sub LoadToAPI_All()

If Range("Label_NomeLT").Locked = False Then
    MsgNaoImportado = MsgBox("Importe os dados da LT e realize o preenchimento completo antes de enviar para o banco de dados", vbExclamation, "Dados não importados")
    Exit Sub
End If
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.ListObjects.Count > 0 Then
            For Each tbl In ws.ListObjects
                On Error GoTo filter_ctn
                If tbl.AutoFilter.FilterMode Then
                    MsgBox "A planilha """ & ws.Name & """ contém filtros ativos." & Chr(13) & Chr(13) & "Remova-o e confira os dados antes de enviar.", vbExclamation
                    Exit Sub
                End If
filter_ctn:
                On Error GoTo -1
                On Error GoTo 0
            Next tbl
        End If
    Next ws
        
Dim MsgLoad_All As VbMsgBoxResult

MsgLoad_All = MsgBox("Deseja enviar todos os dados para o banco de dados?" & vbCrLf & vbCrLf & _
    "ATENÇÃO: Todas as ZLI's e ZEQ's serão enviadas para o banco de dados, o que pode levar um tempo considerável para ser concluído.", vbExclamation + vbYesNo + vbDefaultButton2, "Enviar todos os dados?")

If MsgLoad_All = vbNo Then
    Exit Sub
End If

    LoadAll = "1"
    TempoInicioAll = Now()
    Call LoadToAPI_zli_transmissao

End Sub


Sub LoadToAPI_zli_transmissao()

If Range("Label_NomeLT").Locked = False Then
    MsgNaoImportado = MsgBox("Importe os dados da LT e realize o preenchimento completo antes de enviar para o banco de dados", vbExclamation, "Dados não importados")
    Exit Sub
End If

APIAtual = "ZLI_TRANSMISSAO"
CodLT = Range("Label_CodLT")

If LoadAll = "1" Then
    Unload UserForm_EnviandoAPI
    UserForm_EnviandoAPI.Label1.Caption = "Enviando: " & APIAtual
    UserForm_EnviandoAPI.Show vbModeless
    GoTo Iniciar
End If

Iniciar_zli_transmissao:

Dim MsgLoad_zli_transmissao As VbMsgBoxResult

MsgLoad_zli_transmissao = MsgBox("Deseja enviar os dados para o banco de dados, na API ""zli_transmissao""?" _
    , vbInformation + vbYesNo + vbDefaultButton2, "Enviar dados para ""zli_transmissao""?")

If MsgLoad_zli_transmissao = vbNo Then
    
    If LoadAll = "1" Then
        Dim MsgLoad_zli_transmissao_reask As VbMsgBoxResult

        MsgLoad_zli_transmissao_reask = MsgBox("Deseja mesmo cancelar o envio dos dados para o banco de dados, na API ""zli_transmissao""?" & vbCrLf & vbCrLf & _
            "ATENÇÃO: O envio de todas as ZLI's e ZEQ's pendentes serão canceladas, e as que já foram enviadas serão mantidas.", vbExclamation + vbYesNo + vbDefaultButton2, "Cancelar envio para ""zli_transmissao""?")
        
        If MsgLoad_zli_transmissao_reask = vbNo Then
            GoTo Iniciar_zli_transmissao
        ElseIf MsgLoad_zli_transmissao_reask = vbYes Then
            MsgLoadCancel = MsgBox("Envio de dados para o banco de dados cancelado!", vbCritical, "Cancelado!")
            Exit Sub
        End If
    End If
    
    Exit Sub

End If


Iniciar:

Application.ScreenUpdating = False

TempoInicio = Now()

    ActiveWorkbook.Connections("Consulta - Query_ID_zlis").Refresh
    Unload UserForm_EnviandoAPI

get_d_classificacao_id = Range("Label_ZLI_Transmissao_Classificacao")
get_d_manutencao_id = Range("Label_ZLI_Transmissao_Manutencao")
get_d_operacao_id = Range("Label_ZLI_Transmissao_Operacao")
get_resistencia_aterramento = Range("Label_ZLI_Transmissao_ResistAterramento")
get_temperatura_longa_duracao = Range("Label_ZLI_Transmissao_TemperLongaDur")
get_temperatura_curta_duracao = Range("Label_ZLI_Transmissao_TemperCurtaDur")
get_distancia_min_condutor_solo = Range("Label_ZLI_Transmissao_DistMinimaCondSolo")
get_velocidade_maxima_vento = Range("Label_ZLI_Transmissao_VelocMaximaVento")
get_extensao_propria = Range("Label_ZLI_Transmissao_ExtensPropria")
get_extensao_total_linha = Range("Label_ZLI_Transmissao_ExtensTotalLinha")
get_quantidade_estruturas = Range("Label_ZLI_Transmissao_QtdeEstruturas")
get_modelo_torre_tipica = Range("Label_ZLI_Transmissao_ModeloTorreTipica")


Select Case get_d_classificacao_id
    Case "", "-": d_classificacao_id = Null
    Case "Demais Instalações de Transmissão (DIT)": d_classificacao_id = "DIT"
    Case "Rede Básica (RB)": d_classificacao_id = "RB"
    Case "Rede Básica de Fronteira (RBF)": d_classificacao_id = "RBF"
    Case "Interligação Internacional (II)": d_classificacao_id = "II"
    Case "Consumidor Industrial (CI)": d_classificacao_id = "CI"
    Case "Instalações Comp. por Geradores": d_classificacao_id = "ICG"
    Case "Instalações Exclu. de Geradores": d_classificacao_id = "IEG"
    Case Else: d_classificacao_id = "Erro"
End Select

Select Case get_d_manutencao_id
    Case "", "-": d_manutencao_id = Null
    Case "Própria": d_manutencao_id = "P"
    Case "Terceiros": d_manutencao_id = "T"
    Case Else: d_manutencao_id = "Erro"
End Select

Select Case get_d_operacao_id
    Case "", "-": d_operacao_id = Null
    Case "Própria": d_operacao_id = "P"
    Case "Terceiros": d_operacao_id = "T"
    Case Else: d_operacao_id = "Erro"
End Select

Select Case get_resistencia_aterramento
    Case "", "-": resistencia_aterramento = Null
    Case Else: resistencia_aterramento = get_resistencia_aterramento
End Select

Select Case get_temperatura_longa_duracao
    Case "", "-": temperatura_longa_duracao = Null
    Case Else: temperatura_longa_duracao = get_temperatura_longa_duracao
End Select

Select Case get_temperatura_curta_duracao
    Case "", "-": temperatura_curta_duracao = Null
    Case Else: temperatura_curta_duracao = get_temperatura_curta_duracao
End Select

Select Case get_distancia_min_condutor_solo
    Case "", "-": distancia_min_condutor_solo = Null
    Case Else: distancia_min_condutor_solo = Application.WorksheetFunction.Ceiling(get_distancia_min_condutor_solo, 1)
End Select

Select Case get_velocidade_maxima_vento
    Case "", "-": velocidade_maxima_vento = Null
    Case Else: velocidade_maxima_vento = get_velocidade_maxima_vento
End Select

Select Case get_extensao_propria
    Case "", "-": extensao_propria = Null
    Case Else: extensao_propria = get_extensao_propria
End Select

Select Case get_extensao_total_linha
    Case "", "-": extensao_total_linha = Null
    Case Else: extensao_total_linha = get_extensao_total_linha
End Select

Select Case get_quantidade_estruturas
    Case "", "-": quantidade_estruturas = Null
    Case Else: quantidade_estruturas = get_quantidade_estruturas
End Select

Select Case get_modelo_torre_tipica
    Case "", "-": modelo_torre_tipica = Null
    Case Else: modelo_torre_tipica = get_modelo_torre_tipica
End Select


ID_ZLITransmissao = Range("Query_ID_zlis[id_zli_li_transmissao]").Rows(1).Value
    
    Dim WinHttpReq As Object
    Set WinHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")

If ID_ZLITransmissao <> "" Then
    WinHttpReq.Open "PUT", "http://apilevantamento.h2m.eng.br:3000/api/zli_li_transmissao/" & ID_ZLITransmissao, False
Else:
    WinHttpReq.Open "POST", "http://apilevantamento.h2m.eng.br:3000/api/zli_li_transmissao/", False
End If

    WinHttpReq.SetRequestHeader "Content-Type", "application/json"
    Dim json As Object
    Set json = CreateObject("Scripting.Dictionary")
    
        json("d_classificacao_id") = d_classificacao_id
        json("d_manutencao_id") = d_manutencao_id
        json("d_operacao_id") = d_operacao_id
        json("resistencia_aterramento") = resistencia_aterramento
        json("temperatura_longa_duracao") = temperatura_longa_duracao
        json("temperatura_curta_duracao") = temperatura_curta_duracao
        json("distancia_min_condutor_solo") = distancia_min_condutor_solo
        json("velocidade_maxima_vento") = velocidade_maxima_vento
        json("extensao_propria") = extensao_propria
        json("extensao_total_linha") = extensao_total_linha
        json("quantidade_estruturas") = quantidade_estruturas
        json("modelo_torre_tipica") = modelo_torre_tipica
        json("linha_transmissao_codigo_ativo_concessionaria") = CodLT
    
    Dim jsonData As String
    jsonData = JsonConverter.ConvertToJson(json)
    
    WinHttpReq.Send jsonData
      
If ID_ZLITransmissao <> "" Then
    If InStr(1, WinHttpReq.ResponseText, "EDITADO(A) COM SUCESSO") = 0 Then
        Qtde_Erros = 1
    End If
Else:
    If InStr(1, WinHttpReq.ResponseText, "CRIADO(A) COM SUCESSO") = 0 Then
        Qtde_Erros = 1
    End If
End If

Application.ScreenUpdating = True

TempoFim = Now()
    
    If Qtde_Erros = "" Then
        MsgLog = "Dados de " & APIAtual & " da LT " & CodLT & " enviados para o banco de dados com sucesso!" & Chr(13) & Chr(13) & _
        "Tempo de execução: " & TimeValue(Format(TempoFim - TempoInicio, "dd/mm/yyyy hh:mm:ss"))
        
        If LoadAll = "" Then
            ShowMsgLog = MsgBox(MsgLog, vbInformation, "Sucesso!")
        End If
            IdErro = 0
    Else:
        MsgLog = "Ocorreu um erro ao enviar os dados de " & APIAtual & " da LT " & CodLT & " para o banco de dados." & _
        Chr(13) & Chr(13) & _
        "Revise os dados e tente novamente!" & _
        Chr(13) & Chr(13) & _
        "Tempo de execução: " & TimeValue(Format(TempoFim - TempoInicio, "dd/mm/yyyy hh:mm:ss"))
        
        If LoadAll = "" Then
            ShowMsgLog = MsgBox(MsgLog, vbCritical, "Erro(s)!")
        End If
            IdErro = 1
    
    End If


'Gerando registro de log:

        On Error GoTo AddNewLog
            RowLog = Application.WorksheetFunction.Match(APIAtual, Range("TabLogErros[API]"), 0)
                Range("TabLogErros[API]").Rows(RowLog).Value = APIAtual
                Range("TabLogErros[Erro]").Rows(RowLog).Value = IdErro
                Range("TabLogErros[Início]").Rows(RowLog).Value = TempoInicio
                Range("TabLogErros[Fim]").Rows(RowLog).Value = TempoFim
                Range("TabLogErros[Msg]").Rows(RowLog).Value = MsgLog
            GoTo Finalizar
AddNewLog:
        On Error GoTo -1
        On Error GoTo 0
                Range("TabLogErros[API]").SpecialCells(xlCellTypeBlanks).Rows(1).Value = APIAtual
                Range("TabLogErros[Erro]").SpecialCells(xlCellTypeBlanks).Rows(1).Value = IdErro
                Range("TabLogErros[Início]").SpecialCells(xlCellTypeBlanks).Rows(1).Value = TempoInicio
                Range("TabLogErros[Fim]").SpecialCells(xlCellTypeBlanks).Rows(1).Value = TempoFim
                Range("TabLogErros[Msg]").SpecialCells(xlCellTypeBlanks).Rows(1).Value = MsgLog

Finalizar:
ActiveWorkbook.Save

    On Error GoTo -1
    On Error GoTo 0

If LoadAll = "1" Then
    Call LoadToAPI_zli_parametros_op
End If

End Sub

Sub LoadToAPI_zli_parametros_op()


If Range("Label_NomeLT").Locked = False Then
    MsgNaoImportado = MsgBox("Importe os dados da LT e realize o preenchimento completo antes de enviar para o banco de dados", vbExclamation, "Dados não importados")
    Exit Sub
End If

APIAtual = "ZLI_PARAMETROS_OP"
CodLT = Range("Label_CodLT")

If LoadAll = "1" Then
    Unload UserForm_EnviandoAPI
    UserForm_EnviandoAPI.Label1.Caption = "Enviando: " & APIAtual
    UserForm_EnviandoAPI.Show vbModeless
    GoTo Iniciar
End If

Iniciar_zli_parametros_op:

Dim MsgLoad_zli_parametros_op As VbMsgBoxResult

MsgLoad_zli_parametros_op = MsgBox("Deseja enviar os dados para o banco de dados, na API ""zli_parametros_op""?" _
    , vbInformation + vbYesNo + vbDefaultButton2, "Enviar dados para ""zli_parametros_op""?")

If MsgLoad_zli_parametros_op = vbNo Then
    
    If LoadAll = "1" Then
        Dim MsgLoad_zli_parametros_op_reask As VbMsgBoxResult

        MsgLoad_zli_parametros_op_reask = MsgBox("Deseja mesmo cancelar o envio dos dados para o banco de dados, na API ""zli_parametros_op""?" & vbCrLf & vbCrLf & _
            "ATENÇÃO: O envio de todas as ZLI's e ZEQ's pendentes serão canceladas, e as que já foram enviadas serão mantidas.", vbExclamation + vbYesNo + vbDefaultButton2, "Cancelar envio para ""zli_parametros_op""?")
        
        If MsgLoad_zli_parametros_op_reask = vbNo Then
            GoTo Iniciar_zli_parametros_op
        ElseIf MsgLoad_zli_parametros_op_reask = vbYes Then
            MsgLoadCancel = MsgBox("Envio de dados para o banco de dados cancelado!", vbCritical, "Cancelado!")
            Exit Sub
        End If
    End If
    
    Exit Sub

End If


Iniciar:

Application.ScreenUpdating = False

TempoInicio = Now()

    ActiveWorkbook.Connections("Consulta - Query_ID_zlis").Refresh
    Unload UserForm_EnviandoAPI

get_cap_op_long_duracao_verao_dia = Range("Label_ZLI_ParametrosOp_CapacOperLDVD")
get_cap_op_long_dur_verao_noite = Range("Label_ZLI_ParametrosOp_CapacOperLDVN")
get_cap_op_long_duracao_inver_dia = Range("Label_ZLI_ParametrosOp_CapacOperLDID")
get_cap_op_long_dur_inver_noite = Range("Label_ZLI_ParametrosOp_CapacOperLDIN")
get_cap_op_curt_duracao_verao_dia = Range("Label_ZLI_ParametrosOp_CapacOperCDVD")
get_cap_op_curta_dur_verao_noite = Range("Label_ZLI_ParametrosOp_CapacOperCDVN")
get_cap_op_curta_dur_inverno_dia = Range("Label_ZLI_ParametrosOp_CapacOperCDID")
get_cap_op_curt_dur_inverno_noite = Range("Label_ZLI_ParametrosOp_CapacOperCDIN")
get_flec_max_cond = Range("Label_ZLI_ParametrosOp_FlechaMaxCondut")
get_flecha_max_para_raios = Range("Label_ZLI_ParametrosOp_FlechaMaxPR")
get_resis_sequen_posit_r1 = Range("Label_ZLI_ParametrosOp_ResistSeqPosit")
get_reat_sequen_posit_x1 = Range("Label_ZLI_ParametrosOp_ReatSeqPosit")
get_suscept_seq_posit_b1 = Range("Label_ZLI_ParametrosOp_SuscepSeqPosit")
get_resist_seq_zero_r0 = Range("Label_ZLI_ParametrosOp_ResistSeqZero")
get_reat_seq_zero_x0 = Range("Label_ZLI_ParametrosOp_ReatSeqZero")
get_suscept_seq_zero_b0 = Range("Label_ZLI_ParametrosOp_SuscepSeqZero")
get_capacit_seq_posit_c1 = Range("Label_ZLI_ParametrosOp_CapacitSeqPosit")
get_capacit_seq_zero_c0 = Range("Label_ZLI_ParametrosOp_CapacitSeqZero")
get_impedancia_surto = Range("Label_ZLI_ParametrosOp_ImpedSurto")
get_nivel_ceraunico = Range("Label_ZLI_ParametrosOp_NivelCeraunico")
get_taxa_falhas = Range("Label_ZLI_ParametrosOp_TaxaFalhas")


Select Case get_cap_op_long_duracao_verao_dia
    Case "", "-": cap_op_long_duracao_verao_dia = Null
    Case Else: cap_op_long_duracao_verao_dia = get_cap_op_long_duracao_verao_dia
End Select

Select Case get_cap_op_long_dur_verao_noite
    Case "", "-": cap_op_long_dur_verao_noite = Null
    Case Else: cap_op_long_dur_verao_noite = get_cap_op_long_dur_verao_noite
End Select

Select Case get_cap_op_long_duracao_inver_dia
    Case "", "-": cap_op_long_duracao_inver_dia = Null
    Case Else: cap_op_long_duracao_inver_dia = get_cap_op_long_duracao_inver_dia
End Select

Select Case get_cap_op_long_dur_inver_noite
    Case "", "-": cap_op_long_dur_inver_noite = Null
    Case Else: cap_op_long_dur_inver_noite = get_cap_op_long_dur_inver_noite
End Select

Select Case get_cap_op_curt_duracao_verao_dia
    Case "", "-": cap_op_curt_duracao_verao_dia = Null
    Case Else: cap_op_curt_duracao_verao_dia = get_cap_op_curt_duracao_verao_dia
End Select

Select Case get_cap_op_curta_dur_verao_noite
    Case "", "-": cap_op_curta_dur_verao_noite = Null
    Case Else: cap_op_curta_dur_verao_noite = get_cap_op_curta_dur_verao_noite
End Select

Select Case get_cap_op_curta_dur_inverno_dia
    Case "", "-": cap_op_curta_dur_inverno_dia = Null
    Case Else: cap_op_curta_dur_inverno_dia = get_cap_op_curta_dur_inverno_dia
End Select

Select Case get_cap_op_curt_dur_inverno_noite
    Case "", "-": cap_op_curt_dur_inverno_noite = Null
    Case Else: cap_op_curt_dur_inverno_noite = get_cap_op_curt_dur_inverno_noite
End Select

Select Case get_flec_max_cond
    Case "", "-": flec_max_cond = Null
    Case Else: flec_max_cond = get_flec_max_cond
End Select

Select Case get_flecha_max_para_raios
    Case "", "-": flecha_max_para_raios = Null
    Case Else: flecha_max_para_raios = get_flecha_max_para_raios
End Select

Select Case get_resis_sequen_posit_r1
    Case "", "-": resis_sequen_posit_r1 = Null
    Case Else: resis_sequen_posit_r1 = get_resis_sequen_posit_r1
End Select

Select Case get_reat_sequen_posit_x1
    Case "", "-": reat_sequen_posit_x1 = Null
    Case Else: reat_sequen_posit_x1 = get_reat_sequen_posit_x1
End Select

Select Case get_suscept_seq_posit_b1
    Case "", "-": suscept_seq_posit_b1 = Null
    Case Else: suscept_seq_posit_b1 = get_suscept_seq_posit_b1
End Select

Select Case get_resist_seq_zero_r0
    Case "", "-": resist_seq_zero_r0 = Null
    Case Else: resist_seq_zero_r0 = get_resist_seq_zero_r0
End Select

Select Case get_reat_seq_zero_x0
    Case "", "-": reat_seq_zero_x0 = Null
    Case Else: reat_seq_zero_x0 = get_reat_seq_zero_x0
End Select

Select Case get_suscept_seq_zero_b0
    Case "", "-": suscept_seq_zero_b0 = Null
    Case Else: suscept_seq_zero_b0 = get_suscept_seq_zero_b0
End Select

Select Case get_capacit_seq_posit_c1
    Case "", "-": capacit_seq_posit_c1 = Null
    Case Else: capacit_seq_posit_c1 = get_capacit_seq_posit_c1
End Select

Select Case get_capacit_seq_zero_c0
    Case "", "-": capacit_seq_zero_c0 = Null
    Case Else: capacit_seq_zero_c0 = get_capacit_seq_zero_c0
End Select

Select Case get_impedancia_surto
    Case "", "-": impedancia_surto = Null
    Case Else: impedancia_surto = get_impedancia_surto
End Select

Select Case get_nivel_ceraunico
    Case "", "-": nivel_ceraunico = Null
    Case Else: nivel_ceraunico = get_nivel_ceraunico
End Select

Select Case get_taxa_falhas
    Case "", "-": taxa_falhas = Null
    Case Else: taxa_falhas = get_taxa_falhas
End Select



ID_ZLIParametrosOP = Range("Query_ID_zlis[id_zli_parametros_op]").Rows(1).Value
    
    Dim WinHttpReq As Object
    Set WinHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")

If ID_ZLIParametrosOP <> "" Then
    WinHttpReq.Open "PUT", "http://apilevantamento.h2m.eng.br:3000/api/zli_parametros_op/" & ID_ZLIParametrosOP, False
Else:
    WinHttpReq.Open "POST", "http://apilevantamento.h2m.eng.br:3000/api/zli_parametros_op", False
End If

    WinHttpReq.SetRequestHeader "Content-Type", "application/json"
    Dim json As Object
    Set json = CreateObject("Scripting.Dictionary")
        
        
        json("cap_op_long_duracao_verao_dia") = cap_op_long_duracao_verao_dia
        json("cap_op_long_dur_verao_noite") = cap_op_long_dur_verao_noite
        json("cap_op_long_duracao_inver_dia") = cap_op_long_duracao_inver_dia
        json("cap_op_long_dur_inver_noite") = cap_op_long_dur_inver_noite
        json("cap_op_curt_duracao_verao_dia") = cap_op_curt_duracao_verao_dia
        json("cap_op_curta_dur_verao_noite") = cap_op_curta_dur_verao_noite
        json("cap_op_curta_dur_inverno_dia") = cap_op_curta_dur_inverno_dia
        json("cap_op_curt_dur_inverno_noite") = cap_op_curt_dur_inverno_noite
        json("flec_max_cond") = flec_max_cond
        json("flecha_max_para_raios") = flecha_max_para_raios
        json("resis_sequen_posit_r1") = resis_sequen_posit_r1
        json("reat_sequen_posit_x1") = reat_sequen_posit_x1
        json("suscept_seq_posit_b1") = suscept_seq_posit_b1
        json("resist_seq_zero_r0") = resist_seq_zero_r0
        json("reat_seq_zero_x0") = reat_seq_zero_x0
        json("suscept_seq_zero_b0") = suscept_seq_zero_b0
        json("capacit_seq_posit_c1") = capacit_seq_posit_c1
        json("capacit_seq_zero_c0") = capacit_seq_zero_c0
        json("impedancia_surto") = impedancia_surto
        json("nivel_ceraunico") = nivel_ceraunico
        json("taxa_falhas") = taxa_falhas
        json("linha_transmissao_codigo_ativo_concessionaria") = CodLT
    
    
    Dim jsonData As String
    jsonData = JsonConverter.ConvertToJson(json)
    
    WinHttpReq.Send jsonData
      
If ID_ZLIParametrosOP <> "" Then
    If InStr(1, WinHttpReq.ResponseText, "EDITADO(A) COM SUCESSO") = 0 Then
        Qtde_Erros = 1
    End If
Else:
    If InStr(1, WinHttpReq.ResponseText, "CRIADO(A) COM SUCESSO") = 0 Then
        Qtde_Erros = 1
    End If
End If

Application.ScreenUpdating = True

TempoFim = Now()

    If Qtde_Erros = "" Then
        MsgLog = "Dados de " & APIAtual & " da LT " & CodLT & " enviados para o banco de dados com sucesso!" & Chr(13) & Chr(13) & _
        "Tempo de execução: " & TimeValue(Format(TempoFim - TempoInicio, "dd/mm/yyyy hh:mm:ss"))
        
        If LoadAll = "" Then
            ShowMsgLog = MsgBox(MsgLog, vbInformation, "Sucesso!")
        End If
            IdErro = 0
    
    Else:
        MsgLog = "Ocorreu um erro ao enviar os dados de " & APIAtual & " da LT " & CodLT & " para o banco de dados." & _
        Chr(13) & Chr(13) & _
        "Revise os dados e tente novamente!" & _
        Chr(13) & Chr(13) & _
        "Tempo de execução: " & TimeValue(Format(TempoFim - TempoInicio, "dd/mm/yyyy hh:mm:ss"))
        
        If LoadAll = "" Then
            ShowMsgLog = MsgBox(MsgLog, vbCritical, "Erro(s)!")
        End If
            IdErro = 1
        
    End If


'Gerando registro de log:

        On Error GoTo AddNewLog
            RowLog = Application.WorksheetFunction.Match(APIAtual, Range("TabLogErros[API]"), 0)
                Range("TabLogErros[API]").Rows(RowLog).Value = APIAtual
                Range("TabLogErros[Erro]").Rows(RowLog).Value = IdErro
                Range("TabLogErros[Início]").Rows(RowLog).Value = TempoInicio
                Range("TabLogErros[Fim]").Rows(RowLog).Value = TempoFim
                Range("TabLogErros[Msg]").Rows(RowLog).Value = MsgLog
            GoTo Finalizar
AddNewLog:
        On Error GoTo -1
        On Error GoTo 0
                Range("TabLogErros[API]").SpecialCells(xlCellTypeBlanks).Rows(1).Value = APIAtual
                Range("TabLogErros[Erro]").SpecialCells(xlCellTypeBlanks).Rows(1).Value = IdErro
                Range("TabLogErros[Início]").SpecialCells(xlCellTypeBlanks).Rows(1).Value = TempoInicio
                Range("TabLogErros[Fim]").SpecialCells(xlCellTypeBlanks).Rows(1).Value = TempoFim
                Range("TabLogErros[Msg]").SpecialCells(xlCellTypeBlanks).Rows(1).Value = MsgLog

        
Finalizar:
ActiveWorkbook.Save

    On Error GoTo -1
    On Error GoTo 0

If LoadAll = "1" Then
    Call LoadToAPI_zeq_estru_geral
End If

End Sub

Sub LoadToAPI_zeq_estru_geral()


If Range("Label_NomeLT").Locked = False Then
    MsgNaoImportado = MsgBox("Importe os dados da LT e realize o preenchimento completo antes de enviar para o banco de dados", vbExclamation, "Dados não importados")
    Exit Sub
End If

APIAtual = "ZEQ_ESTRUTURA_GERAL"
CodLT = Range("Label_CodLT")

If LoadAll = "1" Then
    Unload UserForm_EnviandoAPI
    UserForm_EnviandoAPI.Label1.Caption = "Enviando: " & APIAtual
    UserForm_EnviandoAPI.Show vbModeless
    GoTo Iniciar
End If


For Each tbl In ActiveSheet.ListObjects
    If tbl.AutoFilter.FilterMode Then
        MsgBox "Há filtros ativos na planilha." & Chr(13) & Chr(13) & "Remova-o e confira os dados antes de enviar.", vbExclamation
        Exit Sub
    End If
Next tbl


Iniciar_zeq_estru_geral:

Dim MsgLoad_zeq_estru_geral As VbMsgBoxResult

MsgLoad_zeq_estru_geral = MsgBox("Deseja enviar os dados para o banco de dados, na API ""zeq_estru_geral""?" _
    , vbInformation + vbYesNo + vbDefaultButton2, "Enviar dados para ""zeq_estru_geral""?")

If MsgLoad_zeq_estru_geral = vbNo Then
    
    If LoadAll = "1" Then
        Dim MsgLoad_zeq_estru_geral_reask As VbMsgBoxResult

        MsgLoad_zeq_estru_geral_reask = MsgBox("Deseja mesmo cancelar o envio dos dados para o banco de dados, na API ""zeq_estru_geral""?" & vbCrLf & vbCrLf & _
            "ATENÇÃO: O envio de todas as ZLI's e ZEQ's pendentes serão canceladas, e as que já foram enviadas serão mantidas.", vbExclamation + vbYesNo + vbDefaultButton2, "Cancelar envio para ""zeq_estru_geral""?")
        
        If MsgLoad_zeq_estru_geral_reask = vbNo Then
            GoTo Iniciar_zeq_estru_geral
        ElseIf MsgLoad_zeq_estru_geral_reask = vbYes Then
            MsgLoadCancel = MsgBox("Envio de dados para o banco de dados cancelado!", vbCritical, "Cancelado!")
            Exit Sub
        End If
    End If
    
    Exit Sub

End If


Iniciar:

Application.ScreenUpdating = False

TempoInicio = Now()

QtdeTorres = WorksheetFunction.CountA(Range("Tab_zeq_estru_geral[NÚMERO DE OPERAÇÃO]"))

    ActiveWorkbook.Connections("Consulta - Query_ID_zeq_estru_geral").Refresh
    Unload UserForm_EnviandoAPI


Dim get_numero_operacao As String
Dim get_numero_projeto As String
Dim get_silhueta As String
Dim get_d_tipo_estrutura_linha_id As String
Dim get_altura_total As String
Dim get_d_material_construtivo_id As String
Dim get_d_disposicao_fases_id As String
Dim get_desenho_lista_construcao As String
Dim get_desenho_perfil_planta As String
Dim get_menor_distancia_fases_polos As String
Dim get_d_tipo_circuito_id As String
Dim get_altura_misula As String
Dim get_vao_vento As String
Dim get_vao_peso As String
Dim get_comprimento_vao As String
Dim get_distancia_progressiva As String
Dim get_angulo_deflexao As String
Dim get_altitude As String
Dim get_latitude As String
Dim get_longitude As String
Dim get_datum As String


Dim Repete As Integer
Repete = 1

Do While Repete <= QtdeTorres


NumOP = Range("Tab_zeq_estru_geral[NÚMERO DE OPERAÇÃO]").Rows(Repete).text
    
      ActiveWorkbook.Queries.Item("Param_NumOP").Formula = _
        """" & NumOP & """ meta [IsParameterQuery=true, Type=""Text"", IsParameterQueryRequired=true]"
    

sequencia_num_torre = Application.Match(NumOP, Range("Query_ID_zeq_estru_geral[numero_torre]"), 0)

ID = Range("Query_ID_zeq_estru_geral[ID_zeq_estru_geral]").Rows(sequencia_num_torre)

get_numero_operacao = Range("Tab_zeq_estru_geral[NÚMERO DE OPERAÇÃO]").Rows(Repete).text
get_numero_projeto = Range("Tab_zeq_estru_geral[NÚMERO DE PROJETO]").Rows(Repete).text
get_silhueta = Range("Tab_zeq_estru_geral[SILHUETA]").Rows(Repete).text
get_d_tipo_estrutura_linha_id = Range("Tab_zeq_estru_geral[TIPO DE ESTRUTURA DE LINHA]").Rows(Repete).text
get_altura_total = Range("Tab_zeq_estru_geral[ALTURA TOTAL (m)]").Rows(Repete).Value
get_d_material_construtivo_id = Range("Tab_zeq_estru_geral[MATERIAL CONSTRUTIVO]").Rows(Repete).text
get_d_disposicao_fases_id = Range("Tab_zeq_estru_geral[DISPOSIÇÃO DAS FASES]").Rows(Repete).text
get_desenho_lista_construcao = Range("Tab_zeq_estru_geral[DESENHO DA LISTA DE CONSTRUÇÃO]").Rows(Repete).text
get_desenho_perfil_planta = Range("Tab_zeq_estru_geral[DESENHO DO PERFIL E PLANTA]").Rows(Repete).text
get_menor_distancia_fases_polos = Range("Tab_zeq_estru_geral[MENOR DISTÂNCIA FASES/POLOS (m)]").Rows(Repete).Value
get_d_tipo_circuito_id = Range("Tab_zeq_estru_geral[TIPO DE CIRCUITO]").Rows(Repete).text
get_altura_misula = Range("Tab_zeq_estru_geral[ALTURA MISULA (m)]").Rows(Repete).Value
get_vao_vento = Range("Tab_zeq_estru_geral[VÃO DE VENTO (m)]").Rows(Repete).Value
get_vao_peso = Range("Tab_zeq_estru_geral[VÃO DE PESO (m)]").Rows(Repete).Value
get_comprimento_vao = Range("Tab_zeq_estru_geral[COMPRIMENTO DO VÃO (m)]").Rows(Repete).Value
get_distancia_progressiva = Range("Tab_zeq_estru_geral[DISTÂNCIA PROGRESSIVA (m)]").Rows(Repete).Value
get_angulo_deflexao = Range("Tab_zeq_estru_geral[ÂNGULO DEFLEXÃO]").Rows(Repete).text
get_altitude = Range("Tab_zeq_estru_geral[ALTITUDE]").Rows(Repete).Value
get_latitude = Range("Tab_zeq_estru_geral[LATITUDE]").Rows(Repete).text
get_longitude = Range("Tab_zeq_estru_geral[LONGITUDE]").Rows(Repete).text
get_datum = Range("Tab_zeq_estru_geral[DATUM]").Rows(Repete).text


Select Case get_numero_operacao
    Case "", "-": numero_operacao = Null
    Case Else: numero_operacao = get_numero_operacao
End Select

Select Case get_numero_projeto
    Case "", "-": numero_projeto = Null
    Case Else: numero_projeto = get_numero_projeto
End Select

Select Case get_silhueta
    Case "", "-": silhueta = Null
    Case Else: silhueta = get_silhueta
End Select

Select Case get_d_tipo_estrutura_linha_id
    Case "", "-": d_tipo_estrutura_linha_id = Null
    Case "Autoportante": d_tipo_estrutura_linha_id = "A"
    Case "Estaiada": d_tipo_estrutura_linha_id = "E"
    Case Else: d_tipo_estrutura_linha_id = "Erro"
End Select

Select Case get_altura_total
    Case "", "-": altura_total = Null
    Case Else: altura_total = get_altura_total
End Select

Select Case get_d_material_construtivo_id
    Case "", "-": d_material_construtivo_id = Null
    Case "Metálica": d_material_construtivo_id = "MT"
    Case "Compósito": d_material_construtivo_id = "CP"
    Case "Madeira": d_material_construtivo_id = "MD"
    Case "Concreto": d_material_construtivo_id = "CC"
    Case Else: d_material_construtivo_id = "Erro"
End Select

Select Case get_d_disposicao_fases_id
    Case "", "-": d_disposicao_fases_id = Null
    Case "Transposição": d_disposicao_fases_id = "TP"
    Case "Horizontal": d_disposicao_fases_id = "HZ"
    Case "Vertical": d_disposicao_fases_id = "VT"
    Case "Triangular": d_disposicao_fases_id = "TA"
    Case "Trifólio": d_disposicao_fases_id = "TF"
    Case Else: d_disposicao_fases_id = "Erro"
End Select

Select Case get_desenho_lista_construcao
    Case "", "-": desenho_lista_construcao = Null
    Case Else: desenho_lista_construcao = get_desenho_lista_construcao
End Select

Select Case get_desenho_perfil_planta
    Case "", "-": desenho_perfil_planta = Null
    Case Else: desenho_perfil_planta = get_desenho_perfil_planta
End Select

Select Case get_menor_distancia_fases_polos
    Case "", "-": menor_distancia_fases_polos = Null
    Case Else: menor_distancia_fases_polos = get_menor_distancia_fases_polos
End Select

Select Case get_d_tipo_circuito_id
    Case "", "-": d_tipo_circuito_id = Null
    Case "Duplo": d_tipo_circuito_id = "D"
    Case "Simples": d_tipo_circuito_id = "S"
    Case Else: d_tipo_circuito_id = "Erro"
End Select

Select Case get_altura_misula
    Case "", "-": altura_misula = Null
    Case Else: altura_misula = get_altura_misula
End Select

Select Case get_vao_vento
    Case "", "-": vao_vento = Null
    Case Else: vao_vento = get_vao_vento
End Select

Select Case get_vao_peso
    Case "", "-": vao_peso = Null
    Case Else: vao_peso = get_vao_peso
End Select

Select Case get_comprimento_vao
    Case "", "-": comprimento_vao = Null
    Case Else: comprimento_vao = get_comprimento_vao
End Select

Select Case get_distancia_progressiva
    Case "", "-": distancia_progressiva = Null
    Case Else: distancia_progressiva = get_distancia_progressiva
End Select

Select Case get_angulo_deflexao
    Case "", "-", 0: angulo_deflexao = Null
    Case Else: angulo_deflexao = Replace(get_angulo_deflexao, """", ",00""")
End Select

Select Case get_altitude
    Case "", "-": altitude = Null
    Case Else: altitude = get_altitude
End Select

Select Case get_latitude
    Case "", "-": latitude = Null
    Case Else: latitude = get_latitude
End Select

Select Case get_longitude
    Case "", "-": longitude = Null
    Case Else: longitude = get_longitude
End Select

Select Case get_datum
    Case "", "-": datum = Null
    Case Else: datum = get_datum
End Select



    Dim WinHttpReq As Object
    Set WinHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")

    WinHttpReq.Open "PUT", "http://apilevantamento.h2m.eng.br:3000/api/zeq_estru_geral_lt/" & ID, False
    WinHttpReq.SetRequestHeader "Content-Type", "application/json"

    Dim json As Object
    Set json = CreateObject("Scripting.Dictionary")
    
        json("altura_total") = altura_total
        json("d_material_construtivo_id") = d_material_construtivo_id
        json("d_disposicao_fases_id") = d_disposicao_fases_id
        json("desenho_perfil_planta") = desenho_perfil_planta
        json("desenho_lista_construcao") = desenho_lista_construcao
        json("menor_distancia_fases_polos") = menor_distancia_fases_polos
        json("d_tipo_circuito_id") = d_tipo_circuito_id
        json("altura_misula") = altura_misula
        json("vao_vento") = vao_vento
        json("vao_peso") = vao_peso
        json("distancia_progressiva") = distancia_progressiva
        json("comprimento_vao") = comprimento_vao
        json("angulo_deflexao") = angulo_deflexao
        json("silhueta") = silhueta
        json("d_tipo_estrutura_linha_id") = d_tipo_estrutura_linha_id
        json("numero_operacao") = numero_operacao
        json("numero_projeto") = numero_projeto
        json("altitude") = altitude
        json("latitude") = latitude
        json("longitude") = longitude
        json("datum") = datum
    
    Dim jsonData As String
    jsonData = JsonConverter.ConvertToJson(json)
    
    WinHttpReq.Send jsonData
    
    If InStr(1, WinHttpReq.ResponseText, "EDITADO(A) COM SUCESSO") = 0 Then
        
        If Qtde_Erros = "" Then
            Qtde_Erros = 1
        Else: Qtde_Erros = Qtde_Erros + 1
        End If
        
        If OP_Erros = "" Then
            OP_Erros = NumOP
        Else: OP_Erros = OP_Erros & ", " & NumOP
        End If
        
    End If

    Repete = Repete + 1

Set altura_total = Nothing
Set d_material_construtivo_id = Nothing
Set d_disposicao_fases_id = Nothing
Set desenho_perfil_planta = Nothing
Set desenho_lista_construcao = Nothing
Set menor_distancia_fases_polos = Nothing
Set d_tipo_circuito_id = Nothing
Set altura_misula = Nothing
Set vao_vento = Nothing
Set vao_peso = Nothing
Set distancia_progressiva = Nothing
Set comprimento_vao = Nothing
Set angulo_deflexao = Nothing
Set silhueta = Nothing
Set d_tipo_estrutura_linha_id = Nothing
Set numero_operacao = Nothing
Set numero_projeto = Nothing
Set altitude = Nothing
Set latitude = Nothing
Set longitude = Nothing
Set datum = Nothing

Set WinHttpReq = Nothing
Set json = Nothing

Loop

Application.ScreenUpdating = True

TempoFim = Now()
    
    If Qtde_Erros = "" Then
        MsgLog = "Dados de " & APIAtual & " da LT " & CodLT & " enviados para o banco de dados com sucesso!" & Chr(13) & Chr(13) & _
        "Total de torres: " & QtdeTorres & _
        Chr(13) & Chr(13) & _
        "Tempo de execução: " & TimeValue(Format(TempoFim - TempoInicio, "dd/mm/yyyy hh:mm:ss"))

        If LoadAll = "" Then
            ShowMsgLog = MsgBox(MsgLog, vbInformation, "Sucesso!")
        End If
            IdErro = 0
        
    Else:
        MsgLog = Qtde_Erros & " torre(s) não foi(oram) enviada(s): " & Chr(13) & Chr(13) & _
        OP_Erros & Chr(13) & Chr(13) & _
        "Verifique os dados e tente novamente." & _
        Chr(13) & Chr(13) & _
        "Tempo de execução: " & TimeValue(Format(TempoFim - TempoInicio, "dd/mm/yyyy hh:mm:ss"))
    
        If LoadAll = "" Then
            ShowMsgLog = MsgBox(MsgLog, vbCritical, "Erro(s)!")
        End If
            IdErro = 1
        
    End If


'Gerando registro de log:

        On Error GoTo AddNewLog
            RowLog = Application.WorksheetFunction.Match(APIAtual, Range("TabLogErros[API]"), 0)
                Range("TabLogErros[API]").Rows(RowLog).Value = APIAtual
                Range("TabLogErros[Erro]").Rows(RowLog).Value = IdErro
                Range("TabLogErros[Início]").Rows(RowLog).Value = TempoInicio
                Range("TabLogErros[Fim]").Rows(RowLog).Value = TempoFim
                Range("TabLogErros[Msg]").Rows(RowLog).Value = MsgLog
            GoTo Finalizar
AddNewLog:
        On Error GoTo -1
        On Error GoTo 0
                Range("TabLogErros[API]").SpecialCells(xlCellTypeBlanks).Rows(1).Value = APIAtual
                Range("TabLogErros[Erro]").SpecialCells(xlCellTypeBlanks).Rows(1).Value = IdErro
                Range("TabLogErros[Início]").SpecialCells(xlCellTypeBlanks).Rows(1).Value = TempoInicio
                Range("TabLogErros[Fim]").SpecialCells(xlCellTypeBlanks).Rows(1).Value = TempoFim
                Range("TabLogErros[Msg]").SpecialCells(xlCellTypeBlanks).Rows(1).Value = MsgLog

        
Finalizar:
ActiveWorkbook.Save

    On Error GoTo -1
    On Error GoTo 0

If LoadAll = "1" Then
Call LoadToAPI_zeq_estru_autop
End If

End Sub


Sub LoadToAPI_zeq_estru_autop()


If Range("Label_NomeLT").Locked = False Then
    MsgNaoImportado = MsgBox("Importe os dados da LT e realize o preenchimento completo antes de enviar para o banco de dados", vbExclamation, "Dados não importados")
    Exit Sub
End If

APIAtual = "ZEQ_ESTRUTURA_AUTOPORTANTE"
CodLT = Range("Label_CodLT")

If LoadAll = "1" Then
    Unload UserForm_EnviandoAPI
    UserForm_EnviandoAPI.Label1.Caption = "Enviando: " & APIAtual
    UserForm_EnviandoAPI.Show vbModeless
    GoTo Iniciar
End If


For Each tbl In ActiveSheet.ListObjects
    If tbl.AutoFilter.FilterMode Then
        MsgBox "Há filtros ativos na planilha." & Chr(13) & Chr(13) & "Remova-o e confira os dados antes de enviar.", vbExclamation
        Exit Sub
    End If
Next tbl


Iniciar_zeq_estru_autop:

Dim MsgLoad_zeq_estru_autop As VbMsgBoxResult

MsgLoad_zeq_estru_autop = MsgBox("Deseja enviar os dados para o banco de dados, na API ""zeq_estru_autop""?" _
    , vbInformation + vbYesNo + vbDefaultButton2, "Enviar dados para ""zeq_estru_autop""?")

If MsgLoad_zeq_estru_autop = vbNo Then
    
    If LoadAll = "1" Then
        Dim MsgLoad_zeq_estru_autop_reask As VbMsgBoxResult

        MsgLoad_zeq_estru_autop_reask = MsgBox("Deseja mesmo cancelar o envio dos dados para o banco de dados, na API ""zeq_estru_autop""?" & vbCrLf & vbCrLf & _
            "ATENÇÃO: O envio de todas as ZLI's e ZEQ's pendentes serão canceladas, e as que já foram enviadas serão mantidas.", vbExclamation + vbYesNo + vbDefaultButton2, "Cancelar envio para ""zeq_estru_autop""?")
        
        If MsgLoad_zeq_estru_autop_reask = vbNo Then
            GoTo Iniciar_zeq_estru_autop
        ElseIf MsgLoad_zeq_estru_autop_reask = vbYes Then
            MsgLoadCancel = MsgBox("Envio de dados para o banco de dados cancelado!", vbCritical, "Cancelado!")
            Exit Sub
        End If
    End If
    
    Exit Sub

End If


Iniciar:

Application.ScreenUpdating = False

TempoInicio = Now()

QtdeTorres = WorksheetFunction.CountA(Range("Tab_zeq_estru_autop_estai[TORRE]"))

    ActiveWorkbook.Connections("Consulta - Query_ID_zeq_estru_autop").Refresh
    Unload UserForm_EnviandoAPI


Dim get_extensao As Variant
Dim get_altura_perna_a As Variant
Dim get_altura_perna_b As Variant
Dim get_altura_perna_c As Variant
Dim get_altura_perna_d As Variant
Dim get_perna_referencia As String
Dim get_delta_h As Variant
Dim get_d_fundacao_id As String
Dim get_desenho_fundacao As String


Dim Repete As Integer
Repete = 1

Do While Repete <= QtdeTorres


NumOP = Range("Tab_zeq_estru_autop_estai[TORRE]").Rows(Repete).text
    
      ActiveWorkbook.Queries.Item("Param_NumOP").Formula = _
        """" & NumOP & """ meta [IsParameterQuery=true, Type=""Text"", IsParameterQueryRequired=true]"
    

sequencia_num_torre = Application.Match(NumOP, Range("Query_ID_zeq_estru_autop[numero_torre]"), 0)

ID = Range("Query_ID_zeq_estru_autop[ID_zeq_estru_autop]").Rows(sequencia_num_torre)


get_extensao = Range("Tab_zeq_estru_autop_estai[EXTENSÃO (m)]").Rows(Repete).Value
get_altura_perna_a = Range("Tab_zeq_estru_autop_estai[ALTURA PERNA A (m)]").Rows(Repete).Value
get_altura_perna_b = Range("Tab_zeq_estru_autop_estai[ALTURA PERNA B (m)]").Rows(Repete).Value
get_altura_perna_c = Range("Tab_zeq_estru_autop_estai[ALTURA PERNA C (m)]").Rows(Repete).Value
get_altura_perna_d = Range("Tab_zeq_estru_autop_estai[ALTURA PERNA D (m)]").Rows(Repete).Value
get_perna_referencia = Range("Tab_zeq_estru_autop_estai[PERNA DE REFERÊNCIA]").Rows(Repete).text
get_delta_h = Range("Tab_zeq_estru_autop_estai[DELTA H (m)]").Rows(Repete).Value
get_d_fundacao_id = Range("Tab_zeq_estru_autop_estai[FUNDAÇÃO PÉ]").Rows(Repete).text
get_desenho_fundacao = Range("Tab_zeq_estru_autop_estai[DESENHO FUNDAÇÃO PÉ]").Rows(Repete).text


Select Case get_extensao
    Case "", "-": extensao = Null
    Case Else: extensao = get_extensao
End Select

Select Case get_altura_perna_a
    Case "", "-": altura_perna_a = Null
    Case Else: altura_perna_a = get_altura_perna_a
End Select

Select Case get_altura_perna_b
    Case "", "-": altura_perna_b = Null
    Case Else: altura_perna_b = get_altura_perna_b
End Select

Select Case get_altura_perna_c
    Case "", "-": altura_perna_c = Null
    Case Else: altura_perna_c = get_altura_perna_c
End Select

Select Case get_altura_perna_d
    Case "", "-": altura_perna_d = Null
    Case Else: altura_perna_d = get_altura_perna_d
End Select

Select Case get_perna_referencia
    Case "", "-": perna_referencia = Null
    Case "A": perna_referencia = "A"
    Case "B": perna_referencia = "B"
    Case "C": perna_referencia = "C"
    Case "D": perna_referencia = "D"
    Case Else: perna_referencia = "Erro"
End Select

Select Case get_delta_h
    Case "", "-": delta_h = Null
    Case Else: delta_h = get_delta_h
End Select

Select Case get_d_fundacao_id
    Case "", "-": d_fundacao_id = Null
    Case "Engastada": d_fundacao_id = "EG"
    Case "Grelha": d_fundacao_id = "GR"
    Case "Estaqueada": d_fundacao_id = "ET"
    Case "Tubulão": d_fundacao_id = "TB"
    Case "Helicoidal": d_fundacao_id = "HE"
    Case "Sapata": d_fundacao_id = "SP"
    Case Else: d_fundacao_id = "Erro"
End Select

Select Case get_desenho_fundacao
    Case "", "-": desenho_fundacao = Null
    Case Else: desenho_fundacao = get_desenho_fundacao
End Select


    Dim WinHttpReq As Object
    Set WinHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")

    WinHttpReq.Open "PUT", "http://apilevantamento.h2m.eng.br:3000/api/zeq_estru_autop_lt/" & ID, False
    WinHttpReq.SetRequestHeader "Content-Type", "application/json"

    Dim json As Object
    Set json = CreateObject("Scripting.Dictionary")
    
        json("extensao") = extensao
        json("altura_perna_a") = altura_perna_a
        json("altura_perna_b") = altura_perna_b
        json("altura_perna_c") = altura_perna_c
        json("altura_perna_d") = altura_perna_d
        json("perna_referencia") = perna_referencia
        json("delta_h") = delta_h
        json("d_fundacao_id") = d_fundacao_id
        json("desenho_fundacao") = desenho_fundacao
    
    Dim jsonData As String
    jsonData = JsonConverter.ConvertToJson(json)
    
    WinHttpReq.Send jsonData
    
    If InStr(1, WinHttpReq.ResponseText, "EDITADO(A) COM SUCESSO") = 0 Then
        
        If Qtde_Erros = "" Then
            Qtde_Erros = 1
        Else: Qtde_Erros = Qtde_Erros + 1
        End If
        
        If OP_Erros = "" Then
            OP_Erros = NumOP
        Else: OP_Erros = OP_Erros & ", " & NumOP
        End If
        
    End If

    Repete = Repete + 1


Set extensao = Nothing
Set altura_perna_a = Nothing
Set altura_perna_b = Nothing
Set altura_perna_c = Nothing
Set altura_perna_d = Nothing
Set perna_referencia = Nothing
Set delta_h = Nothing
Set d_fundacao_id = Nothing
Set desenho_fundacao = Nothing

Set WinHttpReq = Nothing
Set json = Nothing

Loop

Application.ScreenUpdating = True

TempoFim = Now()

    If Qtde_Erros = "" Then
        MsgLog = "Dados de " & APIAtual & " da LT " & CodLT & " enviados para o banco de dados com sucesso!" & Chr(13) & Chr(13) & _
        "Total de torres: " & QtdeTorres & _
        Chr(13) & Chr(13) & _
        "Tempo de execução: " & TimeValue(Format(TempoFim - TempoInicio, "dd/mm/yyyy hh:mm:ss"))
        
        If LoadAll = "" Then
            ShowMsgLog = MsgBox(MsgLog, vbInformation, "Sucesso!")
        End If
            IdErro = 0
               
    Else:
        MsgLog = Qtde_Erros & " torre(s) não foi(oram) enviada(s): " & Chr(13) & Chr(13) & _
        OP_Erros & Chr(13) & Chr(13) & _
        "Verifique os dados e tente novamente." & _
        Chr(13) & Chr(13) & _
        "Tempo de execução: " & TimeValue(Format(TempoFim - TempoInicio, "dd/mm/yyyy hh:mm:ss"))
        
        If LoadAll = "" Then
            ShowMsgLog = MsgBox(MsgLog, vbCritical, "Erro(s)!")
        End If
            IdErro = 1
    End If


'Gerando registro de log:

        On Error GoTo AddNewLog
            RowLog = Application.WorksheetFunction.Match(APIAtual, Range("TabLogErros[API]"), 0)
                Range("TabLogErros[API]").Rows(RowLog).Value = APIAtual
                Range("TabLogErros[Erro]").Rows(RowLog).Value = IdErro
                Range("TabLogErros[Início]").Rows(RowLog).Value = TempoInicio
                Range("TabLogErros[Fim]").Rows(RowLog).Value = TempoFim
                Range("TabLogErros[Msg]").Rows(RowLog).Value = MsgLog
            GoTo Finalizar
AddNewLog:
        On Error GoTo -1
        On Error GoTo 0
                Range("TabLogErros[API]").SpecialCells(xlCellTypeBlanks).Rows(1).Value = APIAtual
                Range("TabLogErros[Erro]").SpecialCells(xlCellTypeBlanks).Rows(1).Value = IdErro
                Range("TabLogErros[Início]").SpecialCells(xlCellTypeBlanks).Rows(1).Value = TempoInicio
                Range("TabLogErros[Fim]").SpecialCells(xlCellTypeBlanks).Rows(1).Value = TempoFim
                Range("TabLogErros[Msg]").SpecialCells(xlCellTypeBlanks).Rows(1).Value = MsgLog

        
Finalizar:
ActiveWorkbook.Save

    On Error GoTo -1
    On Error GoTo 0

If LoadAll = "1" Then
    Call LoadToAPI_zeq_cadeia_isol
End If

End Sub

Sub LoadToAPI_zeq_estru_estai()

End Sub

Sub LoadToAPI_zeq_cadeia_isol()


If Range("Label_NomeLT").Locked = False Then
    MsgNaoImportado = MsgBox("Importe os dados da LT e realize o preenchimento completo antes de enviar para o banco de dados", vbExclamation, "Dados não importados")
    Exit Sub
End If

APIAtual = "ZEQ_CADEIA_ISOLADORES"
CodLT = Range("Label_CodLT")

If LoadAll = "1" Then
    Unload UserForm_EnviandoAPI
    UserForm_EnviandoAPI.Label1.Caption = "Enviando: " & APIAtual
    UserForm_EnviandoAPI.Show vbModeless
    GoTo Iniciar
End If


For Each tbl In ActiveSheet.ListObjects
    If tbl.AutoFilter.FilterMode Then
        MsgBox "Há filtros ativos na planilha." & Chr(13) & Chr(13) & "Remova-o e confira os dados antes de enviar.", vbExclamation
        Exit Sub
    End If
Next tbl


Iniciar_zeq_cadeia_isol:

Dim MsgLoad_zeq_cadeia_isol As VbMsgBoxResult

MsgLoad_zeq_cadeia_isol = MsgBox("Deseja enviar os dados para o banco de dados, nas APIs ""zeq_cadeia_isol - fases A, B e C""?" _
    , vbInformation + vbYesNo + vbDefaultButton2, "Enviar dados para ""zeq_cadeia_isol""?")

If MsgLoad_zeq_cadeia_isol = vbNo Then
    
    If LoadAll = "1" Then
        Dim MsgLoad_zeq_cadeia_isol_reask As VbMsgBoxResult

        MsgLoad_zeq_cadeia_isol_reask = MsgBox("Deseja mesmo cancelar o envio dos dados para o banco de dados, na APIs ""zeq_cadeia_isol - fases A, B e C""?" & vbCrLf & vbCrLf & _
            "ATENÇÃO: O envio de todas as ZLI's e ZEQ's pendentes serão canceladas, e as que já foram enviadas serão mantidas.", vbExclamation + vbYesNo + vbDefaultButton2, "Cancelar envio para ""zeq_cadeia_isol""?")
        
        If MsgLoad_zeq_cadeia_isol_reask = vbNo Then
            GoTo Iniciar_zeq_cadeia_isol
        ElseIf MsgLoad_zeq_cadeia_isol_reask = vbYes Then
            MsgLoadCancel = MsgBox("Envio de dados para o banco de dados cancelado!", vbCritical, "Cancelado!")
            Exit Sub
        End If
    End If
    
    Exit Sub

End If


Iniciar:

Application.ScreenUpdating = False

TempoInicio = Now()

QtdeTorres = WorksheetFunction.CountA(Range("Tab_zeq_cadeia_isol[TORRE]"))
QtdeArranjosA = WorksheetFunction.CountIf(Range("Tab_zeq_cadeia_isol[FASEAMENTO ELÉTRICO]"), "A")
QtdeArranjosB = WorksheetFunction.CountIf(Range("Tab_zeq_cadeia_isol[FASEAMENTO ELÉTRICO]"), "B")
QtdeArranjosC = WorksheetFunction.CountIf(Range("Tab_zeq_cadeia_isol[FASEAMENTO ELÉTRICO]"), "C")


    ActiveWorkbook.Connections("Consulta - Query_ID_zeq_cadeia_isol").Refresh
    Unload UserForm_EnviandoAPI


Dim get_num_torre As String
Dim get_faseamento_eletrico As String
Dim get_desenho_arranjo As String
Dim get_desenho_isolador As String
Dim get_d_material_isolador_id As String
Dim get_comprimento_cadeia As Variant
Dim get_qtde_total_isolador_arranjo As Variant
Dim get_d_tipo_arranjo_cadeia_id As String
Dim get_d_composicao_arranjo_id As String
Dim get_massa_peso_adicional As Variant


Dim Repete As Integer
Repete = 1

Do While Repete <= QtdeTorres


NumOP = Range("Tab_zeq_cadeia_isol[TORRE]").Rows(Repete).text & "|" & Range("Tab_zeq_cadeia_isol[FASEAMENTO ELÉTRICO]").Rows(Repete).text
    
      ActiveWorkbook.Queries.Item("Param_NumOP").Formula = _
        """" & NumOP & """ meta [IsParameterQuery=true, Type=""Text"", IsParameterQueryRequired=true]"
    

sequencia_num_torre = Application.Match(NumOP, Range("Query_ID_zeq_cadeia_isol[numero_torre|fase_cadeia_isol]"), 0)

ID = Range("Query_ID_zeq_cadeia_isol[ID_zeq_cadeia_isol]").Rows(sequencia_num_torre)

get_num_torre = Range("Tab_zeq_cadeia_isol[TORRE]").Rows(Repete).text
get_faseamento_eletrico = Range("Tab_zeq_cadeia_isol[FASEAMENTO ELÉTRICO]").Rows(Repete).text
get_desenho_arranjo = Range("Tab_zeq_cadeia_isol[DESENHO DO ARRANJO]").Rows(Repete).text
get_desenho_isolador = Range("Tab_zeq_cadeia_isol[DESENHO DO ISOLADOR]").Rows(Repete).text
get_d_material_isolador_id = Range("Tab_zeq_cadeia_isol[MATERIAL DO ISOLADOR]").Rows(Repete).text
get_comprimento_cadeia = Range("Tab_zeq_cadeia_isol[COMPRIMENTO DA CADEIA (m)]").Rows(Repete).Value
get_qtde_total_isolador_arranjo = Range("Tab_zeq_cadeia_isol[QUANTIDADE TOTAL ISOL ARRANJO]").Rows(Repete).Value
get_d_tipo_arranjo_cadeia_id = Range("Tab_zeq_cadeia_isol[TIPO DE ARRANJO DA CADEIA]").Rows(Repete).text
get_d_composicao_arranjo_id = Range("Tab_zeq_cadeia_isol[COMPOSIÇÃO DO ARRANJO]").Rows(Repete).text
get_massa_peso_adicional = Range("Tab_zeq_cadeia_isol[MASSA DO PESO ADICIONAL (kg)]").Rows(Repete).Value


Select Case get_num_torre
    Case "", "-": num_torre = Null
    Case Else: num_torre = get_num_torre
End Select

Select Case get_faseamento_eletrico
    Case "", "-": faseamento_eletrico = Null
    Case "A": faseamento_eletrico = "A"
    Case "B": faseamento_eletrico = "B"
    Case "C": faseamento_eletrico = "C"
    Case Else: faseamento_eletrico = "Erro"
End Select

Select Case get_desenho_arranjo
    Case "", "-": desenho_arranjo = Null
    Case Else: desenho_arranjo = get_desenho_arranjo
End Select

Select Case get_desenho_isolador
    Case "", "-": desenho_isolador = Null
    Case Else: desenho_isolador = get_desenho_isolador
End Select

Select Case get_d_material_isolador_id
    Case "", "-": d_material_isolador_id = Null
    Case "Vidro e Porcelana": d_material_isolador_id = "VP"
    Case "Vidro": d_material_isolador_id = "VI"
    Case "Porcelana": d_material_isolador_id = "PC"
    Case "Polimérico": d_material_isolador_id = "PM"
    Case Else: d_material_isolador_id = "Erro"
End Select

Select Case get_comprimento_cadeia
    Case "", "-": comprimento_cadeia = Null
    Case Else: comprimento_cadeia = get_comprimento_cadeia
End Select

Select Case get_qtde_total_isolador_arranjo
    Case "", "-": qtde_total_isolador_arranjo = Null
    Case Else: qtde_total_isolador_arranjo = get_qtde_total_isolador_arranjo
End Select

Select Case get_d_tipo_arranjo_cadeia_id
    Case "", "-": d_tipo_arranjo_cadeia_id = Null
    Case "Semi Ancoragem": d_tipo_arranjo_cadeia_id = "SA"
    Case "Suspensão I": d_tipo_arranjo_cadeia_id = "SI"
    Case "Suspensão V": d_tipo_arranjo_cadeia_id = "SV"
    Case "Suspensão L": d_tipo_arranjo_cadeia_id = "SL"
    Case "Ancoragem": d_tipo_arranjo_cadeia_id = "AC"
    Case "Ancoragem com Cadeia de Jumper": d_tipo_arranjo_cadeia_id = "AJ"
    Case Else: d_tipo_arranjo_cadeia_id = "Erro"
End Select

Select Case get_d_composicao_arranjo_id
    Case "", "-": d_composicao_arranjo_id = Null
    Case "Simples": d_composicao_arranjo_id = "S"
    Case "Dupla": d_composicao_arranjo_id = "D"
    Case "Tripla": d_composicao_arranjo_id = "T"
    Case "Quadrupla": d_composicao_arranjo_id = "Q"
    Case Else: d_composicao_arranjo_id = "Erro"
End Select

Select Case get_massa_peso_adicional
    Case "", "-": massa_peso_adicional = Null
    Case Else: massa_peso_adicional = get_massa_peso_adicional
End Select



Select Case faseamento_eletrico
    Case "A": API = "zeq_cadeia_isol_lt_fase_1"
    Case "B": API = "zeq_cadeia_isol_lt_fase_2"
    Case "C": API = "zeq_cadeia_isol_lt_fase_3"
End Select

    Dim WinHttpReq As Object
    Set WinHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
    WinHttpReq.Open "PUT", "http://apilevantamento.h2m.eng.br:3000/api/" & API & "/" & ID, False
    WinHttpReq.SetRequestHeader "Content-Type", "application/json"

    Dim json As Object
    Set json = CreateObject("Scripting.Dictionary")
    
        json("faseamento_eletrico") = faseamento_eletrico
        json("desenho_arranjo") = desenho_arranjo
        json("desenho_isolador") = desenho_isolador
        json("d_material_isolador_id") = d_material_isolador_id
        json("comprimento_cadeia") = comprimento_cadeia
        json("qtde_total_isolador_arranjo") = qtde_total_isolador_arranjo
        json("d_tipo_arranjo_cadeia_id") = d_tipo_arranjo_cadeia_id
        json("d_composicao_arranjo_id") = d_composicao_arranjo_id
        json("massa_peso_adicional") = massa_peso_adicional
    
    Dim jsonData As String
    jsonData = JsonConverter.ConvertToJson(json)
    
    WinHttpReq.Send jsonData
    
    If InStr(1, WinHttpReq.ResponseText, " COM SUCESSO") = 0 Then
        
        If Qtde_Erros = "" Then
            Qtde_Erros = 1
        Else: Qtde_Erros = Qtde_Erros + 1
        End If
        
        If OP_Erros = "" Then
            OP_Erros = num_torre & "(" & faseamento_eletrico & ")"
        Else: OP_Erros = OP_Erros & ", " & num_torre & "(" & faseamento_eletrico & ")"
        End If
        
    End If

    Repete = Repete + 1


Set num_torre = Nothing
Set faseamento_eletrico = Nothing
Set desenho_arranjo = Nothing
Set desenho_isolador = Nothing
Set d_material_isolador_id = Nothing
Set comprimento_cadeia = Nothing
Set qtde_total_isolador_arranjo = Nothing
Set d_tipo_arranjo_cadeia_id = Nothing
Set d_composicao_arranjo_id = Nothing
Set massa_peso_adicional = Nothing
Set API = Nothing

Set WinHttpReq = Nothing
Set json = Nothing

Loop

Application.ScreenUpdating = True

TempoFim = Now()

    If Qtde_Erros = "" Then
        MsgLog = "Dados de " & APIAtual & " da LT " & CodLT & " enviados para o banco de dados com sucesso!" & Chr(13) & Chr(13) & _
        "Total de arranjos: " & Chr(13) & Chr(13) & _
        "Fase A: " & QtdeArranjosA & Chr(13) & _
        "Fase B: " & QtdeArranjosB & Chr(13) & _
        "Fase C: " & QtdeArranjosC & Chr(13) & _
        Chr(13) & Chr(13) & _
        "Tempo de execução: " & TimeValue(Format(TempoFim - TempoInicio, "dd/mm/yyyy hh:mm:ss"))
        
        If LoadAll = "" Then
            ShowMsgLog = MsgBox(MsgLog, vbInformation, "Sucesso!")
        End If
            IdErro = 0
            
    Else:
        MsgLog = Qtde_Erros & " arranjos(s) não foi(oram) enviado(s): " & Chr(13) & Chr(13) & _
        OP_Erros & Chr(13) & Chr(13) & _
        "Verifique os dados e tente novamente." & _
        Chr(13) & Chr(13) & _
        "Tempo de execução: " & TimeValue(Format(TempoFim - TempoInicio, "dd/mm/yyyy hh:mm:ss"))
    
        If LoadAll = "" Then
            ShowMsgLog = MsgBox(MsgLog, vbCritical, "Erro(s)!")
        End If
            IdErro = 1
            
    End If


'Gerando registro de log:

        On Error GoTo AddNewLog
            RowLog = Application.WorksheetFunction.Match(APIAtual, Range("TabLogErros[API]"), 0)
                Range("TabLogErros[API]").Rows(RowLog).Value = APIAtual
                Range("TabLogErros[Erro]").Rows(RowLog).Value = IdErro
                Range("TabLogErros[Início]").Rows(RowLog).Value = TempoInicio
                Range("TabLogErros[Fim]").Rows(RowLog).Value = TempoFim
                Range("TabLogErros[Msg]").Rows(RowLog).Value = MsgLog
            GoTo Finalizar
AddNewLog:
        On Error GoTo -1
        On Error GoTo 0
                Range("TabLogErros[API]").SpecialCells(xlCellTypeBlanks).Rows(1).Value = APIAtual
                Range("TabLogErros[Erro]").SpecialCells(xlCellTypeBlanks).Rows(1).Value = IdErro
                Range("TabLogErros[Início]").SpecialCells(xlCellTypeBlanks).Rows(1).Value = TempoInicio
                Range("TabLogErros[Fim]").SpecialCells(xlCellTypeBlanks).Rows(1).Value = TempoFim
                Range("TabLogErros[Msg]").SpecialCells(xlCellTypeBlanks).Rows(1).Value = MsgLog

        
Finalizar:
ActiveWorkbook.Save

    On Error GoTo -1
    On Error GoTo 0

If LoadAll = "1" Then
    Call LoadToAPI_zeq_aterramento
End If

End Sub


Sub LoadToAPI_zeq_aterramento()


If Range("Label_NomeLT").Locked = False Then
    MsgNaoImportado = MsgBox("Importe os dados da LT e realize o preenchimento completo antes de enviar para o banco de dados", vbExclamation, "Dados não importados")
    Exit Sub
End If

APIAtual = "ZEQ_ATERRAMENTO"
CodLT = Range("Label_CodLT")

If LoadAll = "1" Then
    Unload UserForm_EnviandoAPI
    UserForm_EnviandoAPI.Label1.Caption = "Enviando: " & APIAtual
    UserForm_EnviandoAPI.Show vbModeless
    GoTo Iniciar
End If


For Each tbl In ActiveSheet.ListObjects
    If tbl.AutoFilter.FilterMode Then
        MsgBox "Há filtros ativos na planilha." & Chr(13) & Chr(13) & "Remova-o e confira os dados antes de enviar.", vbExclamation
        Exit Sub
    End If
Next tbl


Iniciar_zeq_aterramento:

Dim MsgLoad_zeq_aterramento As VbMsgBoxResult

MsgLoad_zeq_aterramento = MsgBox("Deseja enviar os dados para o banco de dados, na API ""zeq_aterramento""?" _
    , vbInformation + vbYesNo + vbDefaultButton2, "Enviar dados para ""zeq_aterramento""?")

If MsgLoad_zeq_aterramento = vbNo Then
    
    If LoadAll = "1" Then
        Dim MsgLoad_zeq_aterramento_reask As VbMsgBoxResult

        MsgLoad_zeq_aterramento_reask = MsgBox("Deseja mesmo cancelar o envio dos dados para o banco de dados, na API ""zeq_aterramento""?" & vbCrLf & vbCrLf & _
            "ATENÇÃO: O envio de todas as ZLI's e ZEQ's pendentes serão canceladas, e as que já foram enviadas serão mantidas.", vbExclamation + vbYesNo + vbDefaultButton2, "Cancelar envio para ""zeq_aterramento""?")
        
        If MsgLoad_zeq_aterramento_reask = vbNo Then
            GoTo Iniciar_zeq_aterramento
        ElseIf MsgLoad_zeq_aterramento_reask = vbYes Then
            MsgLoadCancel = MsgBox("Envio de dados para o banco de dados cancelado!", vbCritical, "Cancelado!")
            Exit Sub
        End If
    End If
    
    Exit Sub

End If


Iniciar:

Application.ScreenUpdating = False

TempoInicio = Now()

QtdeTorres = WorksheetFunction.CountA(Range("Tab_zeq_aterramento[TORRE]"))

    ActiveWorkbook.Connections("Consulta - Query_ID_zeq_aterramento").Refresh
    Unload UserForm_EnviandoAPI


Dim get_configuracao_aterramento As String
Dim get_tipo_cabo_contrapeso As String
Dim get_comp_tot_cabo_contrapeso As Variant
Dim get_tipo_fio_interligacao As String
Dim get_comp_total_fio_interligacao As Variant
Dim get_resistividade_solo As Variant
Dim get_classificacao_solo As String

Dim Repete As Integer
Repete = 1

Do While Repete <= QtdeTorres


NumOP = Range("Tab_zeq_aterramento[TORRE]").Rows(Repete).text
    
      ActiveWorkbook.Queries.Item("Param_NumOP").Formula = _
        """" & NumOP & """ meta [IsParameterQuery=true, Type=""Text"", IsParameterQueryRequired=true]"
    

sequencia_num_torre = Application.Match(NumOP, Range("Query_ID_zeq_aterramento[numero_torre]"), 0)

ID = Range("Query_ID_zeq_aterramento[ID_zeq_aterramento]").Rows(sequencia_num_torre)


get_configuracao_aterramento = Range("Tab_zeq_aterramento[CONFIGURAÇÃO DE ATERRAMENTO]").Rows(Repete).text
get_tipo_cabo_contrapeso = Range("Tab_zeq_aterramento[TIPO DE CABO CONTRAPESO]").Rows(Repete).text
get_comp_tot_cabo_contrapeso = Range("Tab_zeq_aterramento[COMP TOT CABO CONTRAPESO (m)]").Rows(Repete).Value
get_tipo_fio_interligacao = Range("Tab_zeq_aterramento[TIPO DE FIO DE INTERLIGAÇÃO]").Rows(Repete).text
get_comp_total_fio_interligacao = Range("Tab_zeq_aterramento[COMP TOT FIO INTERLIGACAO (m)]").Rows(Repete).Value
get_resistividade_solo = Range("Tab_zeq_aterramento[RESISTIVIDADE DO SOLO ('#m)]").Rows(Repete).Value
get_classificacao_solo = Range("Tab_zeq_aterramento[CLASSIFICAÇÃO DO SOLO]").Rows(Repete).text


Select Case get_configuracao_aterramento
    Case "", "-": configuracao_aterramento = Null
    Case Else: configuracao_aterramento = get_configuracao_aterramento
End Select

Select Case get_tipo_cabo_contrapeso
    Case "", "-": tipo_cabo_contrapeso = Null
    Case Else: tipo_cabo_contrapeso = get_tipo_cabo_contrapeso
End Select

Select Case get_comp_tot_cabo_contrapeso
    Case "", "-": comp_tot_cabo_contrapeso = Null
    Case Else: comp_tot_cabo_contrapeso = get_comp_tot_cabo_contrapeso
End Select

Select Case get_tipo_fio_interligacao
    Case "", "-": tipo_fio_interligacao = Null
    Case Else: tipo_fio_interligacao = get_tipo_fio_interligacao
End Select

Select Case get_comp_total_fio_interligacao
    Case "", "-": comp_total_fio_interligacao = Null
    Case Else: comp_total_fio_interligacao = get_comp_total_fio_interligacao
End Select

Select Case get_resistividade_solo
    Case "", "-": resistividade_solo = Null
    Case Else: resistividade_solo = get_resistividade_solo
End Select

Select Case get_classificacao_solo
    Case "", "-": classificacao_solo = Null
    Case Else: classificacao_solo = get_classificacao_solo
End Select


    Dim WinHttpReq As Object
    Set WinHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")

    WinHttpReq.Open "PUT", "http://apilevantamento.h2m.eng.br:3000/api/zeq_aterramento_lt/" & ID, False
    WinHttpReq.SetRequestHeader "Content-Type", "application/json"

    Dim json As Object
    Set json = CreateObject("Scripting.Dictionary")
    
        json("configuracao_aterramento") = configuracao_aterramento
        json("tipo_cabo_contrapeso") = tipo_cabo_contrapeso
        json("comp_tot_cabo_contrapeso") = comp_tot_cabo_contrapeso
        json("tipo_fio_interligacao") = tipo_fio_interligacao
        json("comp_total_fio_interligacao") = comp_total_fio_interligacao
        json("resistividade_solo") = resistividade_solo
        json("classificacao_solo") = classificacao_solo
    
    Dim jsonData As String
    jsonData = JsonConverter.ConvertToJson(json)

    WinHttpReq.Send jsonData
    
    If InStr(1, WinHttpReq.ResponseText, "EDITADO(A) COM SUCESSO") = 0 Then
        
        If Qtde_Erros = "" Then
            Qtde_Erros = 1
        Else: Qtde_Erros = Qtde_Erros + 1
        End If
        
        If OP_Erros = "" Then
            OP_Erros = NumOP
        Else: OP_Erros = OP_Erros & ", " & NumOP
        End If
        
    End If

    Repete = Repete + 1

Set configuracao_aterramento = Nothing
Set d_tipo_cabo_contrapeso_id = Nothing
Set comp_tot_cabo_contrapeso = Nothing
Set d_tipo_fio_interligacao_id = Nothing
Set comp_total_fio_interligacao = Nothing
Set resistividade_solo = Nothing
Set classificacao_solo = Nothing

Set WinHttpReq = Nothing
Set json = Nothing

Loop

Application.ScreenUpdating = True

TempoFim = Now()
    
    If Qtde_Erros = "" Then
        MsgLog = "Dados de " & APIAtual & " da LT " & CodLT & " enviados para o banco de dados com sucesso!" & Chr(13) & Chr(13) & _
        "Total de torres: " & QtdeTorres & _
        Chr(13) & Chr(13) & _
        "Tempo de execução: " & TimeValue(Format(TempoFim - TempoInicio, "dd/mm/yyyy hh:mm:ss"))
        
        If LoadAll = "" Then
            ShowMsgLog = MsgBox(MsgLog, vbInformation, "Sucesso!")
        End If
            IdErro = 0
            
    Else:
        MsgLog = Qtde_Erros & " aterramento(s) de torre(s) não foi(oram) enviada(s): " & Chr(13) & Chr(13) & _
        OP_Erros & Chr(13) & Chr(13) & _
        "Verifique os dados e tente novamente." & _
        Chr(13) & Chr(13) & _
        "Tempo de execução: " & TimeValue(Format(TempoFim - TempoInicio, "dd/mm/yyyy hh:mm:ss"))
        
        If LoadAll = "" Then
            ShowMsgLog = MsgBox(MsgLog, vbCritical, "Erro(s)!")
        End If
            IdErro = 1
            
    End If


'Gerando registro de log:

        On Error GoTo AddNewLog
            RowLog = Application.WorksheetFunction.Match(APIAtual, Range("TabLogErros[API]"), 0)
                Range("TabLogErros[API]").Rows(RowLog).Value = APIAtual
                Range("TabLogErros[Erro]").Rows(RowLog).Value = IdErro
                Range("TabLogErros[Início]").Rows(RowLog).Value = TempoInicio
                Range("TabLogErros[Fim]").Rows(RowLog).Value = TempoFim
                Range("TabLogErros[Msg]").Rows(RowLog).Value = MsgLog
            GoTo Finalizar
AddNewLog:
        On Error GoTo -1
        On Error GoTo 0
                Range("TabLogErros[API]").SpecialCells(xlCellTypeBlanks).Rows(1).Value = APIAtual
                Range("TabLogErros[Erro]").SpecialCells(xlCellTypeBlanks).Rows(1).Value = IdErro
                Range("TabLogErros[Início]").SpecialCells(xlCellTypeBlanks).Rows(1).Value = TempoInicio
                Range("TabLogErros[Fim]").SpecialCells(xlCellTypeBlanks).Rows(1).Value = TempoFim
                Range("TabLogErros[Msg]").SpecialCells(xlCellTypeBlanks).Rows(1).Value = MsgLog


Finalizar:
ActiveWorkbook.Save

    On Error GoTo -1
    On Error GoTo 0

If LoadAll = "1" Then
    Call LoadToAPI_zeq_condutor
End If

End Sub

Sub LoadToAPI_zeq_acessos()


If Range("Label_NomeLT").Locked = False Then
    MsgNaoImportado = MsgBox("Importe os dados da LT e realize o preenchimento completo antes de enviar para o banco de dados", vbExclamation, "Dados não importados")
    Exit Sub
End If

For Each tbl In ActiveSheet.ListObjects
    If tbl.AutoFilter.FilterMode Then
        MsgBox "Há filtros ativos na planilha." & Chr(13) & Chr(13) & "Remova-o e confira os dados antes de enviar.", vbExclamation
        Exit Sub
    End If
Next tbl


MsgDesabilitado = MsgBox("Envio de dados ""ZEQ_ACESSOS"" desabilitado.", vbCritical, "Desabilitado")

End Sub


Sub LoadToAPI_zeq_condutor()


If Range("Label_NomeLT").Locked = False Then
    MsgNaoImportado = MsgBox("Importe os dados da LT e realize o preenchimento completo antes de enviar para o banco de dados", vbExclamation, "Dados não importados")
    Exit Sub
End If

APIAtual = "ZEQ_CONDUTOR"
CodLT = Range("Label_CodLT")

If LoadAll = "1" Then
    Unload UserForm_EnviandoAPI
    UserForm_EnviandoAPI.Label1.Caption = "Enviando: " & APIAtual
    UserForm_EnviandoAPI.Show vbModeless
    GoTo Iniciar
End If


For Each tbl In ActiveSheet.ListObjects
    If tbl.AutoFilter.FilterMode Then
        MsgBox "Há filtros ativos na planilha." & Chr(13) & Chr(13) & "Remova-o e confira os dados antes de enviar.", vbExclamation
        Exit Sub
    End If
Next tbl


Iniciar_zeq_condutor:

Dim MsgLoad_zeq_condutor As VbMsgBoxResult

MsgLoad_zeq_condutor = MsgBox("Deseja enviar os dados para o banco de dados, nas APIs ""zeq_condutor - fases A, B e C""?" _
    , vbInformation + vbYesNo + vbDefaultButton2, "Enviar dados para ""zeq_condutor""?")

If MsgLoad_zeq_condutor = vbNo Then
    
    If LoadAll = "1" Then
        Dim MsgLoad_zeq_condutor_reask As VbMsgBoxResult

        MsgLoad_zeq_condutor_reask = MsgBox("Deseja mesmo cancelar o envio dos dados para o banco de dados, na APIs ""zeq_condutor - fases A, B e C""?" & vbCrLf & vbCrLf & _
            "ATENÇÃO: O envio de todas as ZLI's e ZEQ's pendentes serão canceladas, e as que já foram enviadas serão mantidas.", vbExclamation + vbYesNo + vbDefaultButton2, "Cancelar envio para ""zeq_condutor""?")
        
        If MsgLoad_zeq_condutor_reask = vbNo Then
            GoTo Iniciar_zeq_condutor
        ElseIf MsgLoad_zeq_condutor_reask = vbYes Then
            MsgLoadCancel = MsgBox("Envio de dados para o banco de dados cancelado!", vbCritical, "Cancelado!")
            Exit Sub
        End If
    End If
    
    Exit Sub

End If


Iniciar:

Application.ScreenUpdating = False

TempoInicio = Now()

QtdeVaos = WorksheetFunction.CountA(Range("Tab_zeq_condutor[VÃO]"))
QtdeVaoA = WorksheetFunction.CountIf(Range("Tab_zeq_condutor[FASE]"), "A")
QtdeVaoB = WorksheetFunction.CountIf(Range("Tab_zeq_condutor[FASE]"), "B")
QtdeVaoC = WorksheetFunction.CountIf(Range("Tab_zeq_condutor[FASE]"), "C")


    ActiveWorkbook.Connections("Consulta - Query_ID_zeq_condutor").Refresh
    Unload UserForm_EnviandoAPI


Dim get_identificacao_vao As String
Dim get_fase_condutor As String
Dim get_tipo_cabo_condutor_1 As String
Dim get_tipo_cabo_condutor_2 As String
Dim get_quantidade_sub_condutores As Variant
Dim get_tracao_eds As Variant
Dim get_cabo_solo_aneel As Variant
Dim get_circuitos_compartilhados As String
Dim get_quantidade_amortecedores As Variant
Dim get_d_zeq_tipo_amortecedor_id As String
Dim get_quantidade_espacadores As Variant
Dim get_d_tipo_espacador_id As String
Dim get_d_zeq_tipo_emenda_1_id As String
Dim get_d_zeq_tipo_emenda_2_id As String

Dim Repete As Integer
Repete = 1

Do While Repete <= QtdeVaos


NumVaoID = Range("Tab_zeq_condutor[VÃO]").Rows(Repete).text & "|" & Range("Tab_zeq_condutor[FASE]").Rows(Repete).text

sequencia_vao = Application.Match(NumVaoID, Range("Query_ID_zeq_condutor[identificacao_vao|fase_condutor]"), 0)

ID = Range("Query_ID_zeq_condutor[ID_zeq_condutor]").Rows(sequencia_vao)

get_identificacao_vao = Range("Tab_zeq_condutor[VÃO]").Rows(Repete).text
get_fase_condutor = Range("Tab_zeq_condutor[FASE]").Rows(Repete).text
get_tipo_cabo_condutor_1 = Range("Tab_zeq_condutor[TIPO CABO CONDUTOR I]").Rows(Repete).text
get_tipo_cabo_condutor_2 = Range("Tab_zeq_condutor[TIPO CABO CONDUTOR II]").Rows(Repete).text
get_quantidade_sub_condutores = Range("Tab_zeq_condutor[QTDE. SUB-CONDUTORES]").Rows(Repete).Value
get_tracao_eds = Range("Tab_zeq_condutor[TRAÇÃO EDS (%)]").Rows(Repete).Value
get_cabo_solo_aneel = Range("Tab_zeq_condutor[CABO-SOLO ANEEL]").Rows(Repete).Value
get_circuitos_compartilhados = Range("Tab_zeq_condutor[CIRCUITOS COMPARTILHADOS]").Rows(Repete).text
get_quantidade_amortecedores = Range("Tab_zeq_condutor[QUANTIDADE AMORTECEDORES]").Rows(Repete).Value
get_d_zeq_tipo_amortecedor_id = Range("Tab_zeq_condutor[TIPO AMORTECEDOR]").Rows(Repete).text
get_quantidade_espacadores = Range("Tab_zeq_condutor[QUANTIDADE DE ESPAÇADORES]").Rows(Repete).Value
get_d_tipo_espacador_id = Range("Tab_zeq_condutor[TIPO ESPACADOR]").Rows(Repete).text
get_d_zeq_tipo_emenda_1_id = Range("Tab_zeq_condutor[TIPO DE EMENDA I]").Rows(Repete).text
get_d_zeq_tipo_emenda_2_id = Range("Tab_zeq_condutor[TIPO DE EMENDA II]").Rows(Repete).text


Select Case get_identificacao_vao
    Case "", "-": identificacao_vao = Null
    Case Else: identificacao_vao = get_identificacao_vao
End Select

Select Case get_fase_condutor
    Case "", "-": fase_condutor = Null
    Case Else: fase_condutor = get_fase_condutor
End Select

Select Case get_tipo_cabo_condutor_1
    Case "", "-": tipo_cabo_condutor_1 = Null
    Case Else: tipo_cabo_condutor_1 = get_tipo_cabo_condutor_1
End Select

Select Case get_tipo_cabo_condutor_2
    Case "", "-": tipo_cabo_condutor_2 = Null
    Case Else: tipo_cabo_condutor_2 = get_tipo_cabo_condutor_2
End Select

Select Case get_quantidade_sub_condutores
    Case "", "-": quantidade_sub_condutores = Null
    Case Else: quantidade_sub_condutores = get_quantidade_sub_condutores
End Select

Select Case get_tracao_eds
    Case "", "-": tracao_eds = Null
    Case Else: tracao_eds = get_tracao_eds
End Select

Select Case get_cabo_solo_aneel
    Case "", "-": cabo_solo_aneel = Null
    Case Else: cabo_solo_aneel = get_cabo_solo_aneel
End Select

Select Case get_circuitos_compartilhados
    Case "", "-": circuitos_compartilhados = Null
    Case Else: circuitos_compartilhados = get_circuitos_compartilhados
End Select

Select Case get_quantidade_amortecedores
    Case "", "-": quantidade_amortecedores = Null
    Case Else: quantidade_amortecedores = get_quantidade_amortecedores
End Select

Select Case get_d_zeq_tipo_amortecedor_id
    Case "", "-": d_zeq_tipo_amortecedor_id = Null
    Case "Outros": d_zeq_tipo_amortecedor_id = "O"
    Case "Stockbridge": d_zeq_tipo_amortecedor_id = "S"
    Case "Helicoidal": d_zeq_tipo_amortecedor_id = "H"
    Case Else: d_zeq_tipo_amortecedor_id = "Erro"
End Select

Select Case get_quantidade_espacadores
    Case "", "-": quantidade_espacadores = Null
    Case Else: quantidade_espacadores = get_quantidade_espacadores
End Select

Select Case get_d_tipo_espacador_id
    Case "", "-": d_tipo_espacador_id = Null
    Case "Parafuso": d_tipo_espacador_id = "PA"
    Case "Pré-Formado": d_tipo_espacador_id = "PF"
    Case Else: d_tipo_espacador_id = "Erro"
End Select

Select Case get_d_zeq_tipo_emenda_1_id
    Case "", "-": d_zeq_tipo_emenda_1_id = Null
    Case "Não possui": d_zeq_tipo_emenda_1_id = "N"
    Case "Compressão": d_zeq_tipo_emenda_1_id = "C"
    Case "Pré-Formada": d_zeq_tipo_emenda_1_id = "P"
    Case Else: d_zeq_tipo_emenda_1_id = "Erro"
End Select

Select Case get_d_zeq_tipo_emenda_2_id
    Case "", "-": d_zeq_tipo_emenda_2_id = Null
    Case "Não possui": d_zeq_tipo_emenda_2_id = "N"
    Case "Compressão": d_zeq_tipo_emenda_2_id = "C"
    Case "Pré-Formada": d_zeq_tipo_emenda_2_id = "P"
    Case Else: d_zeq_tipo_emenda_2_id = "Erro"
End Select


Select Case get_fase_condutor
    Case "A": API = "zeq_condutor_lt_fase_1"
    Case "B": API = "zeq_condutor_lt_fase_2"
    Case "C": API = "zeq_condutor_lt_fase_3"
End Select


    Dim WinHttpReq As Object
    Set WinHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
    WinHttpReq.Open "PUT", "http://apilevantamento.h2m.eng.br:3000/api/" & API & "/" & ID, False
    WinHttpReq.SetRequestHeader "Content-Type", "application/json"

    Dim json As Object
    Set json = CreateObject("Scripting.Dictionary")
    
        json("tipo_cabo_condutor_1") = tipo_cabo_condutor_1
        json("tipo_cabo_condutor_2") = tipo_cabo_condutor_2
        json("quantidade_sub_condutores") = quantidade_sub_condutores
        json("tracao_eds") = tracao_eds
        json("cabo_solo_aneel") = cabo_solo_aneel
        json("circuitos_compartilhados") = circuitos_compartilhados
        json("quantidade_amortecedores") = quantidade_amortecedores
        json("d_zeq_tipo_amortecedor_id") = d_zeq_tipo_amortecedor_id
        json("quantidade_espacadores") = quantidade_espacadores
        json("d_tipo_espacador_id") = d_tipo_espacador_id
        json("d_zeq_tipo_emenda_1_id") = d_zeq_tipo_emenda_1_id
        json("d_zeq_tipo_emenda_2_id") = d_zeq_tipo_emenda_2_id
    
    
    Dim jsonData As String
    jsonData = JsonConverter.ConvertToJson(json)
    
    WinHttpReq.Send jsonData
    
    If InStr(1, WinHttpReq.ResponseText, " COM SUCESSO") = 0 Then
        
        If Qtde_Erros = "" Then
            Qtde_Erros = 1
        Else: Qtde_Erros = Qtde_Erros + 1
        End If
        
        If Vaos_Erros = "" Then
            Vaos_Erros = identificacao_vao & "(" & fase_condutor & ")"
        Else: Vaos_Erros = Vaos_Erros & ", " & identificacao_vao & "(" & fase_condutor & ")"
        End If
        
    End If

    Repete = Repete + 1


Set identificacao_vao = Nothing
Set fase_condutor = Nothing
Set tipo_cabo_condutor_1 = Nothing
Set tipo_cabo_condutor_2 = Nothing
Set quantidade_sub_condutores = Nothing
Set tracao_eds = Nothing
Set cabo_solo_aneel = Nothing
Set circuitos_compartilhados = Nothing
Set quantidade_amortecedores = Nothing
Set d_zeq_tipo_amortecedor_id = Nothing
Set quantidade_espacadores = Nothing
Set d_tipo_espacador_id = Nothing
Set d_zeq_tipo_emenda_1_id = Nothing
Set d_zeq_tipo_emenda_2_id = Nothing
Set API = Nothing

Set WinHttpReq = Nothing
Set json = Nothing

Loop

Application.ScreenUpdating = True

TempoFim = Now()

    If Qtde_Erros = "" Then
        MsgLog = "Dados de " & APIAtual & " da LT " & CodLT & " enviados para o banco de dados com sucesso!" & Chr(13) & Chr(13) & _
        "Total de vãos: " & Chr(13) & Chr(13) & _
        "Fase A: " & QtdeVaoA & Chr(13) & _
        "Fase B: " & QtdeVaoB & Chr(13) & _
        "Fase C: " & QtdeVaoC & Chr(13) & _
        Chr(13) & _
        "Tempo de execução: " & TimeValue(Format(TempoFim - TempoInicio, "dd/mm/yyyy hh:mm:ss"))
        
        If LoadAll = "" Then
            ShowMsgLog = MsgBox(MsgLog, vbInformation, "Sucesso!")
        End If
            IdErro = 0
            
    Else:
        MsgLog = Qtde_Erros & " arranjos(s) não foi(oram) enviado(s): " & Chr(13) & Chr(13) & _
        Vaos_Erros & Chr(13) & Chr(13) & _
        "Verifique os dados e tente novamente." & _
        Chr(13) & Chr(13) & _
        "Tempo de execução: " & TimeValue(Format(TempoFim - TempoInicio, "dd/mm/yyyy hh:mm:ss"))
        
        If LoadAll = "" Then
            ShowMsgLog = MsgBox(MsgLog, vbCritical, "Erro(s)!")
        End If
            IdErro = 1
            
    End If


'Gerando registro de log:

        On Error GoTo AddNewLog
            RowLog = Application.WorksheetFunction.Match(APIAtual, Range("TabLogErros[API]"), 0)
                Range("TabLogErros[API]").Rows(RowLog).Value = APIAtual
                Range("TabLogErros[Erro]").Rows(RowLog).Value = IdErro
                Range("TabLogErros[Início]").Rows(RowLog).Value = TempoInicio
                Range("TabLogErros[Fim]").Rows(RowLog).Value = TempoFim
                Range("TabLogErros[Msg]").Rows(RowLog).Value = MsgLog
            GoTo Finalizar
AddNewLog:
        On Error GoTo -1
        On Error GoTo 0
                Range("TabLogErros[API]").SpecialCells(xlCellTypeBlanks).Rows(1).Value = APIAtual
                Range("TabLogErros[Erro]").SpecialCells(xlCellTypeBlanks).Rows(1).Value = IdErro
                Range("TabLogErros[Início]").SpecialCells(xlCellTypeBlanks).Rows(1).Value = TempoInicio
                Range("TabLogErros[Fim]").SpecialCells(xlCellTypeBlanks).Rows(1).Value = TempoFim
                Range("TabLogErros[Msg]").SpecialCells(xlCellTypeBlanks).Rows(1).Value = MsgLog

        
Finalizar:
ActiveWorkbook.Save

    On Error GoTo -1
    On Error GoTo 0

If LoadAll = "1" Then
    Call LoadToAPI_zeq_pararaio
End If

End Sub


Sub LoadToAPI_zeq_pararaio()


If Range("Label_NomeLT").Locked = False Then
    MsgNaoImportado = MsgBox("Importe os dados da LT e realize o preenchimento completo antes de enviar para o banco de dados", vbExclamation, "Dados não importados")
    Exit Sub
End If

APIAtual = "ZEQ_PARARAIO"
CodLT = Range("Label_CodLT")

If LoadAll = "1" Then
    Unload UserForm_EnviandoAPI
    UserForm_EnviandoAPI.Label1.Caption = "Enviando: " & APIAtual
    UserForm_EnviandoAPI.Show vbModeless
    GoTo Iniciar
End If


For Each tbl In ActiveSheet.ListObjects
    If tbl.AutoFilter.FilterMode Then
        MsgBox "Há filtros ativos na planilha." & Chr(13) & Chr(13) & "Remova-o e confira os dados antes de enviar.", vbExclamation
        Exit Sub
    End If
Next tbl


Iniciar_zeq_pararaio:

Dim MsgLoad_zeq_pararaio As VbMsgBoxResult

MsgLoad_zeq_pararaio = MsgBox("Deseja enviar os dados para o banco de dados, nas APIs ""zeq_pararaio - esquerdo/central e direito""?" _
    , vbInformation + vbYesNo + vbDefaultButton2, "Enviar dados para ""zeq_pararaio""?")

If MsgLoad_zeq_pararaio = vbNo Then
    
    If LoadAll = "1" Then
        Dim MsgLoad_zeq_pararaio_reask As VbMsgBoxResult

        MsgLoad_zeq_pararaio_reask = MsgBox("Deseja mesmo cancelar o envio dos dados para o banco de dados, na APIs ""zeq_pararaio - esquerdo/central e direito""?" & vbCrLf & vbCrLf & _
            "ATENÇÃO: O envio de todas as ZLI's e ZEQ's pendentes serão canceladas, e as que já foram enviadas serão mantidas.", vbExclamation + vbYesNo + vbDefaultButton2, "Cancelar envio para ""zeq_pararaio""?")
        
        If MsgLoad_zeq_pararaio_reask = vbNo Then
            GoTo Iniciar_zeq_pararaio
        ElseIf MsgLoad_zeq_pararaio_reask = vbYes Then
            MsgLoadCancel = MsgBox("Envio de dados para o banco de dados cancelado!", vbCritical, "Cancelado!")
            Exit Sub
        End If
    End If
    
    Exit Sub

End If


Iniciar:

Application.ScreenUpdating = False

TempoInicio = Now()

QtdeVaos = WorksheetFunction.CountA(Range("Tab_zeq_pararaio[VÃO]"))
QtdeVaoEsq = WorksheetFunction.CountIf(Range("Tab_zeq_pararaio[LADO]"), "Esquerdo/Central")
QtdeVaoDir = WorksheetFunction.CountIf(Range("Tab_zeq_pararaio[LADO]"), "Direito")
QtdeVaoIndef = WorksheetFunction.CountIf(Range("Tab_zeq_pararaio[LADO]"), "Indefinido")


    ActiveWorkbook.Connections("Consulta - Query_ID_zeq_pararaio").Refresh
    Unload UserForm_EnviandoAPI

Dim get_identificacao_vao As String
Dim get_lado_pararaio As String
Dim get_tipo_cabo As String
Dim get_d_tipo_arranjo_cabo_id As String
Dim get_desenho_arranjo As String
Dim get_quantidade_amortecedores As Variant
Dim get_d_tipo_amortecedor_id As String
Dim get_quantidade_esfera As Variant
Dim get_d_engate_esfera_id As String
Dim get_d_tipo_emenda_id As String
Dim get_tracao_eds As Variant
Dim get_para_raio_isolados As String

Dim Repete As Integer
Repete = 1

Do While Repete <= QtdeVaos

    If Range("Tab_zeq_pararaio[LADO]").Rows(Repete).text = "Indefinido" Then
        GoTo nextLoop
    End If

NumVaoID = Range("Tab_zeq_pararaio[VÃO]").Rows(Repete).text & "|" & Range("Tab_zeq_pararaio[LADO]").Rows(Repete).text

sequencia_vao = Application.Match(NumVaoID, Range("Query_ID_zeq_pararaio[identificacao_vao|lado_pararaio]"), 0)

ID = Range("Query_ID_zeq_pararaio[ID_zeq_pararaio]").Rows(sequencia_vao)

get_identificacao_vao = Range("Tab_zeq_pararaio[VÃO]").Rows(Repete).text
get_lado_pararaio = Range("Tab_zeq_pararaio[LADO]").Rows(Repete).text
get_tipo_cabo = Range("Tab_zeq_pararaio[TIPO DO CABO]").Rows(Repete).text
get_d_tipo_arranjo_cabo_id = Range("Tab_zeq_pararaio[TIPO DE ARRANJO DO CABO]").Rows(Repete).text
get_desenho_arranjo = Range("Tab_zeq_pararaio[DESENHO DO ARRANJO]").Rows(Repete).text
get_quantidade_amortecedores = Range("Tab_zeq_pararaio[QUANTIDADE AMORTECEDORES]").Rows(Repete).Value
get_d_tipo_amortecedor_id = Range("Tab_zeq_pararaio[TIPO AMORTECEDOR]").Rows(Repete).text
get_quantidade_esfera = Range("Tab_zeq_pararaio[QUANTIDADE DE ESFERAS]").Rows(Repete).Value
get_d_engate_esfera_id = Range("Tab_zeq_pararaio[ENGATE DA ESFERA]").Rows(Repete).text
get_d_tipo_emenda_id = Range("Tab_zeq_pararaio[TIPO DE EMENDA]").Rows(Repete).text
get_tracao_eds = Range("Tab_zeq_pararaio[TRAÇÃO EDS (%)]").Rows(Repete).Value
get_para_raio_isolados = Range("Tab_zeq_pararaio[PARA-RAIO ISOLADOS]").Rows(Repete).text


Select Case get_identificacao_vao
    Case "", "-": identificacao_vao = Null
    Case Else: identificacao_vao = get_identificacao_vao
End Select

Select Case get_lado_pararaio
    Case "", "-": lado_pararaio = Null
    Case Else: lado_pararaio = get_lado_pararaio
End Select
    
Select Case get_tipo_cabo
    Case "", "-": tipo_cabo = Null
    Case Else: tipo_cabo = get_tipo_cabo
End Select

Select Case get_d_tipo_arranjo_cabo_id
    Case "", "-": d_tipo_arranjo_cabo_id = Null
    Case "Ancoragem": d_tipo_arranjo_cabo_id = "A"
    Case "Suspensão": d_tipo_arranjo_cabo_id = "S"
    Case Else: d_tipo_arranjo_cabo_id = "Erro"
End Select

Select Case get_desenho_arranjo
    Case "", "-": desenho_arranjo = Null
    Case Else: desenho_arranjo = get_desenho_arranjo
End Select

Select Case get_quantidade_amortecedores
    Case "", "-": quantidade_amortecedores = Null
    Case Else: quantidade_amortecedores = get_quantidade_amortecedores
End Select

Select Case get_d_tipo_amortecedor_id
    Case "", "-": d_tipo_amortecedor_id = Null
    Case "Outros": d_tipo_amortecedor_id = "O"
    Case "Stockbridge": d_tipo_amortecedor_id = "S"
    Case "Helicoidal": d_tipo_amortecedor_id = "H"
    Case Else: d_tipo_amortecedor_id = "Erro"
End Select

Select Case get_quantidade_esfera
    Case "", "-": quantidade_esfera = Null
    Case Else: quantidade_esfera = get_quantidade_esfera
End Select

Select Case get_d_engate_esfera_id
    Case "", "-": d_engate_esfera_id = Null
    Case "Via helicóptero": d_engate_esfera_id = "VH"
    Case "Via corda": d_engate_esfera_id = "VC"
    Case "Parafuso": d_engate_esfera_id = "PA"
    Case "Pré-Formado": d_engate_esfera_id = "PF"
    Case Else: d_engate_esfera_id = "Erro"
End Select

Select Case get_d_tipo_emenda_id
    Case "", "-": d_tipo_emenda_id = Null
    Case "Não possui": d_tipo_emenda_id = "N"
    Case "Compressão": d_tipo_emenda_id = "C"
    Case "Pré-Formada": d_tipo_emenda_id = "F"
    Case Else: d_tipo_emenda_id = "Erro"
End Select

Select Case get_tracao_eds
    Case "", "-": tracao_eds = Null
    Case Else: tracao_eds = get_tracao_eds
End Select

Select Case get_para_raio_isolados
    Case "", "-": para_raio_isolados = Null
    Case "Sim": para_raio_isolados = "S"
    Case "Não": para_raio_isolados = "N"
    Case Else: para_raio_isolados = "Erro"
End Select


Select Case get_lado_pararaio
    Case "Esquerdo/Central": API = "zeq_pararaio_lt_esquerdo"
    Case "Direito": API = "zeq_pararaio_lt_direito"
    Case "Indefinido": API = "zeq_pararaio_lt_indefinido"
End Select


    Dim WinHttpReq As Object
    Set WinHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
    WinHttpReq.Open "PUT", "http://apilevantamento.h2m.eng.br:3000/api/" & API & "/" & ID, False
    WinHttpReq.SetRequestHeader "Content-Type", "application/json"

    Dim json As Object
    Set json = CreateObject("Scripting.Dictionary")
    
        json("d_tipo_cabo_id") = d_tipo_cabo_id
        json("d_tipo_arranjo_cabo_id") = d_tipo_arranjo_cabo_id
        json("desenho_arranjo") = desenho_arranjo
        json("quantidade_amortecedores") = quantidade_amortecedores
        json("d_tipo_amortecedor_id") = d_tipo_amortecedor_id
        json("quantidade_esfera") = quantidade_esfera
        json("d_engate_esfera_id") = d_engate_esfera_id
        json("d_tipo_emenda_id") = d_tipo_emenda_id
        json("tracao_eds") = tracao_eds
        json("para_raio_isolados") = para_raio_isolados
    
    
    Dim jsonData As String
    jsonData = JsonConverter.ConvertToJson(json)
        
    WinHttpReq.Send jsonData
    
    If InStr(1, WinHttpReq.ResponseText, " COM SUCESSO") = 0 Then
        
        If Qtde_Erros = "" Then
            Qtde_Erros = 1
        Else: Qtde_Erros = Qtde_Erros + 1
        End If
        
        If lado_pararaio = "Esquerdo/Central" Then
            rsm_lado_pararaio = "Esq/Cen"
        ElseIf lado_pararaio = "Direito" Then
            rsm_lado_pararaio = "Dir"
        ElseIf lado_pararaio = "Indefinido" Then
            rsm_lado_pararaio = "Indef"
        End If
        
        If Vaos_Erros = "" Then
            Vaos_Erros = identificacao_vao & "(" & rsm_lado_pararaio & ")"
        Else: Vaos_Erros = Vaos_Erros & ", " & identificacao_vao & "(" & rsm_lado_pararaio & ")"
        End If
        
    End If

nextLoop:

    Repete = Repete + 1

Set identificacao_vao = Nothing
Set lado_pararaio = Nothing
Set d_tipo_cabo_id = Nothing
Set d_tipo_arranjo_cabo_id = Nothing
Set desenho_arranjo = Nothing
Set quantidade_amortecedores = Nothing
Set d_tipo_amortecedor_id = Nothing
Set quantidade_esfera = Nothing
Set d_engate_esfera_id = Nothing
Set d_tipo_emenda_id = Nothing
Set tracao_eds = Nothing
Set para_raio_isolados = Nothing
Set API = Nothing

Set WinHttpReq = Nothing
Set json = Nothing

Loop

Application.ScreenUpdating = True

TempoFim = Now()
    
    If Qtde_Erros = "" Then
        MsgLog = "Dados de " & APIAtual & " da LT " & CodLT & " enviados para o banco de dados com sucesso!" & Chr(13) & Chr(13) & _
        "Total de vãos: " & Chr(13) & Chr(13) & _
        "Para-raios Esq./Cent.: " & QtdeVaoEsq & Chr(13) & _
        "Para-raios Direito: " & QtdeVaoDir & Chr(13) & _
        Chr(13) & Chr(13) & _
        "Tempo de execução: " & TimeValue(Format(TempoFim - TempoInicio, "dd/mm/yyyy hh:mm:ss"))
    
        If LoadAll = "" Then
            ShowMsgLog = MsgBox(MsgLog, vbInformation, "Sucesso!")
        End If
            IdErro = 0
            
    Else:
        MsgLog = Qtde_Erros & " arranjos(s) não foi(oram) enviado(s): " & Chr(13) & Chr(13) & _
        Vaos_Erros & Chr(13) & Chr(13) & _
        "Verifique os dados e tente novamente." & _
        Chr(13) & Chr(13) & _
        "Tempo de execução: " & TimeValue(Format(TempoFim - TempoInicio, "dd/mm/yyyy hh:mm:ss"))
    
        If LoadAll = "" Then
            ShowMsgLog = MsgBox(MsgLog, vbCritical, "Erro(s)!")
        End If
            IdErro = 1
            
    End If


'Gerando registro de log:

        On Error GoTo AddNewLog
            RowLog = Application.WorksheetFunction.Match(APIAtual, Range("TabLogErros[API]"), 0)
                Range("TabLogErros[API]").Rows(RowLog).Value = APIAtual
                Range("TabLogErros[Erro]").Rows(RowLog).Value = IdErro
                Range("TabLogErros[Início]").Rows(RowLog).Value = TempoInicio
                Range("TabLogErros[Fim]").Rows(RowLog).Value = TempoFim
                Range("TabLogErros[Msg]").Rows(RowLog).Value = MsgLog
            GoTo Finalizar
AddNewLog:
        On Error GoTo -1
        On Error GoTo 0
                Range("TabLogErros[API]").SpecialCells(xlCellTypeBlanks).Rows(1).Value = APIAtual
                Range("TabLogErros[Erro]").SpecialCells(xlCellTypeBlanks).Rows(1).Value = IdErro
                Range("TabLogErros[Início]").SpecialCells(xlCellTypeBlanks).Rows(1).Value = TempoInicio
                Range("TabLogErros[Fim]").SpecialCells(xlCellTypeBlanks).Rows(1).Value = TempoFim
                Range("TabLogErros[Msg]").SpecialCells(xlCellTypeBlanks).Rows(1).Value = MsgLog

        
Finalizar:
ActiveWorkbook.Save

    On Error GoTo -1
    On Error GoTo 0

If LoadAll = "1" Then
    Call LoadToAPI_zeq_opgw
End If

End Sub


Sub LoadToAPI_zeq_opgw()


If Range("Label_NomeLT").Locked = False Then
    MsgNaoImportado = MsgBox("Importe os dados da LT e realize o preenchimento completo antes de enviar para o banco de dados", vbExclamation, "Dados não importados")
    Exit Sub
End If

APIAtual = "ZEQ_OPGW"
CodLT = Range("Label_CodLT")

If LoadAll = "1" Then
    Unload UserForm_EnviandoAPI
    UserForm_EnviandoAPI.Label1.Caption = "Enviando: " & APIAtual
    UserForm_EnviandoAPI.Show vbModeless
    GoTo Iniciar
End If


For Each tbl In ActiveSheet.ListObjects
    If tbl.AutoFilter.FilterMode Then
        MsgBox "Há filtros ativos na planilha." & Chr(13) & Chr(13) & "Remova-o e confira os dados antes de enviar.", vbExclamation
        Exit Sub
    End If
Next tbl


Iniciar_zeq_opgw:

Dim MsgLoad_zeq_opgw As VbMsgBoxResult

MsgLoad_zeq_opgw = MsgBox("Deseja enviar os dados para o banco de dados, nas APIs ""zeq_opgw - esquerdo/central e direito""?" _
    , vbInformation + vbYesNo + vbDefaultButton2, "Enviar dados para ""zeq_opgw""?")

If MsgLoad_zeq_opgw = vbNo Then
    
    If LoadAll = "1" Then
        Dim MsgLoad_zeq_opgw_reask As VbMsgBoxResult

        MsgLoad_zeq_opgw_reask = MsgBox("Deseja mesmo cancelar o envio dos dados para o banco de dados, na APIs ""zeq_opgw - esquerdo/central e direito""?" & vbCrLf & vbCrLf & _
            "ATENÇÃO: O envio de todas as ZLI's e ZEQ's pendentes serão canceladas, e as que já foram enviadas serão mantidas.", vbExclamation + vbYesNo + vbDefaultButton2, "Cancelar envio para ""zeq_opgw""?")
        
        If MsgLoad_zeq_opgw_reask = vbNo Then
            GoTo Iniciar_zeq_opgw
        ElseIf MsgLoad_zeq_opgw_reask = vbYes Then
            MsgLoadCancel = MsgBox("Envio de dados para o banco de dados cancelado!", vbCritical, "Cancelado!")
            Exit Sub
        End If
    End If
    
    Exit Sub

End If


Iniciar:

Application.ScreenUpdating = False

TempoInicio = Now()

QtdeVaos = WorksheetFunction.CountA(Range("Tab_zeq_opgw[VÃO]"))
QtdeVaoEsq = WorksheetFunction.CountIf(Range("Tab_zeq_opgw[LADO]"), "Esquerdo/Central")
QtdeVaoDir = WorksheetFunction.CountIf(Range("Tab_zeq_opgw[LADO]"), "Direito")
QtdeVaoIndef = WorksheetFunction.CountIf(Range("Tab_zeq_opgw[LADO]"), "Indefinido")


    ActiveWorkbook.Connections("Consulta - Query_ID_zeq_opgw").Refresh
    Unload UserForm_EnviandoAPI

Dim get_identificacao_vao As String
Dim get_lado_opgw As String
Dim get_d_fabricante_opgw_id As String
Dim get_numero_fibras As Variant
Dim get_secao As Variant
Dim get_diametro As Variant
Dim get_d_zeq_sim_nao_caixa_emenda_id As String

Dim Repete As Integer
Repete = 1

Do While Repete <= QtdeVaos

    If Range("Tab_zeq_opgw[LADO]").Rows(Repete).text = "Indefinido" Then
        GoTo nextLoop
    End If

NumVaoID = Range("Tab_zeq_opgw[VÃO]").Rows(Repete).text & "|" & Range("Tab_zeq_opgw[LADO]").Rows(Repete).text

sequencia_vao = Application.Match(NumVaoID, Range("Query_ID_zeq_opgw[identificacao_vao|lado_opgw]"), 0)

ID = Range("Query_ID_zeq_opgw[ID_zeq_opgw]").Rows(sequencia_vao)


get_identificacao_vao = Range("Tab_zeq_opgw[VÃO]").Rows(Repete).text
get_lado_opgw = Range("Tab_zeq_opgw[LADO]").Rows(Repete).text
get_d_fabricante_opgw_id = Range("Tab_zeq_opgw[FABRICANTE OPGW]").Rows(Repete).text
get_numero_fibras = Range("Tab_zeq_opgw[NÚMERO DE FIBRAS]").Rows(Repete).Value
get_secao = Range("Tab_zeq_opgw[SEÇÃO (mm²)]").Rows(Repete).Value
get_diametro = Range("Tab_zeq_opgw[DIÂMETRO (mm)]").Rows(Repete).Value
get_d_zeq_sim_nao_caixa_emenda_id = Range("Tab_zeq_opgw[CAIXA DE EMENDA]").Rows(Repete).text


Select Case get_identificacao_vao
    Case "", "-": identificacao_vao = Null
    Case Else: identificacao_vao = get_identificacao_vao
End Select

Select Case get_lado_opgw
    Case "", "-": lado_opgw = Null
    Case Else: lado_opgw = get_lado_opgw
End Select

Select Case get_d_fabricante_opgw_id
    Case "", "-": d_fabricante_opgw_id = Null
    Case "AFL": d_fabricante_opgw_id = 10
    Case "BRUGG": d_fabricante_opgw_id = 20
    Case "FICAP": d_fabricante_opgw_id = 30
    Case "FUJIKURA": d_fabricante_opgw_id = 40
    Case "FURUKAWA": d_fabricante_opgw_id = 50
    Case "PFI": d_fabricante_opgw_id = 60
    Case "PIRELLI": d_fabricante_opgw_id = 70
    Case "PRYSMIAN": d_fabricante_opgw_id = 71
    Case "SUMITOMO": d_fabricante_opgw_id = 80
    Case "TELCON": d_fabricante_opgw_id = 90
    Case Else: d_fabricante_opgw_id = "Erro"
End Select

Select Case get_numero_fibras
    Case "", "-": numero_fibras = Null
    Case Else: numero_fibras = get_numero_fibras
End Select

Select Case get_secao
    Case "", "-": secao = Null
    Case Else: secao = get_secao
End Select

Select Case get_diametro
    Case "", "-": diametro = Null
    Case Else: diametro = get_diametro
End Select

Select Case get_d_zeq_sim_nao_caixa_emenda_id
    Case "", "-": d_zeq_sim_nao_caixa_emenda_id = Null
    Case "Sim": d_zeq_sim_nao_caixa_emenda_id = "S"
    Case "Não": d_zeq_sim_nao_caixa_emenda_id = "N"
    Case Else: d_zeq_sim_nao_caixa_emenda_id = "Erro"
End Select


Select Case get_lado_opgw
    Case "Esquerdo/Central": API = "zeq_opgw_lt_esquerdo"
    Case "Direito": API = "zeq_opgw_lt_direito"
    Case "Indefinido": API = "zeq_opgw_lt_indefinido"
End Select


    Dim WinHttpReq As Object
    Set WinHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
    WinHttpReq.Open "PUT", "http://apilevantamento.h2m.eng.br:3000/api/" & API & "/" & ID, False
    WinHttpReq.SetRequestHeader "Content-Type", "application/json"

    Dim json As Object
    Set json = CreateObject("Scripting.Dictionary")
    
        json("d_fabricante_opgw_id") = d_fabricante_opgw_id
        json("numero_fibras") = numero_fibras
        json("secao") = secao
        json("diametro") = diametro
        json("d_zeq_sim_nao_caixa_emenda_id") = d_zeq_sim_nao_caixa_emenda_id
    
   
    Dim jsonData As String
    jsonData = JsonConverter.ConvertToJson(json)
        
    WinHttpReq.Send jsonData
    
    If InStr(1, WinHttpReq.ResponseText, " COM SUCESSO") = 0 Then
        
        If Qtde_Erros = "" Then
            Qtde_Erros = 1
        Else: Qtde_Erros = Qtde_Erros + 1
        End If
        
        If lado_opgw = "Esquerdo/Central" Then
            rsm_lado_opgw = "Esq/Cen"
        ElseIf lado_opgw = "Direito" Then
            rsm_lado_opgw = "Dir"
        ElseIf lado_opgw = "Indefinido" Then
            rsm_lado_opgw = "Indef"
        End If
        
        If Vaos_Erros = "" Then
            Vaos_Erros = identificacao_vao & "(" & rsm_lado_opgw & ")"
        Else: Vaos_Erros = Vaos_Erros & ", " & identificacao_vao & "(" & rsm_lado_opgw & ")"
        End If
        
    End If

nextLoop:

    Repete = Repete + 1

Set identificacao_vao = Nothing
Set lado_opgw = Nothing
Set d_fabricante_opgw_id = Nothing
Set numero_fibras = Nothing
Set secao = Nothing
Set diametro = Nothing
Set d_zeq_sim_nao_caixa_emenda_id = Nothing
Set API = Nothing

Set WinHttpReq = Nothing
Set json = Nothing

Loop

Application.ScreenUpdating = True

TempoFim = Now()

    If Qtde_Erros = "" Then
        MsgLog = "Dados de " & APIAtual & " da LT " & CodLT & " enviados para o banco de dados com sucesso!" & Chr(13) & Chr(13) & _
        "Total de vãos: " & Chr(13) & Chr(13) & _
        "OPGW Esq./Cent.: " & QtdeVaoEsq & Chr(13) & _
        "OPGW Direito: " & QtdeVaoDir & Chr(13) & _
        Chr(13) & Chr(13) & _
        "Tempo de execução: " & TimeValue(Format(TempoFim - TempoInicio, "dd/mm/yyyy hh:mm:ss"))
    
        If LoadAll = "" Then
            ShowMsgLog = MsgBox(MsgLog, vbInformation, "Sucesso!")
        End If
            IdErro = 0
            
    Else:
        MsgLog = Qtde_Erros & " vão(s) de cabo OPGW não foi(oram) enviado(s): " & Chr(13) & Chr(13) & _
        Vaos_Erros & Chr(13) & Chr(13) & _
        "Verifique os dados e tente novamente." & _
        Chr(13) & Chr(13) & _
        "Tempo de execução: " & TimeValue(Format(TempoFim - TempoInicio, "dd/mm/yyyy hh:mm:ss"))
        
        If LoadAll = "" Then
            ShowMsgLog = MsgBox(MsgLog, vbCritical, "Erro(s)!")
        End If
            IdErro = 1
            
    End If
    

'Gerando registro de log:

        On Error GoTo AddNewLog
            RowLog = Application.WorksheetFunction.Match(APIAtual, Range("TabLogErros[API]"), 0)
                Range("TabLogErros[API]").Rows(RowLog).Value = APIAtual
                Range("TabLogErros[Erro]").Rows(RowLog).Value = IdErro
                Range("TabLogErros[Início]").Rows(RowLog).Value = TempoInicio
                Range("TabLogErros[Fim]").Rows(RowLog).Value = TempoFim
                Range("TabLogErros[Msg]").Rows(RowLog).Value = MsgLog
            GoTo Finalizar
AddNewLog:
        On Error GoTo -1
        On Error GoTo 0
                Range("TabLogErros[API]").SpecialCells(xlCellTypeBlanks).Rows(1).Value = APIAtual
                Range("TabLogErros[Erro]").SpecialCells(xlCellTypeBlanks).Rows(1).Value = IdErro
                Range("TabLogErros[Início]").SpecialCells(xlCellTypeBlanks).Rows(1).Value = TempoInicio
                Range("TabLogErros[Fim]").SpecialCells(xlCellTypeBlanks).Rows(1).Value = TempoFim
                Range("TabLogErros[Msg]").SpecialCells(xlCellTypeBlanks).Rows(1).Value = MsgLog

        
Finalizar:
ActiveWorkbook.Save

    On Error GoTo -1
    On Error GoTo 0

If LoadAll = "1" Then
    Call LoadToAPI_zeq_servidao
End If

End Sub


Sub LoadToAPI_zeq_servidao()


If Range("Label_NomeLT").Locked = False Then
    MsgNaoImportado = MsgBox("Importe os dados da LT e realize o preenchimento completo antes de enviar para o banco de dados", vbExclamation, "Dados não importados")
    Exit Sub
End If

APIAtual = "ZEQ_SERVIDAO"
CodLT = Range("Label_CodLT")

If LoadAll = "1" Then
    Unload UserForm_EnviandoAPI
    UserForm_EnviandoAPI.Label1.Caption = "Enviando: " & APIAtual
    UserForm_EnviandoAPI.Show vbModeless
    GoTo Iniciar
End If


For Each tbl In ActiveSheet.ListObjects
    If tbl.AutoFilter.FilterMode Then
        MsgBox "Há filtros ativos na planilha." & Chr(13) & Chr(13) & "Remova-o e confira os dados antes de enviar.", vbExclamation
        Exit Sub
    End If
Next tbl


Iniciar_zeq_servidao:

Dim MsgLoad_zeq_servidao As VbMsgBoxResult

MsgLoad_zeq_servidao = MsgBox("Deseja enviar os dados para o banco de dados, nas APIs ""zeq_servidao""?" _
    , vbInformation + vbYesNo + vbDefaultButton2, "Enviar dados para ""zeq_servidao""?")

If MsgLoad_zeq_servidao = vbNo Then
    
    If LoadAll = "1" Then
        Dim MsgLoad_zeq_servidao_reask As VbMsgBoxResult

        MsgLoad_zeq_servidao_reask = MsgBox("Deseja mesmo cancelar o envio dos dados para o banco de dados, na API ""zeq_servidao""?" & vbCrLf & vbCrLf & _
            "ATENÇÃO: O envio de todas as ZLI's e ZEQ's pendentes serão canceladas, e as que já foram enviadas serão mantidas.", vbExclamation + vbYesNo + vbDefaultButton2, "Cancelar envio para ""zeq_servidao""?")
        
        If MsgLoad_zeq_servidao_reask = vbNo Then
            GoTo Iniciar_zeq_servidao
        ElseIf MsgLoad_zeq_servidao_reask = vbYes Then
            MsgLoadCancel = MsgBox("Envio de dados para o banco de dados cancelado!", vbCritical, "Cancelado!")
            Exit Sub
        End If
    End If
    
    Exit Sub

End If


Iniciar:

Application.ScreenUpdating = False

TempoInicio = Now()

QtdeVaos = WorksheetFunction.CountA(Range("Tab_zeq_servidao[VÃO]"))


    ActiveWorkbook.Connections("Consulta - Query_ID_zeq_servidao").Refresh
    Unload UserForm_EnviandoAPI

Dim get_largura_lado_esquerdo As Variant
Dim get_largura_lado_direito As Variant
Dim get_largura_limpeza As Variant
Dim get_d_caracteristica_regiao_id As String
Dim get_licenca_ambiental_operacao As String
Dim get_autorizacao_supressao_veget As String
Dim get_d_tipo_vegetacao_predominante_id As String
Dim get_d_zeq_sim_nao_restricao_p_supressao_veget_id As String
Dim get_d_zeq_sim_nao_regiao_alagamentos_id As String
Dim get_d_zeq_sim_nao_regiao_vandalismo_id As String
Dim get_d_zeq_sim_nao_regiao_queimadas_id As String
Dim get_d_zeq_sim_nao_regiao_poluicao_id As String
Dim get_d_natureza_travessia_critica_id As String
Dim get_dist_vertic_cabo_travessia As Variant
Dim get_dist_horiz_torre_travessia As Variant
Dim get_observacao_travessia As String

Dim Repete As Integer
Repete = 1

Do While Repete <= QtdeVaos


NumVaoID = Range("Tab_zeq_servidao[VÃO]").Rows(Repete).text

sequencia_vao = Application.Match(NumVaoID, Range("Query_ID_zeq_servidao[identificacao_vao]"), 0)

ID = Range("Query_ID_zeq_servidao[ID_zeq_servidao]").Rows(sequencia_vao)


get_largura_lado_esquerdo = Range("Tab_zeq_servidao[LARGURA LADO ESQUERDO (m)]").Rows(Repete).Value
get_largura_lado_direito = Range("Tab_zeq_servidao[LARGURA LADO DIREITO (m)]").Rows(Repete).Value
get_largura_limpeza = Range("Tab_zeq_servidao[LARGURA LIMPEZA (m)]").Rows(Repete).Value
get_d_caracteristica_regiao_id = Range("Tab_zeq_servidao[CARACTERÍSTICA REGIÃO]").Rows(Repete).text
get_licenca_ambiental_operacao = Range("Tab_zeq_servidao[LICENCA AMBIENTAL DE OPERAÇÃO]").Rows(Repete).text
get_autorizacao_supressao_veget = Range("Tab_zeq_servidao[AUTORIZAÇÃO SUPRESSÃO VEGET]").Rows(Repete).text
get_d_tipo_vegetacao_predominante_id = Range("Tab_zeq_servidao[TIPO VEGETAÇÃO PREDOMINANTE]").Rows(Repete).text
get_d_zeq_sim_nao_restricao_p_supressao_veget_id = Range("Tab_zeq_servidao[RESTRIÇÃO P/ SUPRESSÃO VEGET]").Rows(Repete).text
get_d_zeq_sim_nao_regiao_alagamentos_id = Range("Tab_zeq_servidao[REGIÃO DE ALAGAMENTOS]").Rows(Repete).text
get_d_zeq_sim_nao_regiao_vandalismo_id = Range("Tab_zeq_servidao[REGIÃO DE VANDALISMO]").Rows(Repete).text
get_d_zeq_sim_nao_regiao_queimadas_id = Range("Tab_zeq_servidao[REGIÃO DE QUEIMADAS]").Rows(Repete).text
get_d_zeq_sim_nao_regiao_poluicao_id = Range("Tab_zeq_servidao[REGIÃO DE POLUIÇÃO]").Rows(Repete).text
get_d_natureza_travessia_critica_id = Range("Tab_zeq_servidao[NATUREZA DA TRAVESSIA CRÍTICA]").Rows(Repete).text
get_dist_vertic_cabo_travessia = Range("Tab_zeq_servidao[DIST VERTIC CABO-TRAVESSIA (m)]").Rows(Repete).Value
get_dist_horiz_torre_travessia = Range("Tab_zeq_servidao[DIST HORIZ TORRE-TRAVESSIA (m)]").Rows(Repete).Value
get_observacao_travessia = Range("Tab_zeq_servidao[OBSERVAÇÃO]").Rows(Repete).text


Select Case get_largura_lado_esquerdo
    Case "", "-": largura_lado_esquerdo = Null
    Case Else: largura_lado_esquerdo = get_largura_lado_esquerdo
End Select

Select Case get_largura_lado_direito
    Case "", "-": largura_lado_direito = Null
    Case Else: largura_lado_direito = get_largura_lado_direito
End Select

Select Case get_largura_limpeza
    Case "", "-": largura_limpeza = Null
    Case Else: largura_limpeza = get_largura_limpeza
End Select

Select Case get_d_caracteristica_regiao_id
    Case "", "-": d_caracteristica_regiao_id = Null
    Case "Área de Preservação Permanente": d_caracteristica_regiao_id = "APP"
    Case "Urbana": d_caracteristica_regiao_id = "URB"
    Case "Rural": d_caracteristica_regiao_id = "RUR"
    Case "Reserva Indígena e afins": d_caracteristica_regiao_id = "RI"
    Case "Unidade de Conservação": d_caracteristica_regiao_id = "UC"
    Case Else: d_caracteristica_regiao_id = "Erro"
End Select

Select Case get_licenca_ambiental_operacao
    Case "", "-": licenca_ambiental_operacao = Null
    Case Else: licenca_ambiental_operacao = get_licenca_ambiental_operacao
End Select

Select Case get_autorizacao_supressao_veget
    Case "", "-": autorizacao_supressao_veget = Null
    Case Else: autorizacao_supressao_veget = get_autorizacao_supressao_veget
End Select

Select Case get_d_zeq_sim_nao_restricao_p_supressao_veget_id
    Case "", "-": d_zeq_sim_nao_restricao_p_supressao_veget_id = Null
    Case "Sim": d_zeq_sim_nao_restricao_p_supressao_veget_id = "S"
    Case "Não": d_zeq_sim_nao_restricao_p_supressao_veget_id = "N"
    Case Else: d_zeq_sim_nao_restricao_p_supressao_veget_id = "Erro"
End Select

Select Case get_d_tipo_vegetacao_predominante_id
    Case "", "-": d_tipo_vegetacao_predominante_id = Null
    Case "Reflorestamento": d_tipo_vegetacao_predominante_id = "REF"
    Case "Rasteira (Herbácea)": d_tipo_vegetacao_predominante_id = "RAS"
    Case "Capoeira (Arbustiva)": d_tipo_vegetacao_predominante_id = "CAP"
    Case "Densa (Árborea)": d_tipo_vegetacao_predominante_id = "DEN"
    Case "Crescimento Rápido": d_tipo_vegetacao_predominante_id = "CRA"
    Case "Cultura": d_tipo_vegetacao_predominante_id = "CUL"
    Case Else: d_tipo_vegetacao_predominante_id = "Erro"
End Select

Select Case get_d_zeq_sim_nao_regiao_alagamentos_id
    Case "", "-": d_zeq_sim_nao_regiao_alagamentos_id = Null
    Case "Sim": d_zeq_sim_nao_regiao_alagamentos_id = "S"
    Case "Não": d_zeq_sim_nao_regiao_alagamentos_id = "N"
    Case Else: d_zeq_sim_nao_regiao_alagamentos_id = "Erro"
End Select

Select Case get_d_zeq_sim_nao_regiao_vandalismo_id
    Case "", "-": d_zeq_sim_nao_regiao_vandalismo_id = Null
    Case "Sim": d_zeq_sim_nao_regiao_vandalismo_id = "S"
    Case "Não": d_zeq_sim_nao_regiao_vandalismo_id = "N"
    Case Else: d_zeq_sim_nao_regiao_vandalismo_id = "Erro"
End Select

Select Case get_d_zeq_sim_nao_regiao_queimadas_id
    Case "", "-": d_zeq_sim_nao_regiao_queimadas_id = Null
    Case "Sim": d_zeq_sim_nao_regiao_queimadas_id = "S"
    Case "Não": d_zeq_sim_nao_regiao_queimadas_id = "N"
    Case Else: d_zeq_sim_nao_regiao_queimadas_id = "Erro"
End Select

Select Case get_d_zeq_sim_nao_regiao_poluicao_id
    Case "", "-": d_zeq_sim_nao_regiao_poluicao_id = Null
    Case "Sim": d_zeq_sim_nao_regiao_poluicao_id = "S"
    Case "Não": d_zeq_sim_nao_regiao_poluicao_id = "N"
    Case Else: d_zeq_sim_nao_regiao_poluicao_id = "Erro"
End Select

Select Case get_d_natureza_travessia_critica_id
    Case "", "-": d_natureza_travessia_critica_id = Null
    Case "Hidrovia": d_natureza_travessia_critica_id = "HID"
    Case "Elétrica": d_natureza_travessia_critica_id = "ELE"
    Case "Ferrovia": d_natureza_travessia_critica_id = "FER"
    Case "Gasoduto": d_natureza_travessia_critica_id = "GAS"
    Case "Rodovia": d_natureza_travessia_critica_id = "ROD"
    Case Else: d_natureza_travessia_critica_id = "Erro"
End Select

Select Case get_dist_vertic_cabo_travessia
    Case "", "-": dist_vertic_cabo_travessia = Null
    Case Else: dist_vertic_cabo_travessia = get_dist_vertic_cabo_travessia
End Select

Select Case get_dist_horiz_torre_travessia
    Case "", "-": dist_horiz_torre_travessia = Null
    Case Else: dist_horiz_torre_travessia = get_dist_horiz_torre_travessia
End Select

Select Case get_observacao_travessia
    Case "", "-": observacao_travessia = Null
    Case Else: observacao_travessia = get_observacao_travessia
End Select


    Dim WinHttpReq As Object
    Set WinHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
    WinHttpReq.Open "PUT", "http://apilevantamento.h2m.eng.br:3000/api/zeq_servidao_lt/" & ID, False
    WinHttpReq.SetRequestHeader "Content-Type", "application/json"

    Dim json As Object
    Set json = CreateObject("Scripting.Dictionary")
    
        json("largura_lado_esquerdo") = largura_lado_esquerdo
        json("largura_lado_direito") = largura_lado_direito
        json("largura_limpeza") = largura_limpeza
        json("d_caracteristica_regiao_id") = d_caracteristica_regiao_id
        json("licenca_ambiental_operacao") = licenca_ambiental_operacao
        json("autorizacao_supressao_veget") = autorizacao_supressao_veget
        json("d_zeq_sim_nao_restricao_p_supressao_veget_id") = d_zeq_sim_nao_restricao_p_supressao_veget_id
        json("d_tipo_vegetacao_predominante_id") = d_tipo_vegetacao_predominante_id
        json("d_zeq_sim_nao_regiao_alagamentos_id") = d_zeq_sim_nao_regiao_alagamentos_id
        json("d_zeq_sim_nao_regiao_vandalismo_id") = d_zeq_sim_nao_regiao_vandalismo_id
        json("d_zeq_sim_nao_regiao_queimadas_id") = d_zeq_sim_nao_regiao_queimadas_id
        json("d_zeq_sim_nao_regiao_poluicao_id") = d_zeq_sim_nao_regiao_poluicao_id
        json("d_natureza_travessia_critica_id") = d_natureza_travessia_critica_id
        json("dist_vertic_cabo_travessia") = dist_vertic_cabo_travessia
        json("dist_horiz_torre_travessia") = dist_horiz_torre_travessia
        json("observacao_travessia") = observacao_travessia
    
    
    Dim jsonData As String
    jsonData = JsonConverter.ConvertToJson(json)
    
    WinHttpReq.Send jsonData
    
    If InStr(1, WinHttpReq.ResponseText, " COM SUCESSO") = 0 Then
        
        If Qtde_Erros = "" Then
            Qtde_Erros = 1
        Else: Qtde_Erros = Qtde_Erros + 1
        End If
        
        If Vaos_Erros = "" Then
            Vaos_Erros = NumVaoID
        Else: Vaos_Erros = Vaos_Erros & ", " & NumVaoID
        End If
        
    End If

    Repete = Repete + 1

Set largura_lado_esquerdo = Nothing
Set largura_lado_direito = Nothing
Set largura_limpeza = Nothing
Set d_caracteristica_regiao_id = Nothing
Set licenca_ambiental_operacao = Nothing
Set autorizacao_supressao_veget = Nothing
Set d_tipo_vegetacao_predominante_id = Nothing
Set d_zeq_sim_nao_restricao_p_supressao_veget_id = Nothing
Set d_zeq_sim_nao_regiao_alagamentos_id = Nothing
Set d_zeq_sim_nao_regiao_vandalismo_id = Nothing
Set d_zeq_sim_nao_regiao_queimadas_id = Nothing
Set d_zeq_sim_nao_regiao_poluicao_id = Nothing
Set d_natureza_travessia_critica_id = Nothing
Set dist_vertic_cabo_travessia = Nothing
Set dist_horiz_torre_travessia = Nothing
Set observacao_travessia = Nothing

Set WinHttpReq = Nothing
Set json = Nothing

Loop

Application.ScreenUpdating = True

TempoFim = Now()

    If Qtde_Erros = "" Then
        MsgLog = "Dados de " & APIAtual & " da LT " & CodLT & " enviados para o banco de dados com sucesso!" & Chr(13) & Chr(13) & _
        "Total de vãos: " & Chr(13) & _
        Chr(13) & _
        "Tempo de execução: " & TimeValue(Format(TempoFim - TempoInicio, "dd/mm/yyyy hh:mm:ss"))
        
        If LoadAll = "" Then
            ShowMsgLog = MsgBox(MsgLog, vbInformation, "Sucesso!")
        End If
            IdErro = 0
            
    Else:
        MsgLog = Qtde_Erros & " vão(s) de servidão não foi(oram) enviado(s): " & Chr(13) & Chr(13) & _
        Vaos_Erros & Chr(13) & Chr(13) & _
        "Verifique os dados e tente novamente." & _
        Chr(13) & Chr(13) & _
        "Tempo de execução: " & TimeValue(Format(TempoFim - TempoInicio, "dd/mm/yyyy hh:mm:ss"))
        
        If LoadAll = "" Then
            ShowMsgLog = MsgBox(MsgLog, vbCritical, "Erro(s)!")
        End If
            IdErro = 1
            
    End If


'Gerando registro de log:

        On Error GoTo AddNewLog
            RowLog = Application.WorksheetFunction.Match(APIAtual, Range("TabLogErros[API]"), 0)
                Range("TabLogErros[API]").Rows(RowLog).Value = APIAtual
                Range("TabLogErros[Erro]").Rows(RowLog).Value = IdErro
                Range("TabLogErros[Início]").Rows(RowLog).Value = TempoInicio
                Range("TabLogErros[Fim]").Rows(RowLog).Value = TempoFim
                Range("TabLogErros[Msg]").Rows(RowLog).Value = MsgLog
            GoTo Finalizar
AddNewLog:
        On Error GoTo -1
        On Error GoTo 0
                Range("TabLogErros[API]").SpecialCells(xlCellTypeBlanks).Rows(1).Value = APIAtual
                Range("TabLogErros[Erro]").SpecialCells(xlCellTypeBlanks).Rows(1).Value = IdErro
                Range("TabLogErros[Início]").SpecialCells(xlCellTypeBlanks).Rows(1).Value = TempoInicio
                Range("TabLogErros[Fim]").SpecialCells(xlCellTypeBlanks).Rows(1).Value = TempoFim
                Range("TabLogErros[Msg]").SpecialCells(xlCellTypeBlanks).Rows(1).Value = MsgLog

        
Finalizar:
ActiveWorkbook.Save

    On Error GoTo -1
    On Error GoTo 0

If LoadAll = "1" Then
    'Call
End If


If LoadAll = "1" Then
    Unload UserForm_EnviandoAPI
    TempoFimAll = Now()
    MsgFimLoadAll = MsgBox("O envio das informações todas as ZLI's e ZEQ's para o banco de dados foi concluído!" & _
        Chr(13) & Chr(13) & _
        "Tempo total de execução: " & TimeValue(Format(TempoFimAll - TempoInicioAll, "dd/mm/yyyy hh:mm:ss")) _
        , vbInformation, "Envio concluído!")
    LoadAll = ""

    On Error GoTo FimGeral
    If IsNumeric(Application.WorksheetFunction.Match(1, Range("TabLogErros[Erro]"), 0)) Then
        MsgFimLoadAll_Alert = MsgBox("Foram listados erros durante os envios das informações. " & _
        "Consulte os registros para ver as mensagens novamente. Identifique os erros, corrija-os e realize um novo envio." _
        , vbExclamation, "Erros registrados!")
    End If
End If

FimGeral:

End Sub
