Sub Atualizar_SAP() 'Versão 2.4

    Msgbox "Tu continua berola, só que agora na versão 2.4 uahuashuhusahusha"
    Exit Sub

If Range("Label_NomeLT").Locked = True Then
    MsgBox "Os dados já foram preenchidos!"
    Exit Sub
End If

MsgBox "Os seguintes dados serão solicitados:"

'Application.ScreenUpdating = False

BaseVBA_SAP = Application.ActiveWorkbook.Name

LT_CodLT = Range("Label_CodLT").Value
LT_NomeLT = Range("Label_NomeLT").Value
LT_TensaoLT = Application.WorksheetFunction.VLookup(Range("Label_NomeLT"), Range("BASE_LTs"), 14, False)

    
'**********ATUALIZANDO BASES**********

Atualizar_Bases:

On Error GoTo -1
On Error GoTo 0
On Error GoTo ErrorDownload

    ActiveWorkbook.Queries.Item("Param_CodLT").Formula = _
        """" & LT_CodLT & """ meta [IsParameterQuery=true, Type=""Any"", IsParameterQueryRequired=True]"

    'ActiveWorkbook.Connections("Consulta - BASE_Cabos").Refresh
    'ActiveWorkbook.Connections("Consulta - BASE_BD_SerieEstrutura").Refresh
        'ActiveWorkbook.Connections("Consulta - BASE_BD_SerieEstrutura_Aux").Refresh
    'ActiveWorkbook.Connections("Consulta - BASE_BD_Aterramento").Refresh
        'ActiveWorkbook.Connections("Consulta - BASE_BD_Aterramento_Aux").Refresh
    'ActiveWorkbook.Connections("Consulta - BASE_BD_TorresLTGeral").Refresh
    
    'ActiveWorkbook.Connections("Consulta - BASE_BD_ProjetosLT").Refresh


On Error GoTo -1
On Error GoTo 0

If Range("BASE_BD_ProjetosLT[qtde_total_estruturas]").Rows(1).Value <> Application.WorksheetFunction.CountIfs(Range("BASE_BD_TorresLTGeral[d_portico_id]"), "<>1", Range("BASE_BD_TorresLTGeral[d_portico_id]"), "<>3") Then

    Dim QtdeTorresBDIT_Erradas As VbMsgBoxResult
        QtdeTorresBDIT_Erradas = MsgBox("A contagem de torres cadastradas no banco de dados não coincide com o valor inserido no cadastro da ""LT Geral"", também do banco de dados. " _
        & Chr(13) & Chr(13) & _
        "Revise os dados, prossiga com as devidas correções e tente novamente.", vbCritical, "Quantidade de torres incorreta!")
        Exit Sub

End If

If Range("BASE_BD_ProjetosLT[qtde_total_vaos]").Rows(1).Value <> Application.WorksheetFunction.CountA(Range("BASE_BD_VaosLT[identificacao_vao]")) Then

    Dim QtdeVaosBDIT_Erradas As VbMsgBoxResult
        QtdeVaosBDIT_Erradas = MsgBox("A contagem de vãos cadastrados no banco de dados não coincide com o valor inserido no cadastro da ""LT Geral"", também do banco de dados. " _
        & Chr(13) & Chr(13) & _
        "Revise os dados, prossiga com as devidas correções e tente novamente.", vbCritical, "Quantidade de torres incorreta!")
        Exit Sub

End If


GoTo Inicio_Selecionar_LC

ErrorDownload:

Dim MsgErrorDownload As VbMsgBoxResult
MsgErrorDownload = MsgBox("Ocorreu um erro ao fazer o download dos dados. Verifique sua conexão com a internet e tente novamente.", vbCritical + vbRetryCancel + vbDefaultButton2, "Erro: download dados")

If MsgErrorDownload = vbCancel Then
    CancelaPreenchimento = MsgBox("Processo de preenchimento automático cancelado!", vbCritical + vbOKOnly, "Preenchimento cancelado!")
    Exit Sub
ElseIf MsgErrorDownload = vbRetry Then
    GoTo Atualizar_Bases
End If



'**********SELECIONANDO LISTA DE CONSTRUÇÃO**********

Inicio_Selecionar_LC:

Dim Msg_SelectLC As VbMsgBoxResult
Msg_SelectLC = MsgBox("Selecione a Lista de Construção da LT " & LT_CodLT & ":", vbOKCancel + vbExclamation, "Selecionar Lista de Construção")

If Msg_SelectLC = vbCancel Then
CancelaPreenchimento = MsgBox("Processo de preenchimento automático cancelado!", vbCritical + vbOKOnly, "Preenchimento cancelado!")
Exit Sub
End If

Selecionar_LC:

Application.DisplayAlerts = False

Set Select_LC = Application.FileDialog(msoFileDialogOpen)

With Select_LC

.AllowMultiSelect = False
.Show
.Execute

End With

Application.DisplayAlerts = True

    If Application.ActiveWorkbook.Name = BaseVBA_SAP Then
        CancelaPreenchimento = MsgBox("Processo de preenchimento automático cancelado!", vbCritical + vbOKOnly, "Preenchimento cancelado!")
        Exit Sub
    End If

On Error GoTo CheckLC

    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select

CheckLC:

On Error GoTo -1
On Error GoTo 0
    
If Range("A1") = "Dados Gerais" _
And Range("A2") = "Nome do Circuito" _
And Range("A5") = "NomeCircuito" _
And Range("A6") <> "" _
Then
LC_Selected = "OK"
Else: LC_Selected = "NOK"
End If

    LC_CodLT = Range("B6").Value
    LC_NomeLT = Range("A6").Value
    LC_CaminhoLC = Application.ActiveWorkbook.FullName
    LC_NomeLC = Application.ActiveWorkbook.Name
    LC_Extensao = Round(Application.WorksheetFunction.Sum(Range("ListadeConstrucao[Vao]")) / 1000, 2)
    LC_QtdeEstruturas = Application.WorksheetFunction.CountA(Range("ListadeConstrucao[NumOper]"))
    LC_SomaFaixaServidao = Application.WorksheetFunction.Sum(Range("ListadeConstrucao[[FaixaServ1]:[FaixaServ2]]"))

Application.ActiveWorkbook.Close (savechanges = True)

Dim LCnoOK As VbMsgBoxResult

    If LC_Selected <> "OK" Then
        LCnoOK = MsgBox("O arquivo selecionado não corresponde à um modelo de Lista de Construção. Selecione a Lista de Construção corretamente.", vbCritical + vbRetryCancel, "LC não encontrada!")

        If LCnoOK = vbCancel Then
                CancelaPreenchimento = MsgBox("Processo de preenchimento automático cancelado!", vbCritical + vbOKOnly, "Preenchimento cancelado!")
            Exit Sub
        ElseIf LCnoOK = vbRetry Then
            GoTo Selecionar_LC
        End If

    End If


If LC_CodLT <> LT_CodLT Or LC_NomeLT <> LT_NomeLT Then
Dim NoLCDiferente As VbMsgBoxResult
NoLCDiferente = MsgBox("A Lista de Construção selecionada não coincide com os dados da LT. " & _
"Revise o preenchimento ou selecione o arquivo correto." & Chr(13) & Chr(13) & _
"Arquivo esperado:      " & LT_CodLT & " (" & LT_NomeLT & ")" & Chr(13) & _
"Arquivo selecionado:  " & LC_CodLT & " (" & LC_NomeLT & ")", _
vbRetryCancel + vbCritical)

If NoLCDiferente = vbCancel Then
CancelaPreenchimento = MsgBox("Processo de preenchimento automático cancelado!", vbCritical + vbOKOnly, "Preenchimento cancelado!")
Exit Sub
ElseIf NoLCDiferente = vbRetry Then
GoTo Selecionar_LC

End If
End If

If LC_Extensao <> Range("BASE_BD_ProjetosLT[extensao_total_linha]").Rows(1).Value Then
    
    Dim Extensao_Errada As VbMsgBoxResult
        Extensao_Errada = MsgBox("A somatória de vãos da lista de construção não coincide com a extensão total da linha registrada no banco de dados. " _
        & Chr(13) & Chr(13) & _
        "Revise os dados, prossiga com as devidas correções e tente novamente.", vbCritical, "Extensão incorreta!")
        Exit Sub

End If


If LC_QtdeEstruturas <> Application.WorksheetFunction.CountA(Range("BASE_BD_TorresLTGeral[numero_torre]")) Then
    
    Dim QtdeEstruturasLC_Errada As VbMsgBoxResult
        QtdeEstruturasLC_Errada = MsgBox("A quantidade de estruturas da lista de construção não coincide com a contagem de estruturas da linha registrada no banco de dados." _
        & Chr(13) & Chr(13) & _
        "Revise os dados, prossiga com as devidas correções e tente novamente.", vbCritical, "Quantidade de estruturas incorreta!")
        Exit Sub

End If


YesLC = MsgBox("Lista de Construção da LT " & LT_TensaoLT & " kV - " & LC_CodLT & " (" & LC_NomeLT & ") identificada!" _
& Chr(13) & Chr(13) & "Os dados da LC da LT " & LC_CodLT & " foram armazenados!", vbInformation, "LC " & LC_CodLT & " Identificada!")




'**********SELECIONANDO A PLANILHA COM OS DADOS DO BDIT**********

Dim Msg_SelectBDIT As VbMsgBoxResult
Msg_SelectBDIT = MsgBox("Agora, selecione a planilha preenchida do BDIT correspondente à LT " & LT_TensaoLT & " kV - " & LT_CodLT & " (" & LT_NomeLT & ")", vbOKCancel + vbExclamation, "Selecionar dados do BDIT")

If Msg_SelectBDIT = vbCancel Then
CancelaPreenchimento = MsgBox("Processo de preenchimento automático cancelado!", vbCritical + vbOKOnly, "Preenchimento cancelado!")
Exit Sub
End If

Selecionar_BDIT:

Application.DisplayAlerts = False

Set Select_BDIT = Application.FileDialog(msoFileDialogOpen)

With Select_BDIT
.AllowMultiSelect = False
.Show
.Execute
End With

Application.DisplayAlerts = True

    If Application.ActiveWorkbook.Name = BaseVBA_SAP Then
        CancelaPreenchimento = MsgBox("Processo de preenchimento automático cancelado!", vbCritical + vbOKOnly, "Preenchimento cancelado!")
        Exit Sub
    End If

On Error GoTo -1
On Error GoTo 0
On Error GoTo ErroBDIT



If Range("I3") = "DADOS BDIT" _
And Range("I9") = "DADOS GERAIS DA LT" _
And Range("I21") = "Cod Ident Ativo na Concessionária" _
Then
BDIT_Selected = "OK"
BDIT_Version = 4

ElseIf Range("I3") = "DADOS BDIT" _
And Range("I9") = "DADOS GERAIS DA LT" _
And Range("I22") = "Cod Ident Ativo na Concessionária" _
Then
BDIT_Selected = "OK"
BDIT_Version = 5

Else: BDIT_Selected = "NOK"
End If

If BDIT_Selected = "NOK" Then
GoTo ErroBDIT
End If


'Versão 4:

    If BDIT_Version = 4 Then
    
        Sheets("Serie_Estruturas").Select
            ActiveSheet.Unprotect
        Range("F6").Select
        Range(Selection, Selection.End(xlDown)).Select
            Qtde_SerieEst = Application.WorksheetFunction.CountA(Selection) - 1
        Range("G6").Select
        Range(Selection, Selection.End(xlDown)).Select
            Qtde_Torres = Application.WorksheetFunction.Sum(Selection)
    
    
        Sheets("Serie_Estruturas_C1").Select
            ActiveSheet.Unprotect
        Range("F7").Select
        Range(Selection, Selection.End(xlDown)).Select
            Qtde_SerieEstCircuito = Application.WorksheetFunction.CountA(Selection) - 1
    
        If Qtde_SerieEst <> Qtde_SerieEstCircuito Then
            GoTo ErroBDIT
        End If
        
        
        Sheets("Projetos").Select
            Projeto = Range("G6").Value

        Sheets("Torre").Select
            ActiveSheet.Unprotect
        Range("D1").Select
        Range(Selection, Selection.End(xlDown)).Select
            Qtde_EstruturasTabTorres = Application.WorksheetFunction.CountA(Selection) - 1

        Sheets("LT_Geral").Select
            BDIT_CodLT = Range("J21")


    End If


'Versão 5:

    If BDIT_Version = 5 Then
        Qtde_SerieEst = Application.WorksheetFunction.CountA(Range("Tab_SerieEstrut[Nome Estrutura]"))
        Qtde_Torres = Application.WorksheetFunction.Sum(Range("Tab_SerieEstrut[Qtde. Total na LT]"))
            If Qtde_SerieEst <> Qtde_SerieEstCircuito Then
                GoTo ErroBDIT
            End If
        Qtde_SerieEstCircuito = Application.WorksheetFunction.CountA(Range("Tab_SerieEstrutC1[Nome Estrutura]"))
        Qtde_EstruturasTabTorres = Application.WorksheetFunction.CountA(Range("Tab_Torres[Nome LT]"))
        Projeto = Range("Label_IDProjeto").Value
        BDIT_CodLT = Range("Label_CodAtivoConcessionaria")

    End If



BDIT_CaminhoBDIT = Application.ActiveWorkbook.FullName
BDIT_NomeBDIT = Application.ActiveWorkbook.Name




'Verificando se as torres iniciais e finais são pórticos:
    
    Worksheets("Torre").Activate
        Dim LastRowTorres As Long


    'Versão 4:
    
        If BDIT_Version = 4 Then
        
        LastRowTorres = Qtde_EstruturasTabTorres + 1
        
            If Range("K2") = "1 - Portico Inicial" Then
                PorticoInicialExiste = 1
            Else: PorticoInicialExiste = 0
            End If
        
            If Range("K" & LastRowTorres) = "3 - Portico Final" Then
                PorticoFinalExiste = 1
            Else: PorticoFinalExiste = 0
            End If
    
        End If
    
    
    'Versão 5:
    
        If BDIT_Version = 5 Then
        
        LastRowTorres = Range("Tab_Torres[Nome LT]").Rows.Count + 1
        
            If Range("M2") = "1 - Portico Inicial" Then
                PorticoInicialExiste = 1
            Else: PorticoInicialExiste = 0
            End If
        
            If Range("M" & LastRowTorres) = "3 - Portico Final" Then
                PorticoFinalExiste = 1
            Else: PorticoFinalExiste = 0
            End If
    
        End If


Application.ActiveWorkbook.Close (savechanges = True)

GoTo OKBDIT


ErroBDIT:

Dim NoBDIT As VbMsgBoxResult

Application.ActiveWorkbook.Close (savechanges = True)

NoBDIT = MsgBox("O arquivo secionado não corresponde à um modelo padrão do BDIT. Selecione um arquivo válido ou revise o padrão do arquivo.", _
vbCritical + vbRetryCancel, "BDIT não identificado!")

If NoBDIT = vbRetry Then
On Error GoTo 0
GoTo Selecionar_BDIT

ElseIf NoBDIT = vbCancel Then
CancelaPreenchimento = MsgBox("Processo de preenchimento automático cancelado!", vbCritical + vbOKOnly, "Preenchimento cancelado!")
Exit Sub

End If


OKBDIT:

On Error GoTo -1
On Error GoTo 0

If BDIT_CodLT <> LT_CodLT Then

Dim Msg_BDITIncorreto As VbMsgBoxResult

Msg_BDITIncorreto = MsgBox("O arquivo selecionado se refere aos dados do BDIT da LT " & BDIT_CodLT & "." & _
Chr(13) & Chr(13) & _
"Selecione o arquivo correspondente à LT " & LT_CodLT, vbOKCancel + vbCritical, "Arquivo incorreto")

If Msg_BDITIncorreto = vbOK Then
On Error GoTo -1
On Error GoTo 0
GoTo Selecionar_BDIT

ElseIf Msg_BDITIncorreto = vbCancel Then
CancelaPreenchimento = MsgBox("Processo de preenchimento automático cancelado!", vbCritical + vbOKOnly, "Preenchimento cancelado!")
Exit Sub

End If
End If


'Validando se foi selecionado projeto único ou projeto geral:

    If (Projeto <> 0 And Projeto <> 1) Or Qtde_Torres <> Qtde_EstruturasTabTorres - PorticoInicialExiste - PorticoFinalExiste Then
        ErroProjetoIncompleto = MsgBox("Foi identificado que, apesar do arquivo corresponder à LT " & BDIT_CodLT & ", a tabela ""Série Estrutura"" está incompleta. Caso a LT possua mais de um projeto, certifique-se de selecionar o arquivo ""GERAL"", ou cancele o processo, revise os arquivos e tente novamente.", vbCritical + vbRetryCancel, "Dados incompletos")
        If ErroProjetoIncompleto = vbCancel Then
                CancelaPreenchimento = MsgBox("Processo de preenchimento automático cancelado!", vbCritical + vbOKOnly, "Preenchimento cancelado!")
            Exit Sub
        ElseIf ErroProjetoIncompleto = vbRetry Then
                GoTo Selecionar_BDIT
        End If
    End If


'Coinferindo se quantidade de torres da tab "Serie Estrutura" e o carregado no banco de dados são iguais:

If Qtde_Torres <> Range("BASE_BD_ProjetosLT[qtde_total_estruturas]").Rows(1).Value Then

    Dim QtdeTorresSerieEst_Erradas As VbMsgBoxResult
        QtdeTorresSerieEst_Erradas = MsgBox("A somatória de torres cadastradas no arquivo selecionado não coincide com a contagem de torres cadastradas no banco de dados." _
        & Chr(13) & Chr(13) & _
        "Revise os dados, prossiga com as devidas correções e tente novamente.", vbCritical, "Quantidade de torres incorreta!")
        Exit Sub

End If


'Tudo OK:

YesBDIT = MsgBox("Planilha com os dados do BDIT da LT " & LT_TensaoLT & " kV - " & LT_CodLT & " (" & LT_NomeLT & ") identificado!" _
& Chr(13) & Chr(13) & "Os dados do BDIT foram armazenados!", vbInformation, "BDIT " & LT_CodLT & " Identificado!")




'**********CONFIRMAÇÃO FINAL**********

Proced_ConfirmacaoFinal:

Dim ConfirmacaoFinal As VbMsgBoxResult
ConfirmacaoFinal = MsgBox("Ao prosseguir, os dados disponíveis serão preenchidos automaticamente, e o processo não poderá ser desfeito." & _
Chr(13) & Chr(13) & "Deseja prosseguir com o preenchimento dos dados?", vbInformation + vbYesNo + vbDefaultButton2, "Iniciar preenchimento?")

If ConfirmacaoFinal = vbNo Then

Dim CertezaConfirmacaoFinal As VbMsgBoxResult
CertezaConfirmacaoFinal = MsgBox("Tem certeza que deseja cancelar o processo de preenchimento de dados do SAP da LT " & LT_CodLT & "?", vbExclamation + vbYesNo + vbDefaultButton2, "Cancelar: tem certeza?")

If CertezaConfirmacaoFinal = vbNo Then
GoTo Proced_ConfirmacaoFinal
ElseIf CertezaConfirmacaoFinal = vbYes Then
        CancelaPreenchimento = MsgBox("Processo de preenchimento automático cancelado!", vbCritical + vbOKOnly, "Preenchimento cancelado!")
Exit Sub
End If

End If


On Error GoTo -1
On Error GoTo 0



'**********ZLI TRANSMISSÃO**********

Sheets("zli_transmissao").Unprotect (StrConv(Base64Decode("UGFkaWxoYUgyTSo="), vbUnicode))
Range("Tab_zli_li_transmissao[[Coluna2]]").ClearContents
Range("Tab_zli_li_transmissao[[Coluna2]]").Interior.Pattern = xlNone
Range("Tab_zli_li_transmissao[[Coluna2]]").Locked = False


Range("Label_ZLI_Transmissao_Classificacao") = "Rede Básica (RB)"

If LT_CodLT = "BTOITA2" Then
    Range("Label_ZLI_Transmissao_Manutencao") = "Terceiros"
    Range("Label_ZLI_Transmissao_Operacao") = "Terceiros"
Else:
    Range("Label_ZLI_Transmissao_Manutencao") = "Própria"
    Range("Label_ZLI_Transmissao_Operacao") = "Própria"
End If

Range("Label_ZLI_Transmissao_Classificacao,Label_ZLI_Transmissao_Manutencao,Label_ZLI_Transmissao_Operacao").Locked = True



If Application.WorksheetFunction.Average(Range("BASE_BD_ProjetosLT[temperatura_maxima_condutor_longa_duracao]")) = Range("BASE_BD_ProjetosLT[temperatura_maxima_condutor_longa_duracao]").Cells(1, 1).Value _
    And Application.WorksheetFunction.Average(Range("BASE_BD_ProjetosLT[quantidade_capacidade_operativas_curta_duracao_temperatura_maxima_condutor]")) = Range("BASE_BD_ProjetosLT[quantidade_capacidade_operativas_curta_duracao_temperatura_maxima_condutor]").Cells(1, 1).Value Then
    
    Range("Label_ZLI_Transmissao_TemperLongaDur") = Application.WorksheetFunction.Average(Range("BASE_BD_ProjetosLT[temperatura_maxima_condutor_longa_duracao]"))
    Range("Label_ZLI_Transmissao_TemperCurtaDur") = Application.WorksheetFunction.Average(Range("BASE_BD_ProjetosLT[quantidade_capacidade_operativas_curta_duracao_temperatura_maxima_condutor]"))
    
    Range("Label_ZLI_Transmissao_TemperLongaDur,Label_ZLI_Transmissao_TemperCurtaDur").Locked = True

Else: Aviso1 = True
    Range("Label_ZLI_Transmissao_TemperLongaDur,Label_ZLI_Transmissao_TemperCurtaDur").Interior.Color = 65535

End If



Range("Label_ZLI_Transmissao_ExtensPropria") = Range("BASE_BD_ProjetosLT[extensao_total_linha]").Value
Range("Label_ZLI_Transmissao_QtdeEstruturas") = Range("BASE_BD_ProjetosLT[qtde_total_estruturas]").Value
Range("Label_ZLI_Transmissao_ModeloTorreTipica") = Range("BASE_BD_ProjetosLT[nome_estrutura_tipica]").Rows(1)

Range("Label_ZLI_Transmissao_ExtensPropria, Label_ZLI_Transmissao_QtdeEstruturas, Label_ZLI_Transmissao_ModeloTorreTipica").Locked = True

Sheets("zli_transmissao").Activate
Range("A1").Activate
Sheets("Menu").Activate

Sheets("zli_transmissao").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=False, Password:=StrConv(Base64Decode("UGFkaWxoYUgyTSo="), vbUnicode)




'**********ZLI PARÂMETROS OP.**********

Sheets("zli_parametros_OP").Unprotect (StrConv(Base64Decode("UGFkaWxoYUgyTSo="), vbUnicode))
Range("Tab_zli_parametros_op[[Coluna2]]").ClearContents
Range("Tab_zli_parametros_op[[Coluna2]]").Interior.Pattern = xlNone
Range("Tab_zli_parametros_op[[Coluna2]]").Locked = False

If Application.WorksheetFunction.CountA(Range("BASE_BD_ProjetosLT[id]")) = 1 Then
    ProjetoUnico = True
Else: Aviso2 = True
End If


If ProjetoUnico Then
    
    Range("Label_ZLI_ParametrosOp_FlechaMaxCondut") = Range("BASE_BD_ProjetosLT[flecha_cabo_condutor]").Value
    Range("Label_ZLI_ParametrosOp_FlechaMaxPR") = Range("BASE_BD_ProjetosLT[flecha_cabo_para_raios]").Value
    
    Range("Label_ZLI_ParametrosOp_ResistSeqPosit") = Range("BASE_BD_ProjetosLT[resistencia_longitudinal_sequencia_positiva]").Value
    Range("Label_ZLI_ParametrosOp_ReatSeqPosit") = Range("BASE_BD_ProjetosLT[reatancia_indutiva_longitudinal_sequencia_positiva]").Value
    Range("Label_ZLI_ParametrosOp_SuscepSeqPosit") = Range("BASE_BD_ProjetosLT[susceptancia_capacitiva_transversal_sequencia_positiva]").Value
    Range("Label_ZLI_ParametrosOp_ResistSeqZero") = Range("BASE_BD_ProjetosLT[resistencia_longitudinal_sequencia_zero]").Value
    Range("Label_ZLI_ParametrosOp_ReatSeqZero") = Range("BASE_BD_ProjetosLT[reatancia_indutiva_longitudinal_sequencia_zero]").Value
    Range("Label_ZLI_ParametrosOp_SuscepSeqZero") = Range("BASE_BD_ProjetosLT[susceptancia_capacitiva_transversal_sequencia_zero]").Value

    Range("Label_ZLI_ParametrosOp_FlechaMaxCondut, Label_ZLI_ParametrosOp_FlechaMaxPR").Locked = True
    Range("Label_ZLI_ParametrosOp_ResistSeqPosit, Label_ZLI_ParametrosOp_ReatSeqPosit, Label_ZLI_ParametrosOp_SuscepSeqPosit").Locked = True
    Range("Label_ZLI_ParametrosOp_ResistSeqZero, Label_ZLI_ParametrosOp_ReatSeqZero, Label_ZLI_ParametrosOp_SuscepSeqZero").Locked = True

End If



Sheets("zli_parametros_OP").Activate
Range("A1").Activate
Sheets("Menu").Activate

Sheets("zli_parametros_OP").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=False, Password:=StrConv(Base64Decode("UGFkaWxoYUgyTSo="), vbUnicode)




'**********ZEQ ESTRUTURA GERAL**********

Sheets("zeq_estru_geral").Unprotect (StrConv(Base64Decode("UGFkaWxoYUgyTSo="), vbUnicode))
Range("Tab_zeq_estru_geral").ClearContents
Range("Tab_zeq_estru_geral").Interior.Pattern = xlNone
Range("Tab_zeq_estru_geral").Locked = False

Sheets("zeq_estru_geral").Activate
PrimeiraLinha = Range("Tab_zeq_estru_geral").Row
Rows(PrimeiraLinha + 1 & ":" & PrimeiraLinha + 1).Select
Range(Selection, Selection.End(xlDown)).Delete Shift:=xlUp

QtdeEstruturasTabTorresBD = Application.WorksheetFunction.CountA(Range("BASE_BD_TorresLTGeral[numero_torre]"))

ActiveSheet.ListObjects("Tab_zeq_estru_geral").Resize Range("$B$" & PrimeiraLinha - 1 & ":$V$" & 1 * QtdeEstruturasTabTorresBD + (PrimeiraLinha - 1))


Range("Tab_zeq_estru_geral[NÚMERO DE OPERAÇÃO]").Rows(1).Activate
    ActiveCell.Range("A1:A" & QtdeEstruturasTabTorresBD).Value = Range("BASE_BD_TorresLTGeral[numero_torre]").Value

Range("Tab_zeq_estru_geral[SILHUETA]").Rows(1).Activate
    ActiveCell.Range("A1:A" & QtdeEstruturasTabTorresBD).Value = Range("BASE_BD_TorresLTGeral[nome_estrutura]").Value

Range("Tab_zeq_estru_geral[TIPO DE ESTRUTURA DE LINHA]").Rows(1).Activate
    ActiveCell.Range("A1:A" & QtdeEstruturasTabTorresBD).Value = Range("BASE_BD_TorresLTGeral[Caracteristica2]").Value

Range("Tab_zeq_estru_geral[MATERIAL CONSTRUTIVO]").Rows(1).Activate
    ActiveCell.Range("A1:A" & QtdeEstruturasTabTorresBD).Value = Range("BASE_BD_TorresLTGeral[material_predominante]").Value

Range("Tab_zeq_estru_geral[TIPO DE CIRCUITO]").Rows(1).Activate
    ActiveCell.Range("A1:A" & QtdeEstruturasTabTorresBD).Value = Range("BASE_BD_TorresLTGeral[tipo_circuito]").Value

Range("Tab_zeq_estru_geral[ALTURA MISULA (m)]").Rows(1).Activate
    ActiveCell.Range("A1:A" & QtdeEstruturasTabTorresBD).Value = Range("BASE_BD_TorresLTGeral[altura_torre]").Value

Range("Tab_zeq_estru_geral[ALTITUDE]").Rows(1).Activate
    ActiveCell.Range("A1:A" & QtdeEstruturasTabTorresBD).Value = Range("BASE_BD_TorresLTGeral[altitude]").Value

Range("Tab_zeq_estru_geral[LATITUDE]").Rows(1).Activate
    ActiveCell.Range("A1:A" & QtdeEstruturasTabTorresBD).Value = Range("BASE_BD_TorresLTGeral[latitude]").Value

Range("Tab_zeq_estru_geral[LONGITUDE]").Rows(1).Activate
    ActiveCell.Range("A1:A" & QtdeEstruturasTabTorresBD).Value = Range("BASE_BD_TorresLTGeral[longitude]").Value

Range("Tab_zeq_estru_geral[DATUM]") = "SIRGAS 2000"


    Range("Tab_zeq_estru_geral[NÚMERO DE OPERAÇÃO], Tab_zeq_estru_geral[SILHUETA], Tab_zeq_estru_geral[TIPO DE ESTRUTURA DE LINHA]").Locked = True
    Range("Tab_zeq_estru_geral[MATERIAL CONSTRUTIVO], Tab_zeq_estru_geral[TIPO DE CIRCUITO], Tab_zeq_estru_geral[ALTURA MISULA (m)]").Locked = True
    Range("Tab_zeq_estru_geral[ALTITUDE], Tab_zeq_estru_geral[LATITUDE], Tab_zeq_estru_geral[LONGITUDE]").Locked = True

Sheets("zeq_estru_geral").Activate
Range("A1").Activate
Sheets("Menu").Activate

Sheets("zeq_estru_geral").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=False, Password:=StrConv(Base64Decode("UGFkaWxoYUgyTSo="), vbUnicode)



'**********ZEQ ESTRUTURA AUTOPORTANTE/ESTAI**********

Sheets("zeq_estru_autop&estai").Unprotect (StrConv(Base64Decode("UGFkaWxoYUgyTSo="), vbUnicode))
Range("Tab_zeq_estru_autop_estai").ClearContents
Range("Tab_zeq_estru_autop_estai").Interior.Pattern = xlNone
Range("Tab_zeq_estru_autop_estai").Locked = False

Sheets("zeq_estru_autop&estai").Activate
PrimeiraLinha = Range("Tab_zeq_estru_autop_estai").Row
Rows(PrimeiraLinha + 1 & ":" & PrimeiraLinha + 1).Select
Range(Selection, Selection.End(xlDown)).Delete Shift:=xlUp

QtdeTorresTabTorresBD = Application.WorksheetFunction.CountA(Range("BASE_BD_TorresLTAutopEstai[numero_torre]"))

ActiveSheet.ListObjects("Tab_zeq_estru_autop_estai").Resize Range("$B$" & PrimeiraLinha - 1 & ":$Z$" & 1 * QtdeTorresTabTorresBD + (PrimeiraLinha - 1))


Range("Tab_zeq_estru_autop_estai[TORRE]").Rows(1).Activate
    ActiveCell.Range("A1:A" & QtdeTorresTabTorresBD).Value = Range("BASE_BD_TorresLTAutopEstai[numero_torre]").Value
        Range("Tab_zeq_estru_autop_estai[TORRE]").Locked = True


Range("Tab_zeq_estru_autop_estai[FUNDAÇÃO PÉ]").Rows(1).Activate
    ActiveCell.Range("A1:A" & QtdeTorresTabTorresBD).Value = Range("BASE_BD_TorresLTAutopEstai[fundacao_pe1]").Value
        Dim cel As Range
        For Each cel In Range("Tab_zeq_estru_autop_estai[FUNDAÇÃO PÉ]")
            If cel.Value <> "" Then
                cel.Locked = True
            Else:
                cel.Interior.Color = 65535
            End If
        Next cel


Range("Tab_zeq_estru_autop_estai[FUNDAÇÃO MASTRO]").Rows(1).Activate
    ActiveCell.Range("A1:A" & QtdeTorresTabTorresBD).Value = Range("BASE_BD_TorresLTAutopEstai[fundacao_mastro1]").Value
        For Each cel In Range("Tab_zeq_estru_autop_estai[FUNDAÇÃO MASTRO]")
            If cel.Value <> "" Then
                cel.Locked = True
            Else:
                cel.Interior.Color = 65535
            End If
        Next cel


Range("Tab_zeq_estru_autop_estai[FUNDAÇÃO ESTAI A]").Rows(1).Activate
    ActiveCell.Range("A1:A" & QtdeTorresTabTorresBD).Value = Range("BASE_BD_TorresLTAutopEstai[fundacao_estai1]").Value
        For Each cel In Range("Tab_zeq_estru_autop_estai[FUNDAÇÃO ESTAI A]")
            If cel.Value <> "" Then
                cel.Locked = True
            Else:
                cel.Interior.Color = 65535
            End If
        Next cel


Range("Tab_zeq_estru_autop_estai[FUNDAÇÃO ESTAI B]").Rows(1).Activate
    ActiveCell.Range("A1:A" & QtdeTorresTabTorresBD).Value = Range("BASE_BD_TorresLTAutopEstai[fundacao_estai2]").Value
        For Each cel In Range("Tab_zeq_estru_autop_estai[FUNDAÇÃO ESTAI B]")
            If cel.Value <> "" Then
                cel.Locked = True
            Else:
                cel.Interior.Color = 65535
            End If
        Next cel


Range("Tab_zeq_estru_autop_estai[FUNDAÇÃO ESTAI C]").Rows(1).Activate
    ActiveCell.Range("A1:A" & QtdeTorresTabTorresBD).Value = Range("BASE_BD_TorresLTAutopEstai[fundacao_estai3]").Value
        For Each cel In Range("Tab_zeq_estru_autop_estai[FUNDAÇÃO ESTAI C]")
            If cel.Value <> "" Then
                cel.Locked = True
            Else:
                cel.Interior.Color = 65535
            End If
        Next cel


Range("Tab_zeq_estru_autop_estai[FUNDAÇÃO ESTAI D]").Rows(1).Activate
    ActiveCell.Range("A1:A" & QtdeTorresTabTorresBD).Value = Range("BASE_BD_TorresLTAutopEstai[fundacao_estai4]").Value
        For Each cel In Range("Tab_zeq_estru_autop_estai[FUNDAÇÃO ESTAI D]")
            If cel.Value <> "" Then
                cel.Locked = True
            Else:
                cel.Interior.Color = 65535
            End If
        Next cel


Sheets("zeq_estru_autop&estai").Activate
Range("A1").Activate
Sheets("Menu").Activate

Sheets("zeq_estru_autop&estai").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=False, Password:=StrConv(Base64Decode("UGFkaWxoYUgyTSo="), vbUnicode)



'**********ZEQ CADEIA DE ISOLADORES**********

Sheets("zeq_cadeia_isol").Unprotect (StrConv(Base64Decode("UGFkaWxoYUgyTSo="), vbUnicode))
Range("Tab_zeq_cadeia_isol").ClearContents
Range("Tab_zeq_cadeia_isol").Interior.Pattern = xlNone
Range("Tab_zeq_cadeia_isol").Locked = False

Sheets("zeq_cadeia_isol").Activate
PrimeiraLinha = Range("Tab_zeq_cadeia_isol").Row
Rows(PrimeiraLinha + 1 & ":" & PrimeiraLinha + 1).Select
Range(Selection, Selection.End(xlDown)).Delete Shift:=xlUp

ActiveSheet.ListObjects("Tab_zeq_cadeia_isol").Resize Range("$B$" & PrimeiraLinha - 1 & ":$K$" & 3 * QtdeTorresTabTorresBD + (PrimeiraLinha - 1))


Range("Tab_zeq_cadeia_isol[TORRE]").Rows(1).Activate
    ActiveCell.Range("A1:A" & QtdeTorresTabTorresBD).Value = Range("BASE_BD_TorresLTAutopEstai[numero_torre]").Value
Range("Tab_zeq_cadeia_isol[TORRE]").Rows(QtdeTorresTabTorresBD + 1).Activate
    ActiveCell.Range("A1:A" & QtdeTorresTabTorresBD).Value = Range("BASE_BD_TorresLTAutopEstai[numero_torre]").Value
Range("Tab_zeq_cadeia_isol[TORRE]").Rows(QtdeTorresTabTorresBD * 2 + 1).Activate
    ActiveCell.Range("A1:A" & QtdeTorresTabTorresBD).Value = Range("BASE_BD_TorresLTAutopEstai[numero_torre]").Value

Range("Tab_zeq_cadeia_isol[FASEAMENTO ELÉTRICO]").Rows(1).Activate
    ActiveCell.Range("A1:A" & QtdeTorresTabTorresBD).Value = "A"
Range("Tab_zeq_cadeia_isol[FASEAMENTO ELÉTRICO]").Rows(QtdeTorresTabTorresBD + 1).Activate
    ActiveCell.Range("A1:A" & QtdeTorresTabTorresBD).Value = "B"
Range("Tab_zeq_cadeia_isol[FASEAMENTO ELÉTRICO]").Rows(QtdeTorresTabTorresBD * 2 + 1).Activate
    ActiveCell.Range("A1:A" & QtdeTorresTabTorresBD).Value = "C"
        Range("Tab_zeq_cadeia_isol[TORRE], Tab_zeq_cadeia_isol[FASEAMENTO ELÉTRICO]").Locked = True


Sheets("zeq_cadeia_isol").Activate
Range("A1").Activate
Sheets("Menu").Activate

Sheets("zeq_cadeia_isol").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=False, Password:=StrConv(Base64Decode("UGFkaWxoYUgyTSo="), vbUnicode)



'**********ZEQ ATERRAMENTO**********

Sheets("zeq_aterramento").Unprotect (StrConv(Base64Decode("UGFkaWxoYUgyTSo="), vbUnicode))
Range("Tab_zeq_aterramento").ClearContents
Range("Tab_zeq_aterramento").Interior.Pattern = xlNone
Range("Tab_zeq_aterramento").Locked = False

Sheets("zeq_aterramento").Activate
PrimeiraLinha = Range("Tab_zeq_aterramento").Row
Rows(PrimeiraLinha + 1 & ":" & PrimeiraLinha + 1).Select
Range(Selection, Selection.End(xlDown)).Delete Shift:=xlUp

ActiveSheet.ListObjects("Tab_zeq_aterramento").Resize Range("$B$" & PrimeiraLinha - 1 & ":$I$" & 1 * QtdeTorresTabTorresBD + (PrimeiraLinha - 1))


Range("Tab_zeq_aterramento[TORRE]").Rows(1).Activate
    ActiveCell.Range("A1:A" & QtdeTorresTabTorresBD).Value = Range("BASE_BD_TorresLTAutopEstai[numero_torre]").Value
        Range("Tab_zeq_aterramento[TORRE]").Locked = True

Range("Tab_zeq_aterramento[CONFIGURAÇÃO DE ATERRAMENTO]").Rows(1).Activate
    ActiveCell.Range("A1:A" & QtdeTorresTabTorresBD).Value = Range("BASE_BD_TorresLTAutopEstai[nome_fase_aterramento]").Value
        Range("Tab_zeq_aterramento[CONFIGURAÇÃO DE ATERRAMENTO]").Locked = True

Range("Tab_zeq_aterramento[TIPO DE CABO CONTRAPESO]").Rows(1).Activate
    ActiveCell.Range("A1:A" & QtdeTorresTabTorresBD).Value = Range("BASE_BD_TorresLTAutopEstai[NomeCabo]").Value
        Range("Tab_zeq_aterramento[TIPO DE CABO CONTRAPESO]").Locked = True

Range("Tab_zeq_aterramento[COMP TOT CABO CONTRAPESO (m)]").Rows(1).Activate
    ActiveCell.Range("A1:A" & QtdeTorresTabTorresBD).Value = Range("BASE_BD_TorresLTAutopEstai[Comprimento_Contrapeso]").Value
        For Each cel In Range("Tab_zeq_aterramento[COMP TOT CABO CONTRAPESO (m)]")
            If cel.Value <> "" Then
                cel.Locked = True
            Else:
                cel.Interior.Color = 65535
            End If
        Next cel


Sheets("zeq_aterramento").Activate
Range("A1").Activate
Sheets("Menu").Activate

Sheets("zeq_aterramento").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=False, Password:=StrConv(Base64Decode("UGFkaWxoYUgyTSo="), vbUnicode)




'**********ZEQ ACESSOS**********

Sheets("zeq_acessos").Unprotect (StrConv(Base64Decode("UGFkaWxoYUgyTSo="), vbUnicode))
Range("Tab_zeq_acessos").ClearContents
Range("Tab_zeq_acessos").Interior.Pattern = xlNone
Range("Tab_zeq_acessos").Locked = False

Sheets("zeq_acessos").Activate
PrimeiraLinha = Range("Tab_zeq_acessos").Row
Rows(PrimeiraLinha + 1 & ":" & PrimeiraLinha + 1).Select
Range(Selection, Selection.End(xlDown)).Delete Shift:=xlUp

ActiveSheet.ListObjects("Tab_zeq_acessos").Resize Range("$B$" & PrimeiraLinha - 1 & ":$F$" & 1 * QtdeTorresTabTorresBD + (PrimeiraLinha - 1))


Range("Tab_zeq_acessos[TORRE]").Rows(1).Activate
    ActiveCell.Range("A1:A" & QtdeTorresTabTorresBD).Value = Range("BASE_BD_TorresLTAutopEstai[numero_torre]").Value

    Range("Tab_zeq_acessos[TORRE]").Locked = True


Sheets("zeq_acessos").Activate
Range("A1").Activate
Sheets("Menu").Activate

Sheets("zeq_acessos").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=False, Password:=StrConv(Base64Decode("UGFkaWxoYUgyTSo="), vbUnicode)




'**********ZEQ CONDUTOR**********

Sheets("zeq_condutor").Unprotect (StrConv(Base64Decode("UGFkaWxoYUgyTSo="), vbUnicode))
Range("Tab_zeq_condutor").ClearContents
Range("Tab_zeq_condutor").Interior.Pattern = xlNone
Range("Tab_zeq_condutor").Locked = False

Sheets("zeq_condutor").Activate
PrimeiraLinha = Range("Tab_zeq_condutor").Row
Rows(PrimeiraLinha + 1 & ":" & PrimeiraLinha + 1).Select
Range(Selection, Selection.End(xlDown)).Delete Shift:=xlUp

QtdeVaosTabVaosBD = Application.WorksheetFunction.CountA(Range("BASE_BD_VaosLT[identificacao_vao]"))

ActiveSheet.ListObjects("Tab_zeq_condutor").Resize Range("$B$" & PrimeiraLinha - 1 & ":$O$" & 3 * QtdeVaosTabVaosBD + (PrimeiraLinha - 1))


Range("Tab_zeq_condutor[VÃO]").Rows(1).Activate
    ActiveCell.Range("A1:A" & QtdeVaosTabVaosBD).Value = Range("BASE_BD_VaosLT[identificacao_vao]").Value
Range("Tab_zeq_condutor[VÃO]").Rows(QtdeVaosTabVaosBD + 1).Activate
    ActiveCell.Range("A1:A" & QtdeVaosTabVaosBD).Value = Range("BASE_BD_VaosLT[identificacao_vao]").Value
Range("Tab_zeq_condutor[VÃO]").Rows(QtdeVaosTabVaosBD * 2 + 1).Activate
    ActiveCell.Range("A1:A" & QtdeVaosTabVaosBD).Value = Range("BASE_BD_VaosLT[identificacao_vao]").Value

Range("Tab_zeq_condutor[FASE]").Rows(1).Activate
    ActiveCell.Range("A1:A" & QtdeVaosTabVaosBD).Value = "A"
Range("Tab_zeq_condutor[FASE]").Rows(QtdeVaosTabVaosBD + 1).Activate
    ActiveCell.Range("A1:A" & QtdeVaosTabVaosBD).Value = "B"
Range("Tab_zeq_condutor[FASE]").Rows(QtdeVaosTabVaosBD * 2 + 1).Activate
    ActiveCell.Range("A1:A" & QtdeVaosTabVaosBD).Value = "C"

Range("Tab_zeq_condutor[TIPO CABO CONDUTOR I]").Rows(1).Activate
    ActiveCell.Range("A1:A" & QtdeVaosTabVaosBD).Value = Range("BASE_BD_VaosLT[NomeCabo]").Value
Range("Tab_zeq_condutor[TIPO CABO CONDUTOR I]").Rows(QtdeVaosTabVaosBD + 1).Activate
    ActiveCell.Range("A1:A" & QtdeVaosTabVaosBD).Value = Range("BASE_BD_VaosLT[NomeCabo]").Value
Range("Tab_zeq_condutor[TIPO CABO CONDUTOR I]").Rows(QtdeVaosTabVaosBD * 2 + 1).Activate
    ActiveCell.Range("A1:A" & QtdeVaosTabVaosBD).Value = Range("BASE_BD_VaosLT[NomeCabo]").Value

Range("Tab_zeq_condutor[TIPO CABO CONDUTOR II]").Value = "-"

Range("Tab_zeq_condutor[QTDE. SUB-CONDUTORES]").Rows(1).Activate
    ActiveCell.Range("A1:A" & QtdeVaosTabVaosBD).Value = Range("BASE_BD_VaosLT[quantidade_subcondutores_fase]").Value
Range("Tab_zeq_condutor[QTDE. SUB-CONDUTORES]").Rows(QtdeVaosTabVaosBD + 1).Activate
    ActiveCell.Range("A1:A" & QtdeVaosTabVaosBD).Value = Range("BASE_BD_VaosLT[quantidade_subcondutores_fase]").Value
Range("Tab_zeq_condutor[QTDE. SUB-CONDUTORES]").Rows(QtdeVaosTabVaosBD * 2 + 1).Activate
    ActiveCell.Range("A1:A" & QtdeVaosTabVaosBD).Value = Range("BASE_BD_VaosLT[quantidade_subcondutores_fase]").Value

        Range("Tab_zeq_condutor[VÃO], Tab_zeq_condutor[FASE], Tab_zeq_condutor[TIPO CABO CONDUTOR I], Tab_zeq_condutor[TIPO CABO CONDUTOR II], Tab_zeq_condutor[QTDE. SUB-CONDUTORES]").Locked = True


Sheets("zeq_condutor").Activate
Range("A1").Activate
Sheets("Menu").Activate

Sheets("zeq_condutor").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=False, Password:=StrConv(Base64Decode("UGFkaWxoYUgyTSo="), vbUnicode)




'**********ZEQ PARA-RAIOS**********

Sheets("zeq_pararaio").Unprotect (StrConv(Base64Decode("UGFkaWxoYUgyTSo="), vbUnicode))
Range("Tab_zeq_pararaio").ClearContents
Range("Tab_zeq_pararaio").Interior.Pattern = xlNone
Range("Tab_zeq_pararaio").Locked = False

Sheets("zeq_pararaio").Activate
PrimeiraLinha = Range("Tab_zeq_pararaio").Row
Rows(PrimeiraLinha + 1 & ":" & PrimeiraLinha + 1).Select
Range(Selection, Selection.End(xlDown)).Delete Shift:=xlUp

ActiveSheet.ListObjects("Tab_zeq_pararaio").Resize Range("$B$" & PrimeiraLinha - 1 & ":$M$" & 3 * QtdeVaosTabVaosBD + (PrimeiraLinha - 1))

Range("Tab_zeq_pararaio[VÃO]").Rows(1).Activate
    ActiveCell.Range("A1:A" & QtdeVaosTabVaosBD).Value = Range("BASE_BD_VaosLT[identificacao_vao]").Value
Range("Tab_zeq_pararaio[VÃO]").Rows(QtdeVaosTabVaosBD + 1).Activate
    ActiveCell.Range("A1:A" & QtdeVaosTabVaosBD).Value = Range("BASE_BD_VaosLT[identificacao_vao]").Value
Range("Tab_zeq_pararaio[VÃO]").Rows(QtdeVaosTabVaosBD * 2 + 1).Activate
    ActiveCell.Range("A1:A" & QtdeVaosTabVaosBD).Value = Range("BASE_BD_VaosLT[identificacao_vao]").Value

Range("Tab_zeq_pararaio[LADO]").Rows(1).Activate
    ActiveCell.Range("A1:A" & QtdeVaosTabVaosBD).Value = "Esquerdo/Central"
Range("Tab_zeq_pararaio[LADO]").Rows(QtdeVaosTabVaosBD + 1).Activate
    ActiveCell.Range("A1:A" & QtdeVaosTabVaosBD).Value = "Direito"
Range("Tab_zeq_pararaio[LADO]").Rows(QtdeVaosTabVaosBD * 2 + 1).Activate
    ActiveCell.Range("A1:A" & QtdeVaosTabVaosBD).Value = "Indefinido"

Range("Tab_zeq_pararaio[TIPO DO CABO]").Rows(1).Activate
    ActiveCell.Range("A1:A" & QtdeVaosTabVaosBD).Value = Range("BASE_BD_VaosLT[PR_esquerdo]").Value
Range("Tab_zeq_pararaio[TIPO DO CABO]").Rows(QtdeVaosTabVaosBD + 1).Activate
    ActiveCell.Range("A1:A" & QtdeVaosTabVaosBD).Value = Range("BASE_BD_VaosLT[PR_direito]").Value
Range("Tab_zeq_pararaio[TIPO DO CABO]").Rows(QtdeVaosTabVaosBD * 2 + 1).Activate
    ActiveCell.Range("A1:A" & QtdeVaosTabVaosBD).Value = Range("BASE_BD_VaosLT[PR_indefinido]").Value

        Range("Tab_zeq_pararaio[VÃO], Tab_zeq_pararaio[LADO], Tab_zeq_pararaio[TIPO DO CABO]").Locked = True

Sheets("zeq_pararaio").Activate
Range("A1").Activate
Sheets("Menu").Activate

Sheets("zeq_pararaio").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=False, Password:=StrConv(Base64Decode("UGFkaWxoYUgyTSo="), vbUnicode)




'**********ZEQ OPGW**********

Sheets("zeq_opgw").Unprotect (StrConv(Base64Decode("UGFkaWxoYUgyTSo="), vbUnicode))
Range("Tab_zeq_opgw").ClearContents
Range("Tab_zeq_opgw").Interior.Pattern = xlNone
Range("Tab_zeq_opgw").Locked = False

Sheets("zeq_opgw").Activate
PrimeiraLinha = Range("Tab_zeq_opgw").Row
Rows(PrimeiraLinha + 1 & ":" & PrimeiraLinha + 1).Select
Range(Selection, Selection.End(xlDown)).Delete Shift:=xlUp

ActiveSheet.ListObjects("Tab_zeq_opgw").Resize Range("$B$" & PrimeiraLinha - 1 & ":$H$" & 3 * QtdeVaosTabVaosBD + (PrimeiraLinha - 1))

Range("Tab_zeq_opgw[VÃO]").Rows(1).Activate
    ActiveCell.Range("A1:A" & QtdeVaosTabVaosBD).Value = Range("BASE_BD_VaosLT[identificacao_vao]").Value
Range("Tab_zeq_opgw[VÃO]").Rows(QtdeVaosTabVaosBD + 1).Activate
    ActiveCell.Range("A1:A" & QtdeVaosTabVaosBD).Value = Range("BASE_BD_VaosLT[identificacao_vao]").Value
Range("Tab_zeq_opgw[VÃO]").Rows(QtdeVaosTabVaosBD * 2 + 1).Activate
    ActiveCell.Range("A1:A" & QtdeVaosTabVaosBD).Value = Range("BASE_BD_VaosLT[identificacao_vao]").Value

Range("Tab_zeq_opgw[LADO]").Rows(1).Activate
    ActiveCell.Range("A1:A" & QtdeVaosTabVaosBD).Value = "Esquerdo/Central"
Range("Tab_zeq_opgw[LADO]").Rows(QtdeVaosTabVaosBD + 1).Activate
    ActiveCell.Range("A1:A" & QtdeVaosTabVaosBD).Value = "Direito"
Range("Tab_zeq_opgw[LADO]").Rows(QtdeVaosTabVaosBD * 2 + 1).Activate
    ActiveCell.Range("A1:A" & QtdeVaosTabVaosBD).Value = "Indefinido"

Range("Tab_zeq_opgw[FABRICANTE OPGW]").Rows(1).Activate
    ActiveCell.Range("A1:A" & QtdeVaosTabVaosBD).Value = Range("BASE_BD_VaosLT[OPGW_Esquerdo.Fabricante]").Value
Range("Tab_zeq_opgw[FABRICANTE OPGW]").Rows(QtdeVaosTabVaosBD + 1).Activate
    ActiveCell.Range("A1:A" & QtdeVaosTabVaosBD).Value = Range("BASE_BD_VaosLT[OPGW_Direito.Fabricante]").Value
Range("Tab_zeq_opgw[FABRICANTE OPGW]").Rows(QtdeVaosTabVaosBD * 2 + 1).Activate
    ActiveCell.Range("A1:A" & QtdeVaosTabVaosBD).Value = Range("BASE_BD_VaosLT[OPGW_Indefinido.Fabricante]").Value

Range("Tab_zeq_opgw[NÚMERO DE FIBRAS]").Rows(1).Activate
    ActiveCell.Range("A1:A" & QtdeVaosTabVaosBD).Value = Range("BASE_BD_VaosLT[OPGW_Esquerdo.NumFibras]").Value
Range("Tab_zeq_opgw[NÚMERO DE FIBRAS]").Rows(QtdeVaosTabVaosBD + 1).Activate
    ActiveCell.Range("A1:A" & QtdeVaosTabVaosBD).Value = Range("BASE_BD_VaosLT[OPGW_Direito.NumFibras]").Value
Range("Tab_zeq_opgw[NÚMERO DE FIBRAS]").Rows(QtdeVaosTabVaosBD * 2 + 1).Activate
    ActiveCell.Range("A1:A" & QtdeVaosTabVaosBD).Value = Range("BASE_BD_VaosLT[OPGW_Indefinido.NumFibras]").Value

Range("Tab_zeq_opgw[SEÇÃO (mm²)]").Rows(1).Activate
    ActiveCell.Range("A1:A" & QtdeVaosTabVaosBD).Value = Range("BASE_BD_VaosLT[OPGW_Esquerdo.Seção]").Value
Range("Tab_zeq_opgw[SEÇÃO (mm²)]").Rows(QtdeVaosTabVaosBD + 1).Activate
    ActiveCell.Range("A1:A" & QtdeVaosTabVaosBD).Value = Range("BASE_BD_VaosLT[OPGW_Direito.Seção]").Value
Range("Tab_zeq_opgw[SEÇÃO (mm²)]").Rows(QtdeVaosTabVaosBD * 2 + 1).Activate
    ActiveCell.Range("A1:A" & QtdeVaosTabVaosBD).Value = Range("BASE_BD_VaosLT[OPGW_Indefinido.Seção]").Value

Range("Tab_zeq_opgw[DIÂMETRO (mm)]").Rows(1).Activate
    ActiveCell.Range("A1:A" & QtdeVaosTabVaosBD).Value = Range("BASE_BD_VaosLT[OPGW_Esquerdo.Diâmetro]").Value
Range("Tab_zeq_opgw[DIÂMETRO (mm)]").Rows(QtdeVaosTabVaosBD + 1).Activate
    ActiveCell.Range("A1:A" & QtdeVaosTabVaosBD).Value = Range("BASE_BD_VaosLT[OPGW_Direito.Diâmetro]").Value
Range("Tab_zeq_opgw[DIÂMETRO (mm)]").Rows(QtdeVaosTabVaosBD * 2 + 1).Activate
    ActiveCell.Range("A1:A" & QtdeVaosTabVaosBD).Value = Range("BASE_BD_VaosLT[OPGW_Indefinido.Diâmetro]").Value


        Range("Tab_zeq_opgw[VÃO], Tab_zeq_opgw[LADO], Tab_zeq_opgw[FABRICANTE OPGW]").Locked = True
        Range("Tab_zeq_opgw[NÚMERO DE FIBRAS], Tab_zeq_opgw[SEÇÃO (mm²)], Tab_zeq_opgw[DIÂMETRO (mm)]").Locked = True


Sheets("zeq_opgw").Activate
Range("A1").Activate
Sheets("Menu").Activate

Sheets("zeq_opgw").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=False, Password:=StrConv(Base64Decode("UGFkaWxoYUgyTSo="), vbUnicode)




'**********ZEQ SERVIDÃO**********

Sheets("zeq_servidao").Unprotect (StrConv(Base64Decode("UGFkaWxoYUgyTSo="), vbUnicode))
Range("Tab_zeq_servidao").ClearContents
Range("Tab_zeq_servidao").Interior.Pattern = xlNone
Range("Tab_zeq_servidao").Locked = False

Sheets("zeq_servidao").Activate
PrimeiraLinha = Range("Tab_zeq_servidao").Row
Rows(PrimeiraLinha + 1 & ":" & PrimeiraLinha + 1).Select
Range(Selection, Selection.End(xlDown)).Delete Shift:=xlUp

ActiveSheet.ListObjects("Tab_zeq_servidao").Resize Range("$B$" & PrimeiraLinha - 1 & ":$R$" & 1 * QtdeVaosTabVaosBD + (PrimeiraLinha - 1))

Range("Tab_zeq_servidao[VÃO]").Rows(1).Activate
    ActiveCell.Range("A1:A" & QtdeVaosTabVaosBD).Value = Range("BASE_BD_VaosLT[identificacao_vao]").Value


        Range("Tab_zeq_servidao[VÃO]").Locked = True


Sheets("zeq_servidao").Activate
Range("A1").Activate
Sheets("Menu").Activate

Sheets("zeq_servidao").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=False, Password:=StrConv(Base64Decode("UGFkaWxoYUgyTSo="), vbUnicode)




'***AVISOS FINAIS***
   
    If Aviso1 Then
        MsgAviso1 = MsgBox("Foi identificado distinções entre os projetos nas temperaturas das capacidades operativas, " & _
        "Portanto, os dados não foram preenchidos automaticamente", vbExclamation, "Capacidade operativa")
    End If

    If Aviso2 Then
        MsgAviso2 = MsgBox("Devido a existência de cadastro de mais de um projeto ativo para a LT, " & LT_CodLT & _
            "não foi possível obter um único valor para preenchimento automático de flechas e parâmetros elétricos.", vbExclamation, "Capacidade operativa")
    End If


Exit Sub 'TEMPORÁRIO




Application.DisplayAlerts = False

Application.Workbooks.Open (LC_CaminhoLC)

Application.DisplayAlerts = True


On Error GoTo Continue1

    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select

Continue1:

On Error GoTo -1
On Error GoTo 0


Windows(LC_NomeLC).Activate

 

    ActiveSheet.Next.Select '*TABELA RESUMO*
    
    Range("A1").Select
    Cells.Find(What:="Nomenclatura do cabo condutor", After:=ActiveCell, _
        LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False).Activate
    ActiveCell.Offset(0, 1).Range("A1").Select
    LC_NomeCondutor = ActiveCell.Value
        
        If ActiveCell.Offset(0, 1).Range("A1").Value <> "" Then
            LC_NomeCondutor = "{VARIÁVEL}"
        End If


    Range("A1").Select
    Cells.Find(What:="Subcondutor", After:=ActiveCell, LookIn:=xlValues, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    ActiveCell.Offset(0, 1).Range("A1").Select
    LC_NumSubcondutores = ActiveCell.Value
    
        If LC_NomeCondutor = "{VARIÁVEL}" Then
            LC_NumSubcondutores = "{VARIÁVEL}"
        End If

'top demais

End Sub

