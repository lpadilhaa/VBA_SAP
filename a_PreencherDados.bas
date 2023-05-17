Public VBA_SAP_Versao As String
Public BaseVBA_SAP As String
Public LT_CodLT As String
Public LT_NomeLT As String

Sub Preencher_Dados()


If Range("Label_NomeLT").Locked = True Then
    VBA_SAP_Versao = Mid(ThisWorkbook.Name, InStr(ThisWorkbook.Name, "_v") + 2, (InStr(InStr(ThisWorkbook.Name, "_v") + 2, ThisWorkbook.Name, "_") - 1) - (InStr(ThisWorkbook.Name, "_v") + 2) + 1)
    Else: VBA_SAP_Versao = Replace(Right(ThisWorkbook.Name, Len(ThisWorkbook.Name) - InStrRev(ThisWorkbook.Name, "v")), ".xlsm", "")
End If


Dim MsgInicial As VbMsgBoxResult
MsgInicial = MsgBox("Os seguintes dados serão solicitados:" & vbCrLf & vbCrLf & _
            " - Lista de Construção;" & vbCrLf & _
            " - Dados BDIT;" & vbCrLf & _
            " - Base auxiliar preenchida (arranjos, fundações e travessias)." & vbCrLf & vbCrLf & _
                "Deseja iniciar o processo de preenchimento automático dos dados SAP?" _
                    , vbInformation + vbYesNo + vbDefaultButton2, "Iniciar preenchimento?")
    
    If MsgInicial = vbNo Then
        CancelaPreenchimento = MsgBox("Processo de preenchimento automático cancelado!", vbCritical + vbOKOnly, "Preenchimento cancelado!")
        Exit Sub
    End If


If Range("Label_NomeLT").Locked = True Then
    Dim QuestionDownload As VbMsgBoxResult
        QuestionDownload = MsgBox("Uma vez que já houve atualizações anteriores, as informações disponíveis nas bases online e no banco de dados já estão acessíveis." & _
            " Portanto, considere a necessidade de baixar novamente somente se há atualizações nas bases ou no banco." & Chr(13) & Chr(13) & _
                "Deseja baixar os dados novamente?", vbExclamation + vbYesNo + vbDefaultButton2, "Dados já disponíveis")
    GoTo Continuar
End If

Dim MsgInicial_reask As VbMsgBoxResult
MsgInicial_reask = MsgBox("Para iniciar o procedimento, antes, serão realizadas consultas no Banco de Dados." & vbCrLf & vbCrLf & _
                        "ATENÇÃO: A consulta no banco de dados poderá levar vários minutos. É importante que a conexão com a internet seja estável." & vbCrLf & vbCrLf & _
                            "Aperte ""OK"" para iniciar, ou ""CANCELAR"" para sair." _
                                , vbExclamation + vbOKCancel + vbDefaultButton2, "Iniciar preenchimento: certeza?")

    If MsgInicial_reask = vbCancel Then
        CancelaPreenchimento = MsgBox("Processo de preenchimento automático cancelado!", vbCritical + vbOKOnly, "Preenchimento cancelado!")
        Exit Sub
    End If

Continuar:

Application.ScreenUpdating = False

BaseVBA_SAP = Application.ActiveWorkbook.Name

LT_CodLT = Range("Label_CodLT").Value
LT_NomeLT = Range("Label_NomeLT").Value
LT_TensaoLT = Application.WorksheetFunction.VLookup(Range("Label_NomeLT"), Range("BASE_LTs"), 14, False)

    
'**********ATUALIZANDO BASES**********


If QuestionDownload = vbNo Then
    GoTo Inicio_Selecionar_LC
End If


Atualizar_Bases:

On Error GoTo -1
On Error GoTo 0
On Error GoTo ErrorDownload


    ActiveWorkbook.Queries.Item("Param_CodLT").Formula = _
        """" & LT_CodLT & """ meta [IsParameterQuery=true, Type=""Any"", IsParameterQueryRequired=True]"

    ActiveWorkbook.Connections("Consulta - BASE_LTs").Refresh
    ActiveWorkbook.Connections("Consulta - BASE_Cabos").Refresh
    ActiveWorkbook.Connections("Consulta - BASE_CabosWithOPGW").Refresh
    ActiveWorkbook.Connections("Consulta - BASE_Tracoes").Refresh  '/incluído posteriormente
    ActiveWorkbook.Connections("Consulta - BASE_BD_OPGWLT").Refresh
    ActiveWorkbook.Connections("Consulta - BASE_BD_SerieEstrutura").Refresh
    ActiveWorkbook.Connections("Consulta - BASE_BD_Aterramento").Refresh
    ActiveWorkbook.Connections("Consulta - BASE_BD_ParaRaiosLT").Refresh
    ActiveWorkbook.Connections("Consulta - BASE_BD_TorresLTGeral").Refresh
    ActiveWorkbook.Connections("Consulta - BASE_BD_TorresLTAutopEstai").Refresh
    Range("BASE_BD_ProjetosLT").ListObject.QueryTable.Refresh
    ActiveWorkbook.Connections("Consulta - BASE_BD_VaosLT").Refresh

    MsgConsultasOK = MsgBox("Consultas concluídas com sucesso. As informações do Banco de Dados foram armazenadas.", vbInformation, "Consultas: OK")

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

If Application.WorksheetFunction.CountIf(Range("BASE_BD_TorresLTGeral[caracteristica2]"), "Estaiada") >= 1 Then

    Dim PossuiEstaiadas As VbMsgBoxResult
        PossuiEstaiadas = MsgBox("A LT selecionada possui torres estaiadas, qual o código para preenchimento automático dos dados ainda está em desenvolvimento. " & _
                "Portanto, ainda não será possível preencher os dados da LT " & LT_CodLT & "." _
        & Chr(13) & Chr(13) & _
        "Selecione outra LT ou aguarde a finalização do código.", vbCritical, "Torres estaiadas - em desenvolvimento!")
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


Inicio_Selecionar_LC:

Call Preencher_Dados2

End Sub


Sub Preencher_Dados2()


'**********SELECIONANDO LISTA DE CONSTRUÇÃO**********


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

On Error GoTo err1

    LC_CodLT = Range("B6").Value
    LC_NomeLT = Range("A6").Value
    LC_CaminhoLC = Application.ActiveWorkbook.FullName
    LC_NomeLC = Application.ActiveWorkbook.Name
    LC_Extensao = Round(Application.WorksheetFunction.Sum(Range("ListadeConstrucao[Vao]")) / 1000, 2)
    LC_QtdeEstruturas = Application.WorksheetFunction.CountA(Range("ListadeConstrucao[NumOper]"))
    LC_SomaFaixaServidao = Application.WorksheetFunction.Sum(Range("ListadeConstrucao[[FaixaServ1]:[FaixaServ2]]"))

On Error GoTo -1
On Error GoTo 0

Application.ActiveWorkbook.Close (savechanges = True)
GoTo ctn1

err1:
Application.ActiveWorkbook.Close (savechanges = True)
LC_Selected = "NOK"
On Error GoTo -1
On Error GoTo 0

ctn1:

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
    MsgBox Range("BASE_BD_ProjetosLT[extensao_total_linha]").Rows(1).Value
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

If LC_SomaFaixaServidao <> Application.WorksheetFunction.Sum(Range("BASE_BD_VaosLT[largura_faixa_servidao]")) Then
    
    Dim SomaFaixaServidao_Errada As VbMsgBoxResult
        SomaFaixaServidao_Errada = MsgBox("A soma das faixas de servidão, esquerda e direita, da lista de construção não coincide com a soma das faixas de servidão registradas no banco de dados." _
        & Chr(13) & Chr(13) & _
        "Revise os dados, prossiga com as devidas correções e tente novamente.", vbCritical, "Soma das faixas de servidão incorreta!")
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




'**********SELECIONANDO A BASE AUXILIAR PREENCHIDA**********


Inicio_Selecionar_BaseAux:

Dim Msg_SelectBaseAux As VbMsgBoxResult
Msg_SelectBaseAux = MsgBox("Selecione a Base Auxiliar, preenchida, da LT " & LT_CodLT & ":", vbOKCancel + vbExclamation, "Selecionar Base Auxiliar")

If Msg_SelectBaseAux = vbCancel Then
CancelaPreenchimento = MsgBox("Processo de preenchimento automático cancelado!", vbCritical + vbOKOnly, "Preenchimento cancelado!")
Exit Sub
End If


Selecionar_BaseAux:

On Error GoTo -1
On Error GoTo 0

Application.DisplayAlerts = False

Set Select_BaseAux = Application.FileDialog(msoFileDialogOpen)

With Select_BaseAux
    .AllowMultiSelect = False
    .Show
    If .SelectedItems.Count > 0 Then
        BaseAux_Caminho = .SelectedItems(1)
        Workbooks.Open BaseAux_Caminho, UpdateLinks:=False
    End If
End With

BaseAux_Nome = ActiveWorkbook.Name
BaseAux_Caminho = ActiveWorkbook.FullName

Application.DisplayAlerts = True

    If Application.ActiveWorkbook.Name = BaseVBA_SAP Then
        CancelaPreenchimento = MsgBox("Processo de preenchimento automático cancelado!", vbCritical + vbOKOnly, "Preenchimento cancelado!")
        Exit Sub
    End If


On Error GoTo err_baseaux
    Sheets("Início").Activate
    If Range("Label_CodLT") = "" Or Range("Label_NomeLT") = "" Or Range("E6") <> "Dados da LT" Or ActiveWorkbook.ActiveSheet.Name <> "Início" Then
        GoTo err_baseaux
    End If
On Error GoTo -1
On Error GoTo 0


On Error GoTo incorreta_baseaux
    If Range("Label_CodLT") <> LT_CodLT Or Range("Label_NomeLT") <> LT_NomeLT Then
        BaseAux_CodLT = Range("Label_CodLT")
        GoTo incorreta_baseaux
    End If
On Error GoTo -1
On Error GoTo 0


Application.DisplayAlerts = False
Workbooks(BaseAux_Nome).Close
Application.DisplayAlerts = True

GoTo BaseAux_OK


err_baseaux:

Application.DisplayAlerts = False
Workbooks(BaseAux_Nome).Close
Application.DisplayAlerts = True

Dim BaseAuxErro As VbMsgBoxResult
BaseAuxErro = MsgBox("O arquivo selecionado não corresponde à um modelo da Base Auxiliar SAP. Selecione o arquivo correto.", vbCritical + vbRetryCancel, "Base Auxiliar não encontrada!")

    If BaseAuxErro = vbCancel Then
            CancelaPreenchimento = MsgBox("Processo de preenchimento automático cancelado!", vbCritical + vbOKOnly, "Preenchimento cancelado!")
        Exit Sub
    ElseIf BaseAuxErro = vbRetry Then
        GoTo Selecionar_BaseAux
    End If


incorreta_baseaux:

Application.DisplayAlerts = False
Workbooks("BaseAux_Nome").Close
Application.DisplayAlerts = True

Dim BaseAuxIncorreta As VbMsgBoxResult
BaseAuxIncorreta = MsgBox("Foi selecionado o arquivo da LT " & BaseAux_CodLT & ", enquanto era esperado o arquivo da LT " & LT_CodLT & _
    Chr(13) & Chr(13) & ". Para prosseguir, selecione o arquivo da LT " & LT_CodLT & ".", vbCritical + vbRetryCancel, "Base Auxiliar incorreta!")

    If BaseAuxIncorreta = vbCancel Then
            CancelaPreenchimento = MsgBox("Processo de preenchimento automático cancelado!", vbCritical + vbOKOnly, "Preenchimento cancelado!")
        Exit Sub
    ElseIf BaseAuxIncorreta = vbRetry Then
        GoTo Selecionar_BaseAux
    End If



BaseAux_OK:

Msg_BaseAuxOK = MsgBox("Base auxiliar da LT " & LT_TensaoLT & " kV - " & LT_CodLT & " (" & LT_NomeLT & ") identificada!" _
& Chr(13) & Chr(13) & "Os dados foram armazenados!", vbInformation, "Base Auxiliar " & LT_CodLT & " Identificada!")




'**********SELECIONANDO DIRETÓRIO DE SALVAMENTO**********

If Range("Label_NomeLT").Locked = True Then
    filePath = ThisWorkbook.FullName
    Filename = Right(filePath, Len(filePath) - InStrRev(filePath, "\"))
    DiretorioSave = Replace(filePath, "\" & Filename, "")
    GoTo Proced_ConfirmacaoFinal
End If


SelecionarDiretorioSalvamento:
    
    Dim Msg_SelectDiretorioSave As VbMsgBoxResult
    Msg_SelectDiretorioSave = MsgBox("Por fim, selecione a pasta que deseja salvar os Dados SAP da LT " & LT_CodLT & ":", vbOKCancel + vbInformation, "Selecionar diretório de salvamento:")
    
    If Msg_SelectDiretorioSave = vbCancel Then
        CancelaPreenchimento = MsgBox("Processo de preenchimento automático cancelado!", vbCritical + vbOKOnly, "Preenchimento cancelado!")
    Exit Sub
    End If


Selecionar_DiretorioSave:

    On Error GoTo -1
    On Error GoTo 0
    On Error GoTo Msg_DiretorioNull
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Diretório de salvamento"
        .InitialFileName = ThisWorkbook.Path
        .Show
        .AllowMultiSelect = False
        DiretorioSave = .SelectedItems(1)
    End With


Msg_DiretorioNull:
    
    On Error GoTo -1
    On Error GoTo 0
    
        If DiretorioSave = "" Then
            CancelaPreenchimento = MsgBox("Processo de preenchimento automático cancelado!", vbCritical + vbOKOnly, "Preenchimento cancelado!")
            Exit Sub
        End If
    
    
    Dim Msg_DiretorioSaveSelected As VbMsgBoxResult
    
    Msg_DiretorioSaveSelected = MsgBox("Pasta selecionada:" _
    & Chr(13) & Chr(13) & _
    DiretorioSave & "\" _
    & Chr(13) & Chr(13) & _
    "Deseja salvar a os Dados SAP da LT " & LT_CodLT & " nesta pasta?", vbQuestion + vbYesNoCancel, "Confirma diretório de salvamento?")
    
    
    If Msg_DiretorioSaveSelected = vbCancel Then
        CancelaImportação = MsgBox("Processo de importação cancelado!", vbCritical + vbOKOnly)
        Exit Sub
    ElseIf Msg_DiretorioSaveSelected = vbNo Then
        GoTo Selecionar_DiretorioSave
    ElseIf Msg_DiretorioSaveSelected = vbYes Then
    End If



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

Range("Label_ZLI_Transmissao_TemperLongaDur") = Range("BASE_BD_ProjetosLT[Temp_ref_LD (°C)]").Rows(1).Value
Range("Label_ZLI_Transmissao_TemperCurtaDur") = Range("BASE_BD_ProjetosLT[Temp_ref_CD (°C)]").Rows(1).Value
Range("Label_ZLI_Transmissao_ExtensPropria") = Range("BASE_BD_ProjetosLT[extensao_total_linha]").Rows(1).Value
Range("Label_ZLI_Transmissao_QtdeEstruturas") = Range("BASE_BD_ProjetosLT[qtde_total_estruturas]").Rows(1).Value
Range("Label_ZLI_Transmissao_ModeloTorreTipica") = Range("BASE_BD_ProjetosLT[nome_estrutura_tipica]").Rows(1).Value

    
    Dim cel As Range
        For Each cel In Range("Tab_zli_li_transmissao[Coluna2]")
            If cel.Value <> "" Then
                cel.Locked = True
            Else:
                cel.Interior.Color = 65535
            End If
        Next cel


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

   
    Range("Label_ZLI_ParametrosOp_CapacOperLDVD") = Range("BASE_BD_ProjetosLT[LD_VD (A)]").Rows(1).Value
    Range("Label_ZLI_ParametrosOp_CapacOperLDVN") = Range("BASE_BD_ProjetosLT[LD_VN (A)]").Rows(1).Value
    Range("Label_ZLI_ParametrosOp_CapacOperLDID") = Range("BASE_BD_ProjetosLT[LD_ID (A)]").Rows(1).Value
    Range("Label_ZLI_ParametrosOp_CapacOperLDIN") = Range("BASE_BD_ProjetosLT[LD_IN (A)]").Rows(1).Value
    Range("Label_ZLI_ParametrosOp_CapacOperCDVD") = Range("BASE_BD_ProjetosLT[CD_VD (A)]").Rows(1).Value
    Range("Label_ZLI_ParametrosOp_CapacOperCDVN") = Range("BASE_BD_ProjetosLT[CD_VN (A)]").Rows(1).Value
    Range("Label_ZLI_ParametrosOp_CapacOperCDID") = Range("BASE_BD_ProjetosLT[CD_ID (A)]").Rows(1).Value
    Range("Label_ZLI_ParametrosOp_CapacOperCDIN") = Range("BASE_BD_ProjetosLT[CD_IN (A)]").Rows(1).Value
    
    'Range("Label_ZLI_ParametrosOp_FlechaMaxCondut") = Range("BASE_BD_ProjetosLT[flecha_cabo_condutor]").Rows(1).Value
    'Range("Label_ZLI_ParametrosOp_FlechaMaxPR") = Range("BASE_BD_ProjetosLT[flecha_cabo_para_raios]").Rows(1).Value
    
    Range("Label_ZLI_ParametrosOp_ResistSeqPosit") = Range("BASE_BD_ProjetosLT[r1_km]").Rows(1).Value
    Range("Label_ZLI_ParametrosOp_ReatSeqPosit") = Range("BASE_BD_ProjetosLT[x1_km]").Rows(1).Value
    Range("Label_ZLI_ParametrosOp_SuscepSeqPosit") = Range("BASE_BD_ProjetosLT[b1_km]").Rows(1).Value
    Range("Label_ZLI_ParametrosOp_ResistSeqZero") = Range("BASE_BD_ProjetosLT[r0_km]").Rows(1).Value
    Range("Label_ZLI_ParametrosOp_ReatSeqZero") = Range("BASE_BD_ProjetosLT[x0_km]").Rows(1).Value
    Range("Label_ZLI_ParametrosOp_SuscepSeqZero") = Range("BASE_BD_ProjetosLT[b0_km]").Rows(1).Value

    Range("Label_ZLI_ParametrosOp_CapacitSeqPosit") = Range("BASE_BD_ProjetosLT[c1]").Rows(1).Value
    Range("Label_ZLI_ParametrosOp_CapacitSeqZero") = Range("BASE_BD_ProjetosLT[c0]").Rows(1).Value


        'For Each cel In Range("Tab_zli_parametros_op[Coluna2]")
            'If cel.Value <> "" Then
                'cel.Locked = True
            'Else:
                'cel.Interior.Color = 65535
            'End If
        'Next cel


'Sheets("zli_parametros_OP").Activate
'Range("A1").Activate
'Sheets("Menu").Activate

'Sheets("zli_parametros_OP").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=False, Password:=StrConv(Base64Decode("UGFkaWxoYUgyTSo="), vbUnicode)




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
    Range("Tab_zeq_estru_geral[ALTITUDE], Tab_zeq_estru_geral[LATITUDE], Tab_zeq_estru_geral[LONGITUDE], Tab_zeq_estru_geral[DATUM]").Locked = True

Sheets("zeq_estru_geral").Activate
Range("A1").Activate
Sheets("Menu").Activate



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


Range("Tab_zeq_condutor[TRAÇÃO EDS (%)]").FormulaR1C1 = "=IFERROR(SWITCH([@[TIPO CABO CONDUTOR I]],""-"",""-"",INDEX(BASE_CabosWithOPGW[EDS (%)],MATCH([@[TIPO CABO CONDUTOR I]],BASE_CabosWithOPGW[NomeCabo],0))*100),"""")"
    Range("Tab_zeq_condutor[TRAÇÃO EDS (%)]").Value = Range("Tab_zeq_condutor[TRAÇÃO EDS (%)]").Value

        Range("Tab_zeq_condutor[VÃO], Tab_zeq_condutor[FASE], Tab_zeq_condutor[TIPO CABO CONDUTOR I], Tab_zeq_condutor[TIPO CABO CONDUTOR II], Tab_zeq_condutor[QTDE. SUB-CONDUTORES], Tab_zeq_condutor[TRAÇÃO EDS (%)]").Locked = True


Sheets("zeq_condutor").Activate
Range("A1").Activate
Sheets("Menu").Activate





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

Range("Tab_zeq_pararaio[TRAÇÃO EDS (%)]").FormulaR1C1 = "=IFERROR(SWITCH([@[TIPO DO CABO]],""-"",""-"",INDEX(BASE_CabosWithOPGW[EDS (%)],MATCH([@[TIPO DO CABO]],BASE_CabosWithOPGW[NomeCabo],0))*100),"""")"
    Range("Tab_zeq_pararaio[TRAÇÃO EDS (%)]").Value = Range("Tab_zeq_pararaio[TRAÇÃO EDS (%)]").Value

Range("Tab_zeq_pararaio[TIPO DO CABO]").Rows(1).Activate
    ActiveCell.Range("A1:A" & QtdeVaosTabVaosBD).Value = Range("BASE_BD_VaosLT[PR_esquerdo1]").Value
Range("Tab_zeq_pararaio[TIPO DO CABO]").Rows(QtdeVaosTabVaosBD + 1).Activate
    ActiveCell.Range("A1:A" & QtdeVaosTabVaosBD).Value = Range("BASE_BD_VaosLT[PR_direito1]").Value
Range("Tab_zeq_pararaio[TIPO DO CABO]").Rows(QtdeVaosTabVaosBD * 2 + 1).Activate
    ActiveCell.Range("A1:A" & QtdeVaosTabVaosBD).Value = Range("BASE_BD_VaosLT[PR_indefinido1]").Value

        Range("Tab_zeq_pararaio[VÃO], Tab_zeq_pararaio[LADO], Tab_zeq_pararaio[TIPO DO CABO], Tab_zeq_pararaio[TRAÇÃO EDS (%)]").Locked = True

Sheets("zeq_pararaio").Activate
Range("A1").Activate
Sheets("Menu").Activate





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




'**********IMPORTANDO INFORMAÇÕES DE LISTA DE CONSTRUÇÃO**********

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


Windows(BaseVBA_SAP).Activate

    Sheets("zeq_estru_geral").Range("Tab_zeq_estru_geral[NÚMERO DE PROJETO]").NumberFormat = "General"
    Sheets("zeq_estru_geral").Range("Tab_zeq_estru_geral[NÚMERO DE PROJETO]").FormulaR1C1 = "=INDEX('" & LC_NomeLC & "'!ListadeConstrucao[NumProj],MATCH([@[NÚMERO DE OPERAÇÃO]],'" & LC_NomeLC & "'!ListadeConstrucao[NumOper],0))"
    Sheets("zeq_estru_geral").Range("Tab_zeq_estru_geral[NÚMERO DE PROJETO]").NumberFormat = "@"
    Sheets("zeq_estru_geral").Range("Tab_zeq_estru_geral[NÚMERO DE PROJETO]").Value = Sheets("zeq_estru_geral").Range("Tab_zeq_estru_geral[NÚMERO DE PROJETO]").Value

    Sheets("zeq_estru_geral").Range("Tab_zeq_estru_geral[COMPRIMENTO DO VÃO (m)]").FormulaR1C1 = "=INDEX('" & LC_NomeLC & "'!ListadeConstrucao[Vao],MATCH([@[NÚMERO DE OPERAÇÃO]],'" & LC_NomeLC & "'!ListadeConstrucao[NumOper],0))"
    Sheets("zeq_estru_geral").Range("Tab_zeq_estru_geral[COMPRIMENTO DO VÃO (m)]").Value = Sheets("zeq_estru_geral").Range("Tab_zeq_estru_geral[COMPRIMENTO DO VÃO (m)]").Value

    Sheets("zeq_estru_geral").Range("Tab_zeq_estru_geral[ÂNGULO DEFLEXÃO]").FormulaR1C1 = "=INDEX('" & LC_NomeLC & "'!ListadeConstrucao[Deflexao],MATCH([@[NÚMERO DE OPERAÇÃO]],'" & LC_NomeLC & "'!ListadeConstrucao[NumOper],0))"
    Sheets("zeq_estru_geral").Range("Tab_zeq_estru_geral[ÂNGULO DEFLEXÃO]").Value = Sheets("zeq_estru_geral").Range("Tab_zeq_estru_geral[ÂNGULO DEFLEXÃO]").Value

    Sheets("zeq_estru_geral").Range("Tab_zeq_estru_geral[VÃO DE VENTO (m)]").FormulaR1C1 = "=(IFERROR(VALUE(OFFSET([@[COMPRIMENTO DO VÃO (m)]],-1,0)),0)+IFERROR(VALUE([@[COMPRIMENTO DO VÃO (m)]]),0))/2"
    Sheets("zeq_estru_geral").Range("Tab_zeq_estru_geral[VÃO DE VENTO (m)]").Value = Sheets("zeq_estru_geral").Range("Tab_zeq_estru_geral[VÃO DE VENTO (m)]").Value

    Sheets("zeq_estru_geral").Range("Tab_zeq_estru_geral[DISTÂNCIA PROGRESSIVA (m)]").FormulaR1C1 = "=IFERROR(VALUE(OFFSET([@[DISTÂNCIA PROGRESSIVA (m)]],-1,0)),0)+IFERROR(VALUE(OFFSET([@[COMPRIMENTO DO VÃO (m)]],-1,0)),0)"
    Sheets("zeq_estru_geral").Range("Tab_zeq_estru_geral[DISTÂNCIA PROGRESSIVA (m)]").Value = Sheets("zeq_estru_geral").Range("Tab_zeq_estru_geral[DISTÂNCIA PROGRESSIVA (m)]").Value

        Range("Tab_zeq_estru_geral[NÚMERO DE PROJETO], Tab_zeq_estru_geral[COMPRIMENTO DO VÃO (m)], Tab_zeq_estru_geral[ÂNGULO DEFLEXÃO], Tab_zeq_estru_geral[VÃO DE VENTO (m)], Tab_zeq_estru_geral[DISTÂNCIA PROGRESSIVA (m)]").Locked = True


Windows(LC_NomeLC).Activate
    Range("ListadeConstrucao[NumOper]").Select
    Selection.Copy
Workbooks.Add
    WBTemp = ActiveWorkbook.Name
    Range("A1").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False

Windows(LC_NomeLC).Activate
    Range("ListadeConstrucao[[NumLC]:[NumFolhaPP]]").Select
    Selection.Copy
Windows(WBTemp).Activate
    Range("B1").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False

Windows(WBTemp).Activate
    Range("F1").Formula2R1C1 = _
        "=TEXT(LEFT(SUBSTITUTE(RC[-4],""-"",""||"",LEN(RC[-4])-LEN(SUBSTITUTE(RC[-4],""-"",""""))),FIND(""||"",SUBSTITUTE(RC[-4],""-"",""||"",LEN(RC[-4])-LEN(SUBSTITUTE(RC[-4],""-"",""""))))-1),""0000000"")&""-FL.""&SUBSTITUTE(TEXT(TEXTJOIN("""",TRUE,IFERROR(MID(RC[-3],ROW(INDIRECT(""1:100"")),1)+0,"""")),""000""),TEXTJOIN("""",TRUE,IFERROR(MID(RC[-3],ROW(INDIRECT(""1:100"")" & _
        "),1)+0,"""")),RC[-3])&""/""&RIGHT(SUBSTITUTE(RC[-4],""-"",""||"",LEN(RC[-4])-LEN(SUBSTITUTE(RC[-4],""-"",""""))),LEN(SUBSTITUTE(RC[-4],""-"",""||"",LEN(RC[-4])-LEN(SUBSTITUTE(RC[-4],""-"",""""))))-FIND(""||"",SUBSTITUTE(RC[-4],""-"",""||"",LEN(RC[-4])-LEN(SUBSTITUTE(RC[-4],""-"",""""))))-1)" & _
        ""
    Range("G1").Formula2R1C1 = _
        "=TEXT(LEFT(SUBSTITUTE(RC[-3],""-"",""||"",LEN(RC[-3])-LEN(SUBSTITUTE(RC[-3],""-"",""""))),FIND(""||"",SUBSTITUTE(RC[-3],""-"",""||"",LEN(RC[-3])-LEN(SUBSTITUTE(RC[-3],""-"",""""))))-1),""0000000"")&""-FL.""&SUBSTITUTE(TEXT(TEXTJOIN("""",TRUE,IFERROR(MID(RC[-2],ROW(INDIRECT(""1:100"")),1)+0,"""")),""000""),TEXTJOIN("""",TRUE,IFERROR(MID(RC[-2],ROW(INDIRECT(""1:100"")" & _
        "),1)+0,"""")),RC[-2])&""/""&RIGHT(SUBSTITUTE(RC[-3],""-"",""||"",LEN(RC[-3])-LEN(SUBSTITUTE(RC[-3],""-"",""""))),LEN(SUBSTITUTE(RC[-3],""-"",""||"",LEN(RC[-3])-LEN(SUBSTITUTE(RC[-3],""-"",""""))))-FIND(""||"",SUBSTITUTE(RC[-3],""-"",""||"",LEN(RC[-3])-LEN(SUBSTITUTE(RC[-3],""-"",""""))))-1)" & _
        ""
    Range("F1:G1").Select
    Selection.Copy
    Range("A1").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, 5).Range("A1").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False


Windows(BaseVBA_SAP).Activate

    Sheets("zeq_estru_geral").Range("Tab_zeq_estru_geral[DESENHO DA LISTA DE CONSTRUÇÃO]").NumberFormat = "General"
    Sheets("zeq_estru_geral").Range("Tab_zeq_estru_geral[DESENHO DA LISTA DE CONSTRUÇÃO]").FormulaR1C1 = "=INDEX([" & WBTemp & "]Planilha1!C6,MATCH([@[NÚMERO DE OPERAÇÃO]],[" & WBTemp & "]Planilha1!C1,0))"
    Sheets("zeq_estru_geral").Range("Tab_zeq_estru_geral[DESENHO DA LISTA DE CONSTRUÇÃO]").NumberFormat = "@"
    Sheets("zeq_estru_geral").Range("Tab_zeq_estru_geral[DESENHO DA LISTA DE CONSTRUÇÃO]").Value = Sheets("zeq_estru_geral").Range("Tab_zeq_estru_geral[DESENHO DA LISTA DE CONSTRUÇÃO]").Value

    Sheets("zeq_estru_geral").Range("Tab_zeq_estru_geral[DESENHO DO PERFIL E PLANTA]").NumberFormat = "General"
    Sheets("zeq_estru_geral").Range("Tab_zeq_estru_geral[DESENHO DO PERFIL E PLANTA]").FormulaR1C1 = "=INDEX([" & WBTemp & "]Planilha1!C7,MATCH([@[NÚMERO DE OPERAÇÃO]],[" & WBTemp & "]Planilha1!C1,0))"
    Sheets("zeq_estru_geral").Range("Tab_zeq_estru_geral[DESENHO DO PERFIL E PLANTA]").NumberFormat = "@"
    Sheets("zeq_estru_geral").Range("Tab_zeq_estru_geral[DESENHO DO PERFIL E PLANTA]").Value = Sheets("zeq_estru_geral").Range("Tab_zeq_estru_geral[DESENHO DO PERFIL E PLANTA]").Value

Windows(WBTemp).Close (savechanges = True)

        Range("Tab_zeq_estru_geral[DESENHO DA LISTA DE CONSTRUÇÃO], Tab_zeq_estru_geral[DESENHO DO PERFIL E PLANTA]").Locked = True


    Sheets("zeq_estru_autop&estai").Range("Tab_zeq_estru_autop_estai[EXTENSÃO (m)]").FormulaR1C1 = "=INDEX('" & LC_NomeLC & "'!ListadeConstrucao[ExtTorre],MATCH([@[TORRE]],'" & LC_NomeLC & "'!ListadeConstrucao[NumOper],0))"
    Sheets("zeq_estru_autop&estai").Range("Tab_zeq_estru_autop_estai[EXTENSÃO (m)]").Value = Sheets("zeq_estru_autop&estai").Range("Tab_zeq_estru_autop_estai[EXTENSÃO (m)]").Value

    Sheets("zeq_estru_autop&estai").Range("Tab_zeq_estru_autop_estai[ALTURA PERNA A (m)]").FormulaR1C1 = "=INDEX('" & LC_NomeLC & "'!ListadeConstrucao[PernaA],MATCH([@[TORRE]],'" & LC_NomeLC & "'!ListadeConstrucao[NumOper],0))"
    Sheets("zeq_estru_autop&estai").Range("Tab_zeq_estru_autop_estai[ALTURA PERNA A (m)]").Value = Sheets("zeq_estru_autop&estai").Range("Tab_zeq_estru_autop_estai[ALTURA PERNA A (m)]").Value

    Sheets("zeq_estru_autop&estai").Range("Tab_zeq_estru_autop_estai[ALTURA PERNA B (m)]").FormulaR1C1 = "=INDEX('" & LC_NomeLC & "'!ListadeConstrucao[PernaB],MATCH([@[TORRE]],'" & LC_NomeLC & "'!ListadeConstrucao[NumOper],0))"
    Sheets("zeq_estru_autop&estai").Range("Tab_zeq_estru_autop_estai[ALTURA PERNA B (m)]").Value = Sheets("zeq_estru_autop&estai").Range("Tab_zeq_estru_autop_estai[ALTURA PERNA B (m)]").Value

    Sheets("zeq_estru_autop&estai").Range("Tab_zeq_estru_autop_estai[ALTURA PERNA C (m)]").FormulaR1C1 = "=INDEX('" & LC_NomeLC & "'!ListadeConstrucao[PernaC],MATCH([@[TORRE]],'" & LC_NomeLC & "'!ListadeConstrucao[NumOper],0))"
    Sheets("zeq_estru_autop&estai").Range("Tab_zeq_estru_autop_estai[ALTURA PERNA C (m)]").Value = Sheets("zeq_estru_autop&estai").Range("Tab_zeq_estru_autop_estai[ALTURA PERNA C (m)]").Value

    Sheets("zeq_estru_autop&estai").Range("Tab_zeq_estru_autop_estai[ALTURA PERNA D (m)]").FormulaR1C1 = "=INDEX('" & LC_NomeLC & "'!ListadeConstrucao[PernaD],MATCH([@[TORRE]],'" & LC_NomeLC & "'!ListadeConstrucao[NumOper],0))"
    Sheets("zeq_estru_autop&estai").Range("Tab_zeq_estru_autop_estai[ALTURA PERNA D (m)]").Value = Sheets("zeq_estru_autop&estai").Range("Tab_zeq_estru_autop_estai[ALTURA PERNA D (m)]").Value

    Sheets("zeq_estru_autop&estai").Range("Tab_zeq_estru_autop_estai[PERNA DE REFERÊNCIA]").FormulaR1C1 = "=INDEX('" & LC_NomeLC & "'!ListadeConstrucao[PernaRef],MATCH([@[TORRE]],'" & LC_NomeLC & "'!ListadeConstrucao[NumOper],0))"
    Sheets("zeq_estru_autop&estai").Range("Tab_zeq_estru_autop_estai[PERNA DE REFERÊNCIA]").Value = Sheets("zeq_estru_autop&estai").Range("Tab_zeq_estru_autop_estai[PERNA DE REFERÊNCIA]").Value

Sheets("zeq_estru_autop&estai").Activate
    Sheets("zeq_estru_autop&estai").Range("Tab_zeq_estru_autop_estai[DELTA H (m)]").FormulaR1C1 = "=INDEX('" & LC_NomeLC & "'!ListadeConstrucao[ElevacaoPernaRef],MATCH([@[TORRE]],'" & LC_NomeLC & "'!ListadeConstrucao[NumOper],0))"
    Sheets("zeq_estru_autop&estai").Range("Tab_zeq_estru_autop_estai[DELTA H (m)]").Value = Sheets("zeq_estru_autop&estai").Range("Tab_zeq_estru_autop_estai[DELTA H (m)]").Value
    Range("Tab_zeq_estru_autop_estai[DELTA H (m)]").Value = Evaluate("IFERROR(" & Range("Tab_zeq_estru_autop_estai[DELTA H (m)]").Address & "-100, ""-"")")
Sheets("Menu").Activate
    
    Range("Tab_zeq_estru_autop_estai[EXTENSÃO (m)], Tab_zeq_estru_autop_estai[ALTURA PERNA A (m)], Tab_zeq_estru_autop_estai[ALTURA PERNA B (m)], Tab_zeq_estru_autop_estai[ALTURA PERNA C (m)], Tab_zeq_estru_autop_estai[ALTURA PERNA D (m)]").Locked = True
    Range("Tab_zeq_estru_autop_estai[PERNA DE REFERÊNCIA], Tab_zeq_estru_autop_estai[DELTA H (m)]").Locked = True


'//TEMPORÁRIO (INÍCIO)**************************************************************************
    Range("Tab_zeq_estru_autop_estai[[EXTENSÃO MASTRO A (m)]:[TRAÇÃO ESTAI (kgf)]]").Value = "-"
    Range("Tab_zeq_estru_autop_estai[[EXTENSÃO MASTRO A (m)]:[TRAÇÃO ESTAI (kgf)]]").Locked = True
'//TEMPORÁRIO (FIM)*****************************************************************************

    Sheets("zeq_condutor").Range("Tab_zeq_condutor[QUANTIDADE AMORTECEDORES]").FormulaR1C1 = "=IFERROR(ROUND(INDEX('" & LC_NomeLC & "'!ListadeConstrucao[QtdAmort],MATCH(INDEX(BASE_BD_VaosLT[torre_numero_torre_1]," & _
            "MATCH([@VÃO],BASE_BD_VaosLT[identificacao_vao],0)),'" & LC_NomeLC & "'!ListadeConstrucao[NumOper],0))/COUNTIF([VÃO],[@VÃO]),0),""-"")"
    Sheets("zeq_condutor").Range("Tab_zeq_condutor[QUANTIDADE AMORTECEDORES]").Value = Sheets("zeq_condutor").Range("Tab_zeq_condutor[QUANTIDADE AMORTECEDORES]").Value
    Range("Tab_zeq_condutor[QUANTIDADE AMORTECEDORES]").Replace 0, ""
        
        Dim CellsErros As Range
        Set CellsErros = Sheets("zeq_condutor").Range("Tab_zeq_condutor[QUANTIDADE AMORTECEDORES]") '.Replace("-", "TEMPORARIO")
        On Error Resume Next
        Set CellsErros1 = CellsErros.SpecialCells(xlCellTypeConstants, xlErrors)
        CellsErros1.Value = ""
        On Error GoTo 0
        
        For Each cel In Sheets("zeq_condutor").Range("Tab_zeq_condutor[QUANTIDADE AMORTECEDORES]")
            If cel.Value <> "" Then
                cel.Locked = True
            Else:
                cel.Interior.Color = 65535
            End If
        Next cel


    Sheets("zeq_condutor").Range("Tab_zeq_condutor[TIPO AMORTECEDOR]").FormulaR1C1 = "=IF([@[QUANTIDADE AMORTECEDORES]]=""-"",""-"","""")"
    Sheets("zeq_condutor").Range("Tab_zeq_condutor[TIPO AMORTECEDOR]").Value = Sheets("zeq_condutor").Range("Tab_zeq_condutor[TIPO AMORTECEDOR]").Value
        
        Set CellsErros = Sheets("zeq_condutor").Range("Tab_zeq_condutor[TIPO AMORTECEDOR]")
        On Error Resume Next
        Set CellsErros1 = CellsErros.SpecialCells(xlCellTypeConstants, xlErrors)
        CellsErros1.Value = ""
        On Error GoTo 0
        
        For Each cel In Sheets("zeq_condutor").Range("Tab_zeq_condutor[TIPO AMORTECEDOR]")
            If cel.Value <> "" Then
                cel.Locked = True
            End If
        Next cel


    Sheets("zeq_condutor").Range("Tab_zeq_condutor[QUANTIDADE DE ESPAÇADORES]").FormulaR1C1 = "=IFERROR(ROUND(INDEX('" & LC_NomeLC & "'!ListadeConstrucao[QtdEA],MATCH(INDEX(BASE_BD_VaosLT[torre_numero_torre_1]," & _
            "MATCH([@VÃO],BASE_BD_VaosLT[identificacao_vao],0)),'" & LC_NomeLC & "'!ListadeConstrucao[NumOper],0))/COUNTIF([VÃO],[@VÃO]),0),""-"")"
    Sheets("zeq_condutor").Range("Tab_zeq_condutor[QUANTIDADE DE ESPAÇADORES]").Value = Sheets("zeq_condutor").Range("Tab_zeq_condutor[QUANTIDADE DE ESPAÇADORES]").Value
    Range("Tab_zeq_condutor[QUANTIDADE DE ESPAÇADORES]").Replace 0, ""
        
        Set CellsErros = Sheets("zeq_condutor").Range("Tab_zeq_condutor[QUANTIDADE DE ESPAÇADORES]")
        On Error Resume Next
        Set CellsErros1 = CellsErros.SpecialCells(xlCellTypeConstants, xlErrors)
        CellsErros1.Value = ""
        On Error GoTo 0
        
        For Each cel In Sheets("zeq_condutor").Range("Tab_zeq_condutor[QUANTIDADE DE ESPAÇADORES]")
            If cel.Value <> "" Then
                cel.Locked = True
            Else:
                cel.Interior.Color = 65535
            End If
        Next cel


    Sheets("zeq_condutor").Range("Tab_zeq_condutor[TIPO ESPACADOR]").FormulaR1C1 = "=IF([@[QUANTIDADE DE ESPAÇADORES]]=""-"",""-"","""")"
    Sheets("zeq_condutor").Range("Tab_zeq_condutor[TIPO ESPACADOR]").Value = Sheets("zeq_condutor").Range("Tab_zeq_condutor[TIPO ESPACADOR]").Value
        
        Set CellsErros = Sheets("zeq_condutor").Range("Tab_zeq_condutor[TIPO ESPACADOR]")
        On Error Resume Next
        Set CellsErros1 = CellsErros.SpecialCells(xlCellTypeConstants, xlErrors)
        CellsErros1.Value = ""
        On Error GoTo 0
        
        For Each cel In Sheets("zeq_condutor").Range("Tab_zeq_condutor[TIPO ESPACADOR]")
            If cel.Value <> "" Then
                cel.Locked = True
            End If
        Next cel

        Sheets("zeq_condutor").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=False, Password:=StrConv(Base64Decode("UGFkaWxoYUgyTSo="), vbUnicode)


    Sheets("zeq_pararaio").Range("Tab_zeq_pararaio[TIPO DE ARRANJO DO CABO]").FormulaR1C1 = _
        "=SWITCH(INDEX(IF([@LADO]=""Esquerdo/Central"",INDIRECT(""'" & LC_NomeLC & "'!ListadeConstrucao[SuspAncPR1]""),IF([@LADO]=""Direito"",INDIRECT(""'" & LC_NomeLC & "'!ListadeConstrucao[SuspAncPR2]""),IF([@LADO]=""Indefinido""," & _
        "INDIRECT(""'" & LC_NomeLC & "'!ListadeConstrucao[SuspAncPR3]"")))),MATCH(INDEX(BASE_BD_VaosLT[torre_numero_torre_1],MATCH([@VÃO],BASE_BD_VaosLT[identificacao_vao],0)),'" & LC_NomeLC & "'!ListadeConstrucao[NumOper],0)),""A"",""Ancoragem"",""S"",""Suspensão"",""-"",""-"","""")"
    Sheets("zeq_pararaio").Range("Tab_zeq_pararaio[TIPO DE ARRANJO DO CABO]").Value = Sheets("zeq_pararaio").Range("Tab_zeq_pararaio[TIPO DE ARRANJO DO CABO]").Value
        
        Set CellsErros = Sheets("zeq_pararaio").Range("Tab_zeq_pararaio[TIPO DE ARRANJO DO CABO]")
        On Error Resume Next
        Set CellsErros1 = CellsErros.SpecialCells(xlCellTypeConstants, xlErrors)
        CellsErros1.Value = ""
        On Error GoTo 0
        
        For Each cel In Sheets("zeq_pararaio").Range("Tab_zeq_pararaio[TIPO DE ARRANJO DO CABO]")
            If cel.Value <> "" Then
                cel.Locked = True
            Else:
                cel.Interior.Color = 65535
            End If
        Next cel


    Sheets("zeq_pararaio").Range("Tab_zeq_pararaio[QUANTIDADE AMORTECEDORES]").FormulaR1C1 = "=INDEX(SWITCH([@LADO],""Esquerdo/Central"",INDIRECT(""'" & LC_NomeLC & "'!ListadeConstrucao[QtdAmortecedorPR1]""),""Direito"",INDIRECT(""'" & LC_NomeLC & "'!ListadeConstrucao[QtdAmortecedorPR2]""),""Indefinido""," & _
        "INDIRECT(""'" & LC_NomeLC & "'!ListadeConstrucao[QtdAmortecedorPR3]"")),MATCH(INDEX(BASE_BD_VaosLT[torre_numero_torre_1],MATCH([@VÃO],BASE_BD_VaosLT[identificacao_vao],0)),'" & LC_NomeLC & "'!ListadeConstrucao[NumOper],0))"
    Sheets("zeq_pararaio").Range("Tab_zeq_pararaio[QUANTIDADE AMORTECEDORES]").Value = Sheets("zeq_pararaio").Range("Tab_zeq_pararaio[QUANTIDADE AMORTECEDORES]").Value
    Range("Tab_zeq_pararaio[QUANTIDADE AMORTECEDORES]").Replace 0, ""
        Set CellsErros = Sheets("zeq_pararaio").Range("Tab_zeq_pararaio[QUANTIDADE AMORTECEDORES]")
        On Error Resume Next
        Set CellsErros1 = CellsErros.SpecialCells(xlCellTypeConstants, xlErrors)
        CellsErros1.Value = ""
        On Error GoTo 0

        For Each cel In Sheets("zeq_pararaio").Range("Tab_zeq_pararaio[QUANTIDADE AMORTECEDORES]")
            
            If cel.Value <> "" Then
                cel.Locked = True
            Else:
                cel.Interior.Color = 65535
            End If
        Next cel


    Sheets("zeq_pararaio").Range("Tab_zeq_pararaio[TIPO AMORTECEDOR]").FormulaR1C1 = "=IF([@[QUANTIDADE AMORTECEDORES]]=""-"",""-"","""")"
    Sheets("zeq_pararaio").Range("Tab_zeq_pararaio[TIPO AMORTECEDOR]").Value = Sheets("zeq_pararaio").Range("Tab_zeq_pararaio[TIPO AMORTECEDOR]").Value
        For Each cel In Sheets("zeq_pararaio").Range("Tab_zeq_pararaio[TIPO AMORTECEDOR]")
            If cel.Value <> "" Then
                cel.Locked = True
            End If
        Next cel


    Sheets("zeq_pararaio").Range("Tab_zeq_pararaio[QUANTIDADE DE ESFERAS]").FormulaR1C1 = "=INDEX(SWITCH([@LADO],""Esquerdo/Central"",INDIRECT(""'" & LC_NomeLC & "'!ListadeConstrucao[EsferasPR1]""),""Direito"",INDIRECT(""'" & LC_NomeLC & "'!ListadeConstrucao[EsferaPR2]""),""Indefinido""," & _
        "INDIRECT(""'" & LC_NomeLC & "'!ListadeConstrucao[EsferaPR3]"")),MATCH(INDEX(BASE_BD_VaosLT[torre_numero_torre_1],MATCH([@VÃO],BASE_BD_VaosLT[identificacao_vao],0)),'" & LC_NomeLC & "'!ListadeConstrucao[NumOper],0))"
    Sheets("zeq_pararaio").Range("Tab_zeq_pararaio[QUANTIDADE DE ESFERAS]").Value = Sheets("zeq_pararaio").Range("Tab_zeq_pararaio[QUANTIDADE DE ESFERAS]").Value
    Range("Tab_zeq_pararaio[QUANTIDADE DE ESFERAS]").Replace 0, ""
        Set CellsErros = Sheets("zeq_pararaio").Range("Tab_zeq_pararaio[QUANTIDADE DE ESFERAS]")
        On Error Resume Next
        Set CellsErros1 = CellsErros.SpecialCells(xlCellTypeConstants, xlErrors)
        CellsErros1.Value = ""
        On Error GoTo 0

        For Each cel In Sheets("zeq_pararaio").Range("Tab_zeq_pararaio[QUANTIDADE DE ESFERAS]")
            If cel.Value <> "" Then
                cel.Locked = True
            Else:
                cel.Interior.Color = 65535
            End If
        Next cel


    Sheets("zeq_pararaio").Range("Tab_zeq_pararaio[ENGATE DA ESFERA]").FormulaR1C1 = "=IF([@[QUANTIDADE DE ESFERAS]]=""-"",""-"","""")"
    Sheets("zeq_pararaio").Range("Tab_zeq_pararaio[ENGATE DA ESFERA]").Value = Sheets("zeq_pararaio").Range("Tab_zeq_pararaio[ENGATE DA ESFERA]").Value
        For Each cel In Sheets("zeq_pararaio").Range("Tab_zeq_pararaio[ENGATE DA ESFERA]")
            If cel.Value <> "" Then
                cel.Locked = True
            End If
        Next cel


    Sheets("zeq_servidao").Range("Tab_zeq_servidao[LARGURA LADO ESQUERDO (m)]").FormulaR1C1 = "=INDEX('" & LC_NomeLC & "'!ListadeConstrucao[FaixaServ1]," & _
        "MATCH(INDEX(BASE_BD_VaosLT[torre_numero_torre_1],MATCH([@VÃO],BASE_BD_VaosLT[identificacao_vao],0)),'" & LC_NomeLC & "'!ListadeConstrucao[NumOper],0))"
    Sheets("zeq_servidao").Range("Tab_zeq_servidao[LARGURA LADO ESQUERDO (m)]").Value = Sheets("zeq_servidao").Range("Tab_zeq_servidao[LARGURA LADO ESQUERDO (m)]").Value

    Sheets("zeq_servidao").Range("Tab_zeq_servidao[LARGURA LADO DIREITO (m)]").FormulaR1C1 = "=INDEX('" & LC_NomeLC & "'!ListadeConstrucao[FaixaServ2]," & _
        "MATCH(INDEX(BASE_BD_VaosLT[torre_numero_torre_1],MATCH([@VÃO],BASE_BD_VaosLT[identificacao_vao],0)),'" & LC_NomeLC & "'!ListadeConstrucao[NumOper],0))"
    Sheets("zeq_servidao").Range("Tab_zeq_servidao[LARGURA LADO DIREITO (m)]").Value = Sheets("zeq_servidao").Range("Tab_zeq_servidao[LARGURA LADO DIREITO (m)]").Value

        Range("Tab_zeq_servidao[LARGURA LADO ESQUERDO (m)], Tab_zeq_servidao[LARGURA LADO DIREITO (m)]").Locked = True




'**********IMPORTANDO INFORMAÇÕES DA BASE AUXILIAR**********


Application.DisplayAlerts = False
Workbooks.Open BaseAux_Caminho, UpdateLinks:=False
Application.DisplayAlerts = True


Workbooks(BaseVBA_SAP).Activate
    Sheets("zeq_estru_geral").Activate
    
        Range("Tab_zeq_estru_geral[ALTURA MISULA (m)]").FormulaR1C1 = _
            "=IFERROR(IF([@SILHUETA]=""-"",""-"",VLOOKUP([@SILHUETA],'" & BaseAux_Nome & "'!TabTorres,8,0)+VLOOKUP([@[NÚMERO DE OPERAÇÃO]],Tab_zeq_estru_autop_estai,2,0)+VLOOKUP([@[NÚMERO DE OPERAÇÃO]],Tab_zeq_estru_autop_estai,IF(VLOOKUP([@[NÚMERO DE OPERAÇÃO]],Tab_zeq_estru_autop_estai,7,0)=""A"",3,IF(VLOOKUP([@[NÚMERO DE OPERAÇÃO]],Tab_zeq_estru_autop_estai" & _
            ",7,0)=""B"",4,IF(VLOOKUP([@[NÚMERO DE OPERAÇÃO]],Tab_zeq_estru_autop_estai,7,0)=""C"",5,IF(VLOOKUP([@[NÚMERO DE OPERAÇÃO]],Tab_zeq_estru_autop_estai,7,0)=""D"",6,0)))),0)),"""")"
        Range("Tab_zeq_estru_geral[ALTURA MISULA (m)]").Value = Range("Tab_zeq_estru_geral[ALTURA MISULA (m)]").Value
        
        Range("Tab_zeq_estru_geral[ALTURA TOTAL (m)]").FormulaR1C1 = _
            "=IFERROR(IF([@SILHUETA]=""-"",""-"",[@[ALTURA MISULA (m)]]+VLOOKUP([@SILHUETA],'" & BaseAux_Nome & "'!TabTorres,7,0)),"""")"
        Range("Tab_zeq_estru_geral[ALTURA TOTAL (m)]").Value = Range("Tab_zeq_estru_geral[ALTURA TOTAL (m)]").Value

        Range("Tab_zeq_estru_geral[DISPOSIÇÃO DAS FASES]").FormulaR1C1 = _
            "=IFERROR(IF([@SILHUETA]=""-"",""-"",VLOOKUP([@SILHUETA],'" & BaseAux_Nome & "'!TabTorres,3,0)),"""")"
        Range("Tab_zeq_estru_geral[DISPOSIÇÃO DAS FASES]").Value = Range("Tab_zeq_estru_geral[DISPOSIÇÃO DAS FASES]").Value

        Range("Tab_zeq_estru_geral[VÃO DE PESO (m)]").FormulaR1C1 = _
            "=IF([@SILHUETA]=""-"",""-"",[@[VÃO DE VENTO (m)]]-(IFERROR((VLOOKUP(INDEX(BASE_BD_VaosLT[NomeCabo],MATCH(OFFSET([@[NÚMERO DE OPERAÇÃO]],-1,0),BASE_BD_VaosLT[torre_numero_torre_1],0)),BASE_CabosWithOPGW,5,0))*(((IFERROR(VALUE(OFFSET([@[ALTURA MISULA (m)]],-1,0)),0)+IFERROR(VALUE(OFFSET([@ALTITUDE],-1,0)),0))-(IFERROR(VALUE([@[ALTURA MISULA (m)]]),0)+IFERROR(VALUE([@ALTITUDE]),0)))/(OFFSET([@[C" & _
            "OMPRIMENTO DO VÃO (m)]],-1,0))),0)+IFERROR((VLOOKUP(INDEX(BASE_BD_VaosLT[NomeCabo],MATCH(OFFSET([@[NÚMERO DE OPERAÇÃO]],1,0),BASE_BD_VaosLT[torre_numero_torre_1],0)),BASE_CabosWithOPGW,5,0))*(((IFERROR(VALUE(OFFSET([@[ALTURA MISULA (m)]],1,0)),0)+IFERROR(VALUE(OFFSET([@ALTITUDE],1,0)),0))-(IFERROR(VALUE([@[ALTURA MISULA (m)]]),0)+IFERROR(VALUE([@ALTITUDE]),0)))/([@[" & _
            "COMPRIMENTO DO VÃO (m)]])),0)))"
        Range("Tab_zeq_estru_geral[VÃO DE PESO (m)]").Value = Range("Tab_zeq_estru_geral[VÃO DE PESO (m)]").Value

        Range("Tab_zeq_estru_geral[ALTURA MISULA (m)], Tab_zeq_estru_geral[ALTURA TOTAL (m)], Tab_zeq_estru_geral[DISPOSIÇÃO DAS FASES], Tab_zeq_estru_geral[VÃO DE PESO (m)]").Locked = True
        
        Sheets("zeq_estru_geral").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=False, Password:=StrConv(Base64Decode("UGFkaWxoYUgyTSo="), vbUnicode)



Workbooks(BaseVBA_SAP).Activate
    Sheets("zeq_estru_autop&estai").Activate

        Range("Tab_zeq_estru_autop_estai[DESENHO FUNDAÇÃO PÉ]").NumberFormat = "General"
        Range("Tab_zeq_estru_autop_estai[DESENHO FUNDAÇÃO PÉ]").FormulaR1C1 = _
            "=VLOOKUP(LEFT(VLOOKUP([@TORRE],'" & LC_NomeLC & "'!ListadeConstrucao[[NumOper]:[FundPernas]],29,0),IFERROR(FIND(""/"",VLOOKUP([@TORRE],'" & LC_NomeLC & "'!ListadeConstrucao[[NumOper]:[FundPernas]],29,0))-1, " & _
            "LEN(VLOOKUP([@TORRE], '" & LC_NomeLC & "'!ListadeConstrucao[[NumOper]:[FundPernas]],29,0)))),'" & BaseAux_Nome & "'!TabFunPernas,2,0)"
        Range("Tab_zeq_estru_autop_estai[DESENHO FUNDAÇÃO PÉ]").NumberFormat = "@"
        Range("Tab_zeq_estru_autop_estai[DESENHO FUNDAÇÃO PÉ]").Value = Range("Tab_zeq_estru_autop_estai[DESENHO FUNDAÇÃO PÉ]").Value

        Range("Tab_zeq_estru_autop_estai[DESENHO FUNDAÇÃO PÉ]").Locked = True
        
        Sheets("zeq_estru_autop&estai").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=False, Password:=StrConv(Base64Decode("UGFkaWxoYUgyTSo="), vbUnicode)



Workbooks(BaseVBA_SAP).Activate
    Sheets("zeq_cadeia_isol").Activate
        Range("Tab_zeq_estru_autop_estai[TORRE]").Copy

    Workbooks.Add
    PlanTemp = Application.ActiveWorkbook.Name

        Range("A1").PasteSpecial xlPasteValues
        Application.CutCopyMode = False
        
        Range(Range("A1"), Selection.End(xlDown)).Select
        ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlNo).Name = "TabInfoBaseAux"

        Range("B1").FormulaR1C1 = "Coluna2"
        Range("C1").FormulaR1C1 = "Coluna3"
        Range("D1").FormulaR1C1 = "Coluna4"
        Range("E1").FormulaR1C1 = "Coluna5"
        Range("F1").FormulaR1C1 = "Coluna6"
        Range("G1").FormulaR1C1 = "Coluna7"
        
        Range("TabInfoBaseAux[Coluna2]").FormulaR1C1 = _
            "=VLOOKUP([@Coluna1],'" & LC_NomeLC & "'!ListadeConstrucao[[NumOper]:[TipoArranjoCondutor2]],40,0)"
        Range("TabInfoBaseAux[Coluna3]").FormulaR1C1 = _
            "=VLOOKUP([@Coluna1],'" & LC_NomeLC & "'!ListadeConstrucao[[NumOper]:[TipoArranjoCondutor2]],43,0)"
        Range("TabInfoBaseAux[Coluna4]").FormulaR1C1 = _
            "=IF(AND([@Coluna2]<>"""",[@Coluna2]<>""-""),[@Coluna2],[@Coluna3])"
        Range("TabInfoBaseAux[Coluna5]").FormulaR1C1 = _
            "=IFERROR(LEFT([@Coluna2],FIND(""/"",[@Coluna2])-1),[@Coluna2])"
        Range("TabInfoBaseAux[Coluna6]").FormulaR1C1 = _
            "=IFERROR(LEFT(SUBSTITUTE([@Coluna2],[@Coluna5]&""/"",""""),FIND(""/"",SUBSTITUTE([@Coluna2],[@Coluna5]&""/"",""""))-1),SUBSTITUTE([@Coluna2],[@Coluna5]&""/"",""""))"
        Range("TabInfoBaseAux[Coluna7]").FormulaR1C1 = _
            "=IFERROR(RIGHT(SUBSTITUTE([@Coluna2],[@Coluna5]&""/"",""""),LEN(SUBSTITUTE([@Coluna2],[@Coluna5]&""/"",""""))-FIND(""/"",SUBSTITUTE([@Coluna2],[@Coluna5]&""/"",""""))),[@Coluna5])"

Workbooks(BaseVBA_SAP).Activate
    Sheets("zeq_cadeia_isol").Activate

        Range("Tab_zeq_cadeia_isol[DESENHO DO ARRANJO]").NumberFormat = "General"
        Range("Tab_zeq_cadeia_isol[DESENHO DO ARRANJO]").FormulaR1C1 = _
            "=IF([@[FASEAMENTO ELÉTRICO]]=""A"",VLOOKUP(VLOOKUP([@TORRE]," & PlanTemp & "!TabInfoBaseAux,5,0),'" & BaseAux_Nome & "'!TabArranjoCondutor,2,0),IF([@[FASEAMENTO ELÉTRICO]]=""B"",VLOOKUP(VLOOKUP([@TORRE]," & PlanTemp & "!TabInfoBaseAux,6,0),'" & BaseAux_Nome & "'!TabArranjoCondutor,2,0),IF([@[FASEAMENTO ELÉTRICO]]=""C" & _
            """,VLOOKUP(VLOOKUP([@TORRE]," & PlanTemp & "!TabInfoBaseAux,7,0),'" & BaseAux_Nome & "'!TabArranjoCondutor,2,0))))"
        Range("Tab_zeq_cadeia_isol[DESENHO DO ARRANJO]").NumberFormat = "@"
        Range("Tab_zeq_cadeia_isol[DESENHO DO ARRANJO]").Value = Range("Tab_zeq_cadeia_isol[DESENHO DO ARRANJO]").Value
        
        Range("Tab_zeq_cadeia_isol[DESENHO DO ISOLADOR]").NumberFormat = "General"
        Range("Tab_zeq_cadeia_isol[DESENHO DO ISOLADOR]").FormulaR1C1 = _
            "=IF([@[FASEAMENTO ELÉTRICO]]=""A"",VLOOKUP(VLOOKUP([@TORRE]," & PlanTemp & "!TabInfoBaseAux,5,0),'" & BaseAux_Nome & "'!TabArranjoCondutor,3,0),IF([@[FASEAMENTO ELÉTRICO]]=""B"",VLOOKUP(VLOOKUP([@TORRE]," & PlanTemp & "!TabInfoBaseAux,6,0),'" & BaseAux_Nome & "'!TabArranjoCondutor,3,0),IF([@[FASEAMENTO ELÉTRICO]]=""C" & _
            """,VLOOKUP(VLOOKUP([@TORRE]," & PlanTemp & "!TabInfoBaseAux,7,0),'" & BaseAux_Nome & "'!TabArranjoCondutor,3,0))))"
        Range("Tab_zeq_cadeia_isol[DESENHO DO ISOLADOR]").NumberFormat = "@"
        Range("Tab_zeq_cadeia_isol[DESENHO DO ISOLADOR]").Value = Range("Tab_zeq_cadeia_isol[DESENHO DO ISOLADOR]").Value
            Selection.Replace What:="0", Replacement:=vbNullString, LookAt:=xlWhole 'v1.4

        
        Range("Tab_zeq_cadeia_isol[MATERIAL DO ISOLADOR]").FormulaR1C1 = _
            "=IF([@[FASEAMENTO ELÉTRICO]]=""A"",VLOOKUP(VLOOKUP([@TORRE]," & PlanTemp & "!TabInfoBaseAux,5,0),'" & BaseAux_Nome & "'!TabArranjoCondutor,4,0),IF([@[FASEAMENTO ELÉTRICO]]=""B"",VLOOKUP(VLOOKUP([@TORRE]," & PlanTemp & "!TabInfoBaseAux,6,0),'" & BaseAux_Nome & "'!TabArranjoCondutor,4,0),IF([@[FASEAMENTO ELÉTRICO]]=""C" & _
            """,VLOOKUP(VLOOKUP([@TORRE]," & PlanTemp & "!TabInfoBaseAux,7,0),'" & BaseAux_Nome & "'!TabArranjoCondutor,4,0))))"
        Range("Tab_zeq_cadeia_isol[MATERIAL DO ISOLADOR]").Value = Range("Tab_zeq_cadeia_isol[MATERIAL DO ISOLADOR]").Value
        
        Range("Tab_zeq_cadeia_isol[COMPRIMENTO DA CADEIA (m)]").FormulaR1C1 = _
            "=IF([@[FASEAMENTO ELÉTRICO]]=""A"",VLOOKUP(VLOOKUP([@TORRE]," & PlanTemp & "!TabInfoBaseAux,5,0),'" & BaseAux_Nome & "'!TabArranjoCondutor,6,0),IF([@[FASEAMENTO ELÉTRICO]]=""B"",VLOOKUP(VLOOKUP([@TORRE]," & PlanTemp & "!TabInfoBaseAux,6,0),'" & BaseAux_Nome & "'!TabArranjoCondutor,6,0),IF([@[FASEAMENTO ELÉTRICO]]=""C" & _
            """,VLOOKUP(VLOOKUP([@TORRE]," & PlanTemp & "!TabInfoBaseAux,7,0),'" & BaseAux_Nome & "'!TabArranjoCondutor,6,0))))"
        Range("Tab_zeq_cadeia_isol[COMPRIMENTO DA CADEIA (m)]").Value = Range("Tab_zeq_cadeia_isol[COMPRIMENTO DA CADEIA (m)]").Value
        
        Range("Tab_zeq_cadeia_isol[QUANTIDADE TOTAL ISOL ARRANJO]").FormulaR1C1 = _
            "=IF([@[FASEAMENTO ELÉTRICO]]=""A"",VLOOKUP(VLOOKUP([@TORRE]," & PlanTemp & "!TabInfoBaseAux,5,0),'" & BaseAux_Nome & "'!TabArranjoCondutor,5,0),IF([@[FASEAMENTO ELÉTRICO]]=""B"",VLOOKUP(VLOOKUP([@TORRE]," & PlanTemp & "!TabInfoBaseAux,6,0),'" & BaseAux_Nome & "'!TabArranjoCondutor,5,0),IF([@[FASEAMENTO ELÉTRICO]]=""C" & _
            """,VLOOKUP(VLOOKUP([@TORRE]," & PlanTemp & "!TabInfoBaseAux,7,0),'" & BaseAux_Nome & "'!TabArranjoCondutor,5,0))))"
        Range("Tab_zeq_cadeia_isol[QUANTIDADE TOTAL ISOL ARRANJO]").Value = Range("Tab_zeq_cadeia_isol[QUANTIDADE TOTAL ISOL ARRANJO]").Value
        
        Range("Tab_zeq_cadeia_isol[TIPO DE ARRANJO DA CADEIA]").FormulaR1C1 = _
            "=IF(IF([@[FASEAMENTO ELÉTRICO]]=""A"",VLOOKUP(VLOOKUP([@TORRE]," & PlanTemp & "!TabInfoBaseAux,5,0),'" & BaseAux_Nome & "'!TabArranjoCondutor,7,0),IF([@[FASEAMENTO ELÉTRICO]]=""B"",VLOOKUP(VLOOKUP([@TORRE]," & PlanTemp & "!TabInfoBaseAux,6,0),'" & BaseAux_Nome & "'!TabArranjoCondutor,7,0),IF([@[FASEAMENTO ELÉTRICO]]=" & _
            """C"",VLOOKUP(VLOOKUP([@TORRE]," & PlanTemp & "!TabInfoBaseAux,7,0),'" & BaseAux_Nome & "'!TabArranjoCondutor,7,0))))<>""Ancoragem"",IF([@[FASEAMENTO ELÉTRICO]]=""A"",VLOOKUP(VLOOKUP([@TORRE]," & PlanTemp & "!TabInfoBaseAux,5,0),'" & BaseAux_Nome & "'!TabArranjoCondutor,7,0),IF([@[FASEAMENTO ELÉTRICO]]=""B"",VLOOKUP(" & _
            "VLOOKUP([@TORRE]," & PlanTemp & "!TabInfoBaseAux,6,0),'" & BaseAux_Nome & "'!TabArranjoCondutor,7,0),IF([@[FASEAMENTO ELÉTRICO]]=""C"",VLOOKUP(VLOOKUP([@TORRE]," & PlanTemp & "!TabInfoBaseAux,7,0),'" & BaseAux_Nome & "'!TabArranjoCondutor,7,0)))),IF(AND(VLOOKUP([@TORRE],'" & _
            LC_NomeLC & "'!ListadeConstrucao[[NumOper]:[QtdArranjoJumper]],45,0)<>""-"",VLOOKUP([@TORRE],'" & LC_NomeLC & "'!ListadeConstrucao[[NumOper]:[QtdArranjoJumper]],45,0)<>""""),""Ancoragem com Cadeia de Jumper"",""Ancoragem""))"
        Range("Tab_zeq_cadeia_isol[TIPO DE ARRANJO DA CADEIA]").Value = Range("Tab_zeq_cadeia_isol[TIPO DE ARRANJO DA CADEIA]").Value
        
        Range("Tab_zeq_cadeia_isol[COMPOSIÇÃO DO ARRANJO]").FormulaR1C1 = _
            "=IF([@[FASEAMENTO ELÉTRICO]]=""A"",VLOOKUP(VLOOKUP([@TORRE]," & PlanTemp & "!TabInfoBaseAux,5,0),'" & BaseAux_Nome & "'!TabArranjoCondutor,8,0),IF([@[FASEAMENTO ELÉTRICO]]=""B"",VLOOKUP(VLOOKUP([@TORRE]," & PlanTemp & "!TabInfoBaseAux,6,0),'" & BaseAux_Nome & "'!TabArranjoCondutor,8,0),IF([@[FASEAMENTO ELÉTRICO]]=""C" & _
            """,VLOOKUP(VLOOKUP([@TORRE]," & PlanTemp & "!TabInfoBaseAux,7,0),'" & BaseAux_Nome & "'!TabArranjoCondutor,8,0))))"
        Range("Tab_zeq_cadeia_isol[COMPOSIÇÃO DO ARRANJO]").Value = Range("Tab_zeq_cadeia_isol[COMPOSIÇÃO DO ARRANJO]").Value

        Range("Tab_zeq_cadeia_isol[DESENHO DO ARRANJO], Tab_zeq_cadeia_isol[DESENHO DO ISOLADOR], Tab_zeq_cadeia_isol[MATERIAL DO ISOLADOR], Tab_zeq_cadeia_isol[COMPRIMENTO DA CADEIA (m)]").Locked = True
        Range("Tab_zeq_cadeia_isol[QUANTIDADE TOTAL ISOL ARRANJO], Tab_zeq_cadeia_isol[TIPO DE ARRANJO DA CADEIA], Tab_zeq_cadeia_isol[COMPOSIÇÃO DO ARRANJO]").Locked = True
        
        Sheets("zeq_cadeia_isol").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=False, Password:=StrConv(Base64Decode("UGFkaWxoYUgyTSo="), vbUnicode)

        Application.DisplayAlerts = False
        Workbooks(PlanTemp).Close
        Application.DisplayAlerts = True
        
        
Workbooks(BaseVBA_SAP).Activate
    Sheets("zeq_pararaio").Activate
        Range("Tab_zeq_pararaio[[VÃO]:[LADO]]").Copy

    Workbooks.Add
    PlanTemp = Application.ActiveWorkbook.Name

        Range("A1").PasteSpecial xlPasteValues
        Application.CutCopyMode = False
        
        ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlNo).Name = "TabInfoBaseAux"

        Range("C1").FormulaR1C1 = "Coluna3"
        Range("D1").FormulaR1C1 = "Coluna4"
        Range("E1").FormulaR1C1 = "Coluna5"
        Range("F1").FormulaR1C1 = "Coluna6"
        Range("G1").FormulaR1C1 = "Coluna7"
        
        Range("TabInfoBaseAux[Coluna3]").FormulaR1C1 = _
            "=INDEX('" & BaseVBA_SAP & "'!BASE_BD_VaosLT[torre_numero_torre_1],MATCH([@Coluna1],'" & BaseVBA_SAP & "'!BASE_BD_VaosLT[identificacao_vao],0))"
        
        Range("TabInfoBaseAux[Coluna4]").FormulaR1C1 = _
            "=VLOOKUP([@Coluna3],'" & LC_NomeLC & "'!ListadeConstrucao[[NumOper]:[TipoArranjoPR3]],IF([@Coluna2]=""Esquerdo/Central"",58,IF([@Coluna2]=""Direito"",67,IF([@Coluna2]=""Indefinido"",76))),0)"
        
        Range("TabInfoBaseAux[Coluna5]").FormulaR1C1 = _
            "=IFERROR(VLOOKUP(RIGHT([@Coluna4],LEN([@Coluna4])-IFERROR(FIND(""/"",[@Coluna4]),0)),'" & BaseAux_Nome & "'!TabArranjoPR,2,0),""-"")"
        
        Range("TabInfoBaseAux[Coluna6]").FormulaR1C1 = _
            "=IFERROR(VLOOKUP(RIGHT([@Coluna4],LEN([@Coluna4])-IFERROR(FIND(""/"",[@Coluna4]),0)),'" & BaseAux_Nome & "'!TabArranjoPR,4,0),""-"")"
        
        Range("TabInfoBaseAux[Coluna7]").FormulaR1C1 = _
            "=IFERROR(VLOOKUP(RIGHT([@Coluna4],LEN([@Coluna4])-IFERROR(FIND(""/"",[@Coluna4]),0)),'" & BaseAux_Nome & "'!TabArranjoPR,6,0),""-"")"

        Range("TabInfoBaseAux[Coluna5]").Copy

Workbooks(BaseVBA_SAP).Activate
    Sheets("zeq_pararaio").Activate
        Range("Tab_zeq_pararaio[DESENHO DO ARRANJO]").PasteSpecial Paste:=xlPasteValues

Workbooks(PlanTemp).Activate
        Range("TabInfoBaseAux[Coluna6]").Copy

Workbooks(BaseVBA_SAP).Activate
    Sheets("zeq_pararaio").Activate
        Range("Tab_zeq_pararaio[PARA-RAIO ISOLADOS]").PasteSpecial Paste:=xlPasteValues
        
        Range("A1").Select
        Range("Tab_zeq_pararaio[DESENHO DO ARRANJO], Tab_zeq_pararaio[PARA-RAIO ISOLADOS]").Locked = True
        
        Sheets("zeq_pararaio").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=False, Password:=StrConv(Base64Decode("UGFkaWxoYUgyTSo="), vbUnicode)


Workbooks(PlanTemp).Activate
        Range("TabInfoBaseAux[Coluna7]").Copy

Workbooks(BaseVBA_SAP).Activate
    Sheets("zeq_opgw").Activate
        Range("Tab_zeq_opgw[CAIXA DE EMENDA]").PasteSpecial Paste:=xlPasteValues

            LastRow_OPGW = Application.WorksheetFunction.CountA(Range("Tab_zeq_opgw[VÃO]"))
                For i = 1 To LastRow_OPGW
                    If Range("Tab_zeq_opgw[FABRICANTE OPGW]").Cells(i).Value <> "-" And Range("Tab_zeq_opgw[FABRICANTE OPGW]").Cells(i).Value <> "" Then
                        Range("Tab_zeq_opgw[CAIXA DE EMENDA]").Cells(i).Value = Range("Tab_zeq_opgw[CAIXA DE EMENDA]").Cells(i).Value
                    Else
                        Range("Tab_zeq_opgw[CAIXA DE EMENDA]").Cells(i).Value = "-"
                    End If
                Next i

        Range("A1").Select
        Range("Tab_zeq_opgw[CAIXA DE EMENDA]").Locked = True
        
        Sheets("zeq_opgw").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=False, Password:=StrConv(Base64Decode("UGFkaWxoYUgyTSo="), vbUnicode)
        
        Application.DisplayAlerts = False
        Workbooks(PlanTemp).Close
        Application.DisplayAlerts = True



Workbooks(BaseVBA_SAP).Activate
    Sheets("zeq_servidao").Activate

        Range("Tab_zeq_servidao[NATUREZA DA TRAVESSIA CRÍTICA]").FormulaR1C1 = _
            "=IFERROR(IFERROR(INDEX('" & BaseAux_Nome & "'!TabTravAerea[TipoTravessia],MATCH([@VÃO],'" & BaseAux_Nome & "'!TabTravAerea[Vão],0)),INDEX('" & BaseAux_Nome & "'!TabTravObs[TipoTravessia],MATCH([@VÃO],'" & BaseAux_Nome & "'!TabTravObs[Vão],0))),""-"")"
        Range("Tab_zeq_servidao[NATUREZA DA TRAVESSIA CRÍTICA]").Value = Range("Tab_zeq_servidao[NATUREZA DA TRAVESSIA CRÍTICA]").Value
        
            For Each cel In Sheets("zeq_servidao").Range("Tab_zeq_servidao[NATUREZA DA TRAVESSIA CRÍTICA]")
                If cel.Value <> "" And cel.Value <> 0 Then
                    cel.Locked = True
                Else:
                    cel.Value = ""
                    cel.Interior.Color = 65535
                End If
            Next cel

        
        Range("Tab_zeq_servidao[DIST VERTIC CABO-TRAVESSIA (m)]").FormulaR1C1 = _
            "=IFERROR(IFERROR(INDEX('" & BaseAux_Nome & "'!TabTravAerea[Dist Cabo-Travessia (m)],MATCH([@VÃO],'" & BaseAux_Nome & "'!TabTravAerea[Vão],0)),INDEX('" & BaseAux_Nome & "'!TabTravObs[Dist Cabo-Travessia (m)],MATCH([@VÃO],'" & BaseAux_Nome & "'!TabTravObs[Vão],0))),""-"")"
        Range("Tab_zeq_servidao[DIST VERTIC CABO-TRAVESSIA (m)]").Value = Range("Tab_zeq_servidao[DIST VERTIC CABO-TRAVESSIA (m)]").Value
        
        Range("Tab_zeq_servidao[OBSERVAÇÃO]").FormulaR1C1 = _
            "=IFERROR(IFERROR(INDEX('" & BaseAux_Nome & "'!TabTravAerea[Observações],MATCH([@VÃO],'" & BaseAux_Nome & "'!TabTravAerea[Vão],0)),INDEX('" & BaseAux_Nome & "'!TabTravObs[Observações],MATCH([@VÃO],'" & BaseAux_Nome & "'!TabTravObs[Vão],0))),""-"")"
        Range("Tab_zeq_servidao[OBSERVAÇÃO]").Value = Range("Tab_zeq_servidao[OBSERVAÇÃO]").Value


        Range("Tab_zeq_servidao[DIST VERTIC CABO-TRAVESSIA (m)], Tab_zeq_servidao[OBSERVAÇÃO]").Locked = True
        
        Sheets("zeq_servidao").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=False, Password:=StrConv(Base64Decode("UGFkaWxoYUgyTSo="), vbUnicode)


        Application.DisplayAlerts = False
        Workbooks(BaseAux_Nome).Close
        Application.DisplayAlerts = True

        Application.DisplayAlerts = False
        Workbooks(LC_NomeLC).Close
        Application.DisplayAlerts = True


'**********INCLUSÕES DIVERSAS**********


'Flechas em "ZLI_ParametrosOp":


    ActiveWorkbook.Connections("Consulta - Query_CondutorTipico").Refresh
    ActiveWorkbook.Connections("Consulta - Query_PRTipico").Refresh

    Range("Label_ZLI_ParametrosOp_FlechaMaxCondut").Formula = _
        "=(1/((INDEX(BASE_Tracoes[Peso (kg/km)],MATCH(Query_CondutorTipico[CaboCondutor_Tipico],BASE_Tracoes[Nome],0))/1000)/(INDEX(BASE_Tracoes[Carga de Ruptura (kgf)],MATCH(Query_CondutorTipico[CaboCondutor_Tipico],BASE_Tracoes[Nome],0))*INDEX(BASE_Tracoes[EDS (%)],MATCH(Query_CondutorTipico[CaboCondutor_Tipico],BASE_Tracoes[Nome],0)))))*(COSH(((INDEX(BASE_Tracoes[Peso (kg" & _
        "/km)],MATCH(Query_CondutorTipico[CaboCondutor_Tipico],BASE_Tracoes[Nome],0))/1000)/(INDEX(BASE_Tracoes[Carga de Ruptura (kgf)],MATCH(Query_CondutorTipico[CaboCondutor_Tipico],BASE_Tracoes[Nome],0))*INDEX(BASE_Tracoes[EDS (%)],MATCH(Query_CondutorTipico[CaboCondutor_Tipico],BASE_Tracoes[Nome],0))))*((SUM(Tab_zeq_estru_geral[COMPRIMENTO DO VÃO (m)])/INDEX(BASE_BD_Proj" & _
        "etosLT[qtde_total_vaos],1)/2)-(1/((INDEX(BASE_Tracoes[Peso (kg/km)],MATCH(Query_CondutorTipico[CaboCondutor_Tipico],BASE_Tracoes[Nome],0))/1000)/(INDEX(BASE_Tracoes[Carga de Ruptura (kgf)],MATCH(Query_CondutorTipico[CaboCondutor_Tipico],BASE_Tracoes[Nome],0))*INDEX(BASE_Tracoes[EDS (%)],MATCH(Query_CondutorTipico[CaboCondutor_Tipico],BASE_Tracoes[Nome],0)))))*ASINH(" & _
        "(((INDEX(BASE_Tracoes[Peso (kg/km)],MATCH(Query_CondutorTipico[CaboCondutor_Tipico],BASE_Tracoes[Nome],0))/1000)/(INDEX(BASE_Tracoes[Carga de Ruptura (kgf)],MATCH(Query_CondutorTipico[CaboCondutor_Tipico],BASE_Tracoes[Nome],0))*INDEX(BASE_Tracoes[EDS (%)],MATCH(Query_CondutorTipico[CaboCondutor_Tipico],BASE_Tracoes[Nome],0))))*0)/(2*SINH(((INDEX(BASE_Tracoes[Peso (k" & _
        "g/km)],MATCH(Query_CondutorTipico[CaboCondutor_Tipico],BASE_Tracoes[Nome],0))/1000)/(INDEX(BASE_Tracoes[Carga de Ruptura (kgf)],MATCH(Query_CondutorTipico[CaboCondutor_Tipico],BASE_Tracoes[Nome],0))*INDEX(BASE_Tracoes[EDS (%)],MATCH(Query_CondutorTipico[CaboCondutor_Tipico],BASE_Tracoes[Nome],0))))*SUM(Tab_zeq_estru_geral[COMPRIMENTO DO VÃO (m)])/INDEX(BASE_BD_Proje" & _
        "tosLT[qtde_total_vaos],1)/2)))))-1)"
        
    Range("Label_ZLI_ParametrosOp_FlechaMaxCondut").Value = Range("Label_ZLI_ParametrosOp_FlechaMaxCondut").Value
    
    
    Range("Label_ZLI_ParametrosOp_FlechaMaxPR").Formula = _
        "=(1/((INDEX(BASE_Tracoes[Peso (kg/km)],MATCH(Query_PRTipico[CaboPR_Tipico],BASE_Tracoes[Nome],0))/1000)/(INDEX(BASE_Tracoes[Carga de Ruptura (kgf)],MATCH(Query_PRTipico[CaboPR_Tipico],BASE_Tracoes[Nome],0))*INDEX(BASE_Tracoes[EDS (%)],MATCH(Query_PRTipico[CaboPR_Tipico],BASE_Tracoes[Nome],0)))))*(COSH(((INDEX(BASE_Tracoes[Peso (kg/km)],MATCH(Query_PRTipico[CaboPR_Ti" & _
        "pico],BASE_Tracoes[Nome],0))/1000)/(INDEX(BASE_Tracoes[Carga de Ruptura (kgf)],MATCH(Query_PRTipico[CaboPR_Tipico],BASE_Tracoes[Nome],0))*INDEX(BASE_Tracoes[EDS (%)],MATCH(Query_PRTipico[CaboPR_Tipico],BASE_Tracoes[Nome],0))))*((SUM(Tab_zeq_estru_geral[COMPRIMENTO DO VÃO (m)])/INDEX(BASE_BD_ProjetosLT[qtde_total_vaos],1)/2)-(1/((INDEX(BASE_Tracoes[Peso (kg/km)],MATC" & _
        "H(Query_PRTipico[CaboPR_Tipico],BASE_Tracoes[Nome],0))/1000)/(INDEX(BASE_Tracoes[Carga de Ruptura (kgf)],MATCH(Query_PRTipico[CaboPR_Tipico],BASE_Tracoes[Nome],0))*INDEX(BASE_Tracoes[EDS (%)],MATCH(Query_PRTipico[CaboPR_Tipico],BASE_Tracoes[Nome],0)))))*ASINH((((INDEX(BASE_Tracoes[Peso (kg/km)],MATCH(Query_PRTipico[CaboPR_Tipico],BASE_Tracoes[Nome],0))/1000)/(INDEX(" & _
        "BASE_Tracoes[Carga de Ruptura (kgf)],MATCH(Query_PRTipico[CaboPR_Tipico],BASE_Tracoes[Nome],0))*INDEX(BASE_Tracoes[EDS (%)],MATCH(Query_PRTipico[CaboPR_Tipico],BASE_Tracoes[Nome],0))))*0)/(2*SINH(((INDEX(BASE_Tracoes[Peso (kg/km)],MATCH(Query_PRTipico[CaboPR_Tipico],BASE_Tracoes[Nome],0))/1000)/(INDEX(BASE_Tracoes[Carga de Ruptura (kgf)],MATCH(Query_PRTipico[CaboPR_" & _
        "Tipico],BASE_Tracoes[Nome],0))*INDEX(BASE_Tracoes[EDS (%)],MATCH(Query_PRTipico[CaboPR_Tipico],BASE_Tracoes[Nome],0))))*SUM(Tab_zeq_estru_geral[COMPRIMENTO DO VÃO (m)])/INDEX(BASE_BD_ProjetosLT[qtde_total_vaos],1)/2)))))-1)"
    
    Range("Label_ZLI_ParametrosOp_FlechaMaxPR").Value = Range("Label_ZLI_ParametrosOp_FlechaMaxPR").Value


        For Each cel In Range("Tab_zli_parametros_op[Coluna2]")
            If cel.Value <> "" Then
                cel.Locked = True
            Else:
                cel.Interior.Color = 65535
            End If
        Next cel


Sheets("zli_parametros_OP").Activate
Range("A1").Activate
Sheets("Menu").Activate

Sheets("zli_parametros_OP").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=False, Password:=StrConv(Base64Decode("UGFkaWxoYUgyTSo="), vbUnicode)


'(FIM) Flechas em "ZLI_ParametrosOp"




'**********AJUSTE DE FORMATAÇÃO, SALVAMENTO E FINALIZAÇAÕ**********

        Sheets("Menu").Activate
        Sheets("Menu").Unprotect (StrConv(Base64Decode("UGFkaWxoYUgyTSo="), vbUnicode))

        With Range("Label_OndaLT", "Label_NomeLT")
            .Interior.Pattern = xlSolid
            .Interior.PatternColorIndex = xlAutomatic
            .Interior.ThemeColor = xlThemeColorAccent5
            .Interior.TintAndShade = 0.799981688894314
            .Interior.PatternTintAndShade = 0
            .Validation.Delete
            .Locked = True
        End With

        Sheets("Menu").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=StrConv(Base64Decode("UGFkaWxoYUgyTSo="), vbUnicode)


        Revisao = 0
        Do While Dir(DiretorioSave & "\(" & LT_CodLT & ")_v" & VBA_SAP_Versao & "_Dados SAP R." & Revisao & ".xlsm") <> ""
            Revisao = Revisao + 1
        Loop
    
    
        ActiveWorkbook.SaveAs Filename:= _
            DiretorioSave & "\(" & LT_CodLT & ")_v" & VBA_SAP_Versao & "_Dados SAP R." & Revisao & ".xlsm" _
            , FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False


Application.ScreenUpdating = True

        MsgFimPreenchimento = MsgBox("Preenchimento automático concluído com sucesso, com base nas informações disponíveis no Banco de Dados, Lista de Construção, Planilha BDIT e Base Auxiliar." _
            , vbInformation, "Preenchimento concluído com sucesso!")


End Sub


