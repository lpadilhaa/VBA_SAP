Sub LogErrors()


Select Case ActiveSheet.Name
    Case "zli_transmissao": APIAtual = "ZLI_TRANSMISSAO"
    Case "zli_parametros_OP": APIAtual = "ZLI_PARAMETROS_OP"
    Case "zeq_estru_geral": APIAtual = "ZEQ_ESTRUTURA_GERAL"
    'Case "zeq_estru_autop&estai": APIAtual = "ZEQ_ESTRUTURA_AUTOPORTANTE"
    Case "zeq_cadeia_isol": APIAtual = "ZEQ_CADEIA_ISOLADORES"
    Case "zeq_aterramento": APIAtual = "ZEQ_ATERRAMENTO"
    Case "zeq_acessos": APIAtual = "ZEQ_ACESSOS"
    Case "zeq_condutor": APIAtual = "ZEQ_CONDUTOR"
    Case "zeq_pararaio": APIAtual = "ZEQ_PARARAIO"
    Case "zeq_opgw": APIAtual = "ZEQ_OPGW"
    Case "zeq_servidao": APIAtual = "ZEQ_SERVIDAO"
End Select



If ActiveSheet.Name = "zeq_estru_autop&estai" Then
    If Application.WorksheetFunction.CountIf(Range("Tab_zeq_estru_autop_estai[PERNA DE REFERÊNCIA]"), "<>-") > 0 Then
        APIAtual = "ZEQ_ESTRUTURA_AUTOPORTANTE"
        On Error GoTo NoLog1
        RowLog = Application.WorksheetFunction.Match(APIAtual, Range("TabLogErros[API]"), 0)
            MsgLog = Range("TabLogErros[Msg]").Rows(RowLog).Value
            DataHoraEnvio = Range("TabLogErros[Fim]").Rows(RowLog).Value
            
            If Range("TabLogErros[Erro]").Rows(RowLog).Value = 0 Then
                MsgBoxStyle = vbInformation
            ElseIf Range("TabLogErros[Erro]").Rows(RowLog).Value = 1 Then
                MsgBoxStyle = vbCritical
            End If
            
            ShowMsgLog = MsgBox(MsgLog & Chr(13) & Chr(13) & "(Registro: " & DataHoraEnvio & ")", MsgBoxStyle, APIAtual)
            
            If Application.WorksheetFunction.CountIf(Range("Tab_zeq_estru_autop_estai[EXTENSÃO MASTRO A (m)]"), "<>-") = 0 Then
                GoTo Finalizar
            Else: GoTo Log_Estai
            End If
    
NoLog1:
    On Error GoTo 0
    On Error GoTo -1
                ShowMsgNull = MsgBox("Não há registro de envios de dados para a API " & APIAtual & _
                    ".", vbExclamation, APIAtual)
                If Application.WorksheetFunction.CountIf(Range("Tab_zeq_estru_autop_estai[EXTENSÃO MASTRO A (m)]"), "<>-") = 0 Then
                    Exit Sub
                Else: GoTo Log_Estai
                End If
    End If


Log_Estai:

    If Application.WorksheetFunction.CountIf(Range("Tab_zeq_estru_autop_estai[EXTENSÃO MASTRO A (m)]"), "<>-") > 0 Then
        APIAtual = "ZEQ_ESTRUTURA_ESTAIADA"
        On Error GoTo NoLog2
        RowLog = Application.WorksheetFunction.Match(APIAtual, Range("TabLogErros[API]"), 0)
            MsgLog = Range("TabLogErros[Msg]").Rows(RowLog).Value
            DataHoraEnvio = Range("TabLogErros[Fim]").Rows(RowLog).Value
            
            If Range("TabLogErros[Erro]").Rows(RowLog).Value = 0 Then
                MsgBoxStyle = vbInformation
            ElseIf Range("TabLogErros[Erro]").Rows(RowLog).Value = 1 Then
                MsgBoxStyle = vbCritical
            End If
            
            ShowMsgLog = MsgBox(MsgLog & Chr(13) & Chr(13) & "(Registro: " & DataHoraEnvio & ")", MsgBoxStyle, APIAtual)
            
                GoTo Finalizar
    
NoLog2:
    On Error GoTo 0
    On Error GoTo -1
                ShowMsgNull = MsgBox("Não há registro de envios de dados para a API " & APIAtual & _
                    ".", vbExclamation, APIAtual)
                    Exit Sub
    End If
End If



If APIAtual <> "" Then
    
    On Error GoTo NoLog
    RowLog = Application.WorksheetFunction.Match(APIAtual, Range("TabLogErros[API]"), 0)
        MsgLog = Range("TabLogErros[Msg]").Rows(RowLog).Value
        DataHoraEnvio = Range("TabLogErros[Fim]").Rows(RowLog).Value
        
        If Range("TabLogErros[Erro]").Rows(RowLog).Value = 0 Then
            MsgBoxStyle = vbInformation
        ElseIf Range("TabLogErros[Erro]").Rows(RowLog).Value = 1 Then
            MsgBoxStyle = vbCritical
        End If
        
        ShowMsgLog = MsgBox(MsgLog & Chr(13) & Chr(13) & "(Registro: " & DataHoraEnvio & ")", MsgBoxStyle, APIAtual)
        
        GoTo Finalizar

NoLog:
    On Error GoTo 0
    On Error GoTo -1
            ShowMsgNull = MsgBox("Não há registro de envios de dados para a API " & APIAtual & _
                ".", vbExclamation, APIAtual)
            
            Exit Sub

Else:

    NumeroLogs = Application.WorksheetFunction.CountA(Range("TabLogErros[API]"))
    
    If NumeroLogs = 0 Then
        ShowMsgNull = MsgBox("Não há registro de envios de dados para nenhuma das API's." _
            , vbExclamation, APIAtual)
    End If
    
    Repete = 1
    
    Do While Repete <= NumeroLogs
    
        APIAtual = Range("TabLogErros[API]").Rows(Repete).Value
            MsgLog = Range("TabLogErros[Msg]").Rows(Repete).Value
            DataHoraEnvio = Range("TabLogErros[Fim]").Rows(Repete).Value
            
            If Range("TabLogErros[Erro]").Rows(Repete).Value = 0 Then
                MsgBoxStyle = vbInformation
            ElseIf Range("TabLogErros[Erro]").Rows(Repete).Value = 1 Then
                MsgBoxStyle = vbCritical
            End If
            
            ShowMsgLog = MsgBox(MsgLog & Chr(13) & Chr(13) & "(Registro: " & DataHoraEnvio & ")", MsgBoxStyle, APIAtual)
    
    Repete = Repete + 1
    
    Loop

End If


Finalizar:

End Sub

