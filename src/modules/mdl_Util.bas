Attribute VB_Name = "mdl_Util"
Option Explicit

Sub textBoxSomenteNumerosOptionalMoeda(ByVal pKeyAscii As MSForms.ReturnInteger, Optional pMoeda As Boolean = False)

    Dim strValid As String
    
    If pMoeda Then
        'moeda permite virgula para casas decimais de centavos
        strValid = "0123456789,"
    Else
        strValid = "0123456789"
    End If
    
    If InStr(strValid, Chr(pKeyAscii)) = 0 Then
        pKeyAscii = 0
    End If

End Sub
Public Function validaTextBoxes(nomeTextbox As MSForms.textBox, Optional pOptionalCampoNumerico As Boolean = False, Optional pQtdMinimaCaracteres As Integer = 0, Optional pValorMinimo As Double = -1) As Boolean

    validaTextBoxes = True

    '************************************************************************************************************
    'pinta bordas em vermelho em campos em branco
    If nomeTextbox = "" Then
        validaTextBoxes = False
        nomeTextbox.BorderStyle = fmBorderStyleSingle ' borda
        nomeTextbox.BorderColor = &HFF&          ' Vermelho
        nomeTextbox.SetFocus
        MsgBox "CAMPO OBRIGATÓRIO EM BRANCO.", vbCritical, "C a m p o   e m    B r a n c o !"
        Exit Function
    ElseIf pOptionalCampoNumerico And nomeTextbox <= 0 Then
        validaTextBoxes = False
        nomeTextbox.BorderStyle = fmBorderStyleSingle ' borda
        nomeTextbox.BorderColor = &HFF&          ' Vermelho
        nomeTextbox.SetFocus
        MsgBox "CAMPO NUMÉRICO OU MONETÁRIO. VALOR INVÁLIDO", vbCritical, "V A L O R   I N V Á L I D O !"
        Exit Function
    ElseIf pOptionalCampoNumerico And nomeTextbox < pValorMinimo Then
        validaTextBoxes = False
        nomeTextbox.BorderStyle = fmBorderStyleSingle ' borda
        nomeTextbox.BorderColor = &HFF&          ' Vermelho
        nomeTextbox.SetFocus
        MsgBox "VALOR MINIMO PARA O CAMPO ABAIXO DO ACEITÁVEL. O VALOR DEVE SER IGUAL OU MAIOR QUE " & pValorMinimo & ".", vbCritical, "V A L O R   I N V Á L I D O !"
        Exit Function
    ElseIf Len(nomeTextbox) < pQtdMinimaCaracteres Then
        validaTextBoxes = False
        nomeTextbox.BorderStyle = fmBorderStyleSingle ' borda
        nomeTextbox.BorderColor = &HFF&          ' Vermelho
        nomeTextbox.SetFocus
        MsgBox "QUANTIDADE MINIMA DE CARACTERES/DIGITOS INVÁLIDO. O CAMPO DEVE CONTER NO MÍNIMO " & pQtdMinimaCaracteres & " CARACTERES.", vbCritical, "QUANTIDADE MINIMA DE CARACTERES INVÁLIDO!"
        Exit Function
    Else: nomeTextbox.SpecialEffect = fmSpecialEffectSunken ' Normal
    End If


End Function
'
'
'
'
'--------------------------------------------------
'esta função alimenta um array onde suas linhas podem ser
'redimensionadas/incrementadas preservando seus valores antigos
'
'
'Ex.: de redimencionamento
'ReDim Preserve arrayDados(UBound(arrayDados), linArray + 10)
'--------------------------------------------------
'
'
'
Function arrayComDadosDeTabela(codenamePlanilha As Worksheet) As Variant()

    Dim linha               As Long
    Dim coluna              As Long
    Dim ultimaLinha         As Long
    Dim ultimaColuna        As Long
    Dim linArray            As Long
    Dim colArray            As Long
    Dim arrayLocal() As Variant


    coluna = 1
    linha = 2
    ultimaLinha = codenamePlanilha.Cells(1, 1).End(xlDown).Row - 2 '-2 para excluir cabeçalho e matriz inicia com ZERO
    ultimaColuna = codenamePlanilha.Cells(1, coluna).End(xlToRight).Column - 1 'conta a partir dos cabeçalhos para evitar celulas em branco
    ReDim arrayLocal(ultimaColuna, ultimaLinha)
    
    'popula as LINHAS DO ARRAY
    For linArray = 0 To ultimaLinha
    
        'popula as COLUNAS DO ARRAY
        For colArray = 0 To ultimaColuna
        
            'popula array - percorre cada COLUNA populando E M   L I N H A S
            arrayLocal(colArray, linArray) = codenamePlanilha.Cells(linha, coluna).Value2
            
            'incrementa para a proxima coluna
            coluna = coluna + 1
        Next colArray
        
        'reseta a coluna para a proxima linha
        coluna = 1
        'incrementa a linha para reiniciar o loop e adicionar nova linha
        linha = linha + 1
    Next linArray
    
    arrayComDadosDeTabela = arrayLocal

End Function


Function ultimaLinhaEmBranco(pPlanilha As Worksheet, Optional pIncluirPrimeiraLinha As Boolean = False) As Long
    
    If pIncluirPrimeiraLinha Then
        
        If pPlanilha.Cells(1, 1) = "" Then
            ultimaLinhaEmBranco = 1
        Else
            ultimaLinhaEmBranco = pPlanilha.Cells(1, 1).End(xlDown).Row + 1
        End If
        
    ElseIf pPlanilha.Cells(2, 1) = "" Then
        ultimaLinhaEmBranco = 2
    Else
        ultimaLinhaEmBranco = pPlanilha.Cells(1, 1).End(xlDown).Row + 1
    End If

End Function


Function extraiIdDaStringNaCombobox(pItemSelecionadoCombobox As String) As Long

    Dim posicaoCaractereSeparador As Integer

    posicaoCaractereSeparador = InStr(1, pItemSelecionadoCombobox, "-", vbTextCompare) - 1
    extraiIdDaStringNaCombobox = Left(pItemSelecionadoCombobox, posicaoCaractereSeparador)


End Function


Function geraId(pPlanilha As Worksheet) As Long
    
    Dim ultimaLinha As Long

    ultimaLinha = ultimaLinhaEmBranco(pPlanilha)
    
    If pPlanilha.Cells(2, 1) = "" Then
        geraId = 1
    Else
        geraId = pPlanilha.Cells(ultimaLinha - 1, 1) + 1
    End If
    
    
End Function

Function indexLinhaRegistroPorId(ByVal pId As Long, pPlanilha As Worksheet) As Long

    Dim registro As Range

    Set registro = pPlanilha.Range("A:A").Find(What:=pId, LookIn:=xlValues, LookAt:=xlWhole)

    If Not registro Is Nothing Then
        indexLinhaRegistroPorId = registro.Row
    Else: MsgBox "Registro nao encontrado.", vbCritical
        indexLinhaRegistroPorId = 0
    End If


End Function

'=========================================
'FUNÇÃO QUE VERIFICA SE UM ARRAY FOI
'INICIALIZADO, COM DADOS OU SE ESTA VAZIO
'=========================================
Function arrayIniciado(ByRef arr() As Variant, Optional pMostrarMensagem As Boolean = False) As Boolean
    On Error Resume Next
    arrayIniciado = IsNumeric(UBound(arr))
    If arrayIniciado = False Then
        If pMostrarMensagem Then
            MsgBox "Um array de dados vazio/não inicializado está sendo utilizado para coleta de dados." _
                 + vbNewLine + vbNewLine + "Provavelmente alguma pesquisa retornou uma lista vazia ou algum ID/registro não foi encontrado para que fosse possível popular tal array.", vbCritical
        End If
    End If
    On Error GoTo 0
End Function
