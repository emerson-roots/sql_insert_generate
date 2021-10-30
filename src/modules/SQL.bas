Attribute VB_Name = "SQL"
Option Explicit
Sub gerarComandosInsertsSQL()

    Dim arrLocal() As Variant
    Dim SQL As String
    Dim values As String
    Dim dado As String
    Dim i As Integer
    Dim z As Integer
    Dim qtdInserts As Integer
    Dim tipoVariavel As String
    Dim rangeMinimo As Variant
    Dim rangeMaximo As Variant
    Dim colunaTabela As String
    Dim qtdLinhas As Integer
    Dim arrComandosSql() As Variant
    
    ' verifica se o banco de palavras esta populado para aumentar performance
    If mdl_Util.arrayIniciado(arrBancoDePalavrasDoDicionario) = False Then
        arrBancoDePalavrasDoDicionario = arrayComDadosDeTabela(wsBancoDePalavras)
    End If

    'captura estrutura de colunas da tabela
    arrLocal = mdl_Util.arrayComDadosDeTabela(wsTabela)

    qtdInserts = 10000
    qtdLinhas = UBound(arrLocal, 2)

    'seta quantidade de inserts para enviar ao arquivo/script
    ReDim arrComandosSql(qtdInserts - 1)


    For i = 1 To qtdInserts
    
        For z = 0 To qtdLinhas
            On Error GoTo errorHandler
            colunaTabela = colunaTabela & arrLocal(0, z)
            
            tipoVariavel = arrLocal(1, z)
            rangeMinimo = arrLocal(2, z)
            rangeMaximo = arrLocal(3, z)
            dado = geraDadoRandomico(tipoVariavel, rangeMinimo, rangeMaximo)
            values = values & dado
            
            If z < qtdLinhas Then
                values = values & ", "
                colunaTabela = colunaTabela & ", "
            End If
    
errorHandler:

            If Err.Number <> 0 Then
                MsgBox "Ocorreu um erro ao gerar SQL o processo será abortado.", vbCritical
                Exit Sub
            End If
        
        Next z
        
        SQL = "INSERT INTO tbTabela(" & colunaTabela & ") VALUES (" & values & ");"
        
        ' adiciona comando insert no array
        arrComandosSql(i - 1) = SQL
        
        'limpa valores para nova rodada no loop
        values = ""
        SQL = ""
        colunaTabela = ""
    

    Next i
    
    Call geraArquivoScriptSQL(arrComandosSql)
    
End Sub

Sub geraArquivoScriptSQL(arrayInserts As Variant)

    Dim nomeDiretorioArquivo As String
    Dim i As Integer
    Dim abrirArquivo As Variant

    Dim arquivo As Object
    Set arquivo = CreateObject("ADODB.Stream")
    arquivo.Type = 2 'Specify stream type - we want To save text/string data.
    arquivo.Charset = "utf-8" 'Specify charset For the source text data.
    arquivo.Open 'Open the stream And write binary data To the object


    nomeDiretorioArquivo = ThisWorkbook.Path & "\" & "SQL script" & ".sql"

    For i = 0 To UBound(arrayInserts)
        arquivo.WriteText arrayInserts(i) & vbNewLine
    Next i

    arquivo.SaveToFile nomeDiretorioArquivo, 2
    arquivo.Close

    abrirArquivo = Shell("notepad.exe """ & nomeDiretorioArquivo & """", vbMaximizedFocus)


End Sub

Function geraDadoRandomico(tipo As String, pMin As Variant, pMax As Variant) As Variant


    Dim converteStringToEnum As String
    Dim palavraParaString As String

    palavraParaString = StrConv(arrBancoDePalavrasDoDicionario(1, mdl_DadosAleatorios.randomLong(1, 320000)) & " " & mdl_DadosAleatorios.randomString(5), vbProperCase)


    converteStringToEnum = extraiIdDaStringNaCombobox(tipo)

    Select Case converteStringToEnum

        Case TiposVariaveis.TEXTO
            geraDadoRandomico = "'" & palavraParaString & "'"
        
        Case TiposVariaveis.INTEIRO
            geraDadoRandomico = "'" & mdl_DadosAleatorios.randomInteger(pMin, pMax) & "'"
        
        Case TiposVariaveis.LONGO
            geraDadoRandomico = "'" & mdl_DadosAleatorios.randomLong(pMin, pMax) & "'"
        
        Case TiposVariaveis.DOUBLE_NUM
            geraDadoRandomico = "'" & mdl_DadosAleatorios.randomDouble(pMin, pMax) & "'"
        
        Case TiposVariaveis.DATA_
            'em db oracle, padrão data com hora TO_DATE('2021-11-26 00:00:00', 'YYYY-MM-DD HH24:MI:SS')
            geraDadoRandomico = "TO_DATE('" & mdl_DadosAleatorios.randomDateWithHour & "', 'YYYY-MM-DD HH24:MI:SS')"
            
        Case TiposVariaveis.CARACTERE
            geraDadoRandomico = "'" & mdl_DadosAleatorios.randomCaracter(CStr(pMin)) & "'"
            
        Case TiposVariaveis.FK_STRING
            geraDadoRandomico = "'" & mdl_DadosAleatorios.randomFkTipoString(CStr(pMin)) & "'"
    
    End Select


End Function

