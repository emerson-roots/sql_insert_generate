Attribute VB_Name = "Enums"
Option Explicit

Private Const cTexto As String = "TEXTO"
Private Const cInteiro As String = "INTEIRO"
Private Const cLongo As String = "LONGO"
Private Const cDouble_Num As String = "DOUBLE"
Private Const cData_ As String = "DATA"
Private Const cCaractere As String = "CARACTERE"
Private Const cFkString As String = "FK_STRING"

Enum TiposVariaveis
  TEXTO = 1
  INTEIRO = 2
  LONGO = 3
  DOUBLE_NUM = 4
  DATA_ = 5
  CARACTERE = 6
  FK_STRING = 7
End Enum


Function converteEnumTipoVariavelToString(ByVal pIndexEnum As TiposVariaveis) As String
    
    Dim textoEnum As String
    
    Select Case pIndexEnum
    
        Case TiposVariaveis.TEXTO
            textoEnum = cTexto
        Case TiposVariaveis.INTEIRO
            textoEnum = cInteiro
        Case TiposVariaveis.LONGO
            textoEnum = cInteiro
        Case TiposVariaveis.DOUBLE_NUM
            textoEnum = cDouble_Num
        Case TiposVariaveis.DATA_
            textoEnum = cData_
        Case TiposVariaveis.CARACTERE
            textoEnum = cCaractere
        Case TiposVariaveis.FK_STRING
            textoEnum = cFkString
            
    End Select
    
    converteEnumTipoVariavelToString = textoEnum

End Function
