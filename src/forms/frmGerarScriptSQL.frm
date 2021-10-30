VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmGerarScriptSQL 
   Caption         =   "Gerador de Script Randomico INSERT SQL"
   ClientHeight    =   1320
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6465
   OleObjectBlob   =   "frmGerarScriptSQL.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmGerarScriptSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnGerar_Click()
    
    If mdl_Util.validaTextBoxes(txtTableName, True) = False Then Exit Sub
    If mdl_Util.validaTextBoxes(txtQtdScriptsSQL, True) = False Then Exit Sub
    
    
    
    Call gerarComandosInsertsSQL(txtTableName, txtQtdScriptsSQL)
    
    
    
End Sub

Private Sub txtQtdScriptsSQL_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
     Call mdl_Util.textBoxSomenteNumerosOptionalMoeda(KeyAscii)
End Sub
