VERSION 5.00
Begin VB.Form frmEntrada_Dados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entrada de Dados"
   ClientHeight    =   30
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1755
   LinkTopic       =   "Form1"
   ScaleHeight     =   30
   ScaleWidth      =   1755
End
Attribute VB_Name = "frmEntrada_dados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************
'
'*******************************************************************************
Private Sub Form_Load()
    
    'frmEntrada_dados.Hide
    Call Avalia_Aptidao

End Sub
'*******************************************************************************

