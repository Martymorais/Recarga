VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   840
      Left            =   360
      TabIndex        =   2
      Top             =   1500
      Width           =   1410
   End
   Begin VB.TextBox txt_VQ_O 
      Height          =   585
      Left            =   1815
      TabIndex        =   1
      Text            =   "0"
      Top             =   390
      Width           =   1050
   End
   Begin VB.TextBox txt_VQ_Q 
      Height          =   510
      Left            =   150
      TabIndex        =   0
      Text            =   "0"
      Top             =   465
      Width           =   900
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    Call Enche_Matriz_Elementos(200, 20, CInt(Text1.Text), CInt(Text2.Text))
    Call Enche_HE(20, CInt(Text1.Text), CInt(Text2.Text))
    Call Enche_AQ(0.0000001, 20)
    Call Enche_Matriz_Onde_Parar

    Semente = 5252521

    Call Algoritmo_Formiga_Veneno(10, 2, 0.9, 0.3, 1, 0.0000001, CInt(Text1.Text), CInt(Text2.Text))
    

 
End Sub

Private Sub txt_VQ_O_LostFocus()

    If IsNull(txt_VQ_O.Text) Or (txt_VQ_O.Text = "") Or Not IsNumeric(txt_VQ_O.Text) Then
        txt_VQ_O.SetFocus
        txt_VQ_O.Text = 0
        txt_VQ_O.SelStart = 0
        txt_VQ_O.SelLength = 1
    Else
        If CInt(txt_VQ_O.Text) < 0 Or CInt(txt_VQ_O.Text) > 8 Then
            MsgBox "Este valor deve estar entre 0 e 8"
            txt_VQ_O.SetFocus
            txt_VQ_O.Text = 0
            txt_VQ_O.SelStart = 0
            txt_VQ_O.SelLength = 1
        Else
            txt_Num_Agentes.Text = 200 - (CInt(txt_VQ_Q.Text) * 4) - (CInt(txt_VQ_O.Text) * 2)
        End If
    End If
    
End Sub

Private Sub txt_VQ_Q_LostFocus()

    If IsNull(txt_VQ_Q.Text) Or (txt_VQ_Q.Text = "") Or Not IsNumeric(txt_VQ_Q.Text) Then
        txt_VQ_Q.SetFocus
        txt_VQ_Q.Text = 0
        txt_VQ_Q.SelStart = 0
        txt_VQ_Q.SelLength = 1
    Else
        If CInt(txt_VQ_Q.Text) < 0 Or CInt(txt_VQ_Q.Text) > 8 Then
            MsgBox "Este valor deve estar entre 0 e 6"
            txt_VQ_Q.SetFocus
            txt_VQ_Q.Text = 0
            txt_VQ_Q.SelStart = 0
            txt_VQ_Q.SelLength = 1
        Else
            txt_Num_Agentes.Text = 200 - (CInt(txt_VQ_Q.Text) * 4) - (CInt(txt_VQ_O.Text) * 2)
        End If
    End If
    
End Sub

