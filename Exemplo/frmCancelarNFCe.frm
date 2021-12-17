VERSION 5.00
Begin VB.Form frmCancelarNFCe 
   Caption         =   "frmCancelarNFCe"
   ClientHeight    =   6135
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   ScaleHeight     =   6135
   ScaleWidth      =   4590
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnCancelamento 
      Caption         =   "Cancelar NFCe"
      Height          =   735
      Left            =   1080
      TabIndex        =   7
      Top             =   4680
      Width           =   2055
   End
   Begin VB.TextBox txtExibirNaTela 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Text            =   "Exibir na tela True ou False"
      Top             =   3720
      Width           =   3855
   End
   Begin VB.TextBox txtCaminho 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Text            =   "Caminho para Salvar"
      Top             =   3240
      Width           =   3855
   End
   Begin VB.TextBox txtxJust 
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Text            =   "Motivo do cancelamento"
      Top             =   2040
      Width           =   3855
   End
   Begin VB.TextBox txtnProt 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Text            =   "Numero do protocolo"
      Top             =   1560
      Width           =   3855
   End
   Begin VB.TextBox txtdhEvento 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Text            =   "2021-12-07T11:49:10-02:00"
      Top             =   1080
      Width           =   4095
   End
   Begin VB.TextBox txttpAmb 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   "Tipo de Ambiente 1 ou 2"
      Top             =   600
      Width           =   3855
   End
   Begin VB.TextBox txtchNFCe 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "Chave da NFCe "
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frmCancelarNFCe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancelamento_Click()
Dim retorno As String
    
    Dim status As String
    Dim motivo As String
    Dim nfeProc As String
    Dim cStat As String
    Dim xMotivo As String
    Dim chNFe As String
    Dim dhRegEvento As String
    Dim nProt As String
    
    retorno = cancelarNFCe(txtchNFCe, txttpAmb, txtdhEvento, txtnProt, txtxJust, txtCaminho, txtExibirNaTela)
    status = LerDadosJSON(retorno, "status", "", "")
    
    If (status = 135) Then
        motivo = LerDadosJSON(retorno, "motivo", "", "")
        MsgBox (retorno)
    End If
        
        If (status = -135) Then
           motivo = LerDadosJSON(retorno, "motivo", "", "")
           cStat = LerDadosJSON(retorno, "nfeProc", "cStat", "")
           xMotivo = LerDadosJSON(retorno, "motivo", "nfeProc", "xMotivo", "")
           MsgBox (retorno)
        End If
        
End Sub
