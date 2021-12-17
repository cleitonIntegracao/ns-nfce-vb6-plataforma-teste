VERSION 5.00
Begin VB.Form frmDownloadNFCe 
   Caption         =   "frmDownloadNFCe"
   ClientHeight    =   4665
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   ScaleHeight     =   4665
   ScaleWidth      =   4950
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtImprimiePDF 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Text            =   "Imprimir PDF True ou False"
      Top             =   2160
      Width           =   3015
   End
   Begin VB.TextBox txtExibirTela 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Text            =   "Exibir na tela True ou False"
      Top             =   1680
      Width           =   3015
   End
   Begin VB.TextBox txtCaminhoSalvar 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Text            =   "Caminho aonde vai ser salvo o PDF"
      Top             =   1200
      Width           =   3015
   End
   Begin VB.CommandButton cmdFazerDown 
      Caption         =   "Download"
      Height          =   735
      Left            =   1560
      TabIndex        =   2
      Top             =   3000
      Width           =   1575
   End
   Begin VB.TextBox txttpAmb 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   "Tipo de Amb 1 ou 2"
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox txtchNFCe 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "Chave da NFCe para Download"
      Top             =   240
      Width           =   4335
   End
End
Attribute VB_Name = "frmDownloadNFCe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdFazerDown_Click()
    Dim retorno As String
    
    Dim status As String
    Dim motivo As String
    Dim nfeProc As String
    Dim nProt As String
    Dim digVal As String
    Dim chNFe As String
    Dim serie As String
    Dim numero As String
    
    retorno = downloadNFCeESalvar(txtchNFCe, txttpAmb, txtCaminhoSalvar, txtExibirTela, txtImprimiePDF)
    status = LerDadosJSON(retorno, "status", "", "")
    
    If (status = 100) Then
        motivo = LerDadosJSON(retorno, "motivo", "", "")
        MsgBox (retorno)
    End If
        
        If (status = -100) Then
           motivo = LerDadosJSON(retorno, "motivo", "", "")
           MsgBox (motivo)
         End If
End Sub

