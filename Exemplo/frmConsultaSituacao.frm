VERSION 5.00
Begin VB.Form frmConsultaSituacao 
   Caption         =   "frmConsultaSituacao"
   ClientHeight    =   4380
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5325
   LinkTopic       =   "Form1"
   ScaleHeight     =   4380
   ScaleWidth      =   5325
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnConsultar 
      Caption         =   "Consultar Situação NFCe"
      Height          =   735
      Left            =   1440
      TabIndex        =   2
      Top             =   2040
      Width           =   2175
   End
   Begin VB.TextBox txttpAmb 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   "Tipo de Ambiente 1 ou 2"
      Top             =   720
      Width           =   3015
   End
   Begin VB.TextBox txtchNFCe 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "Chave da NFCe"
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "frmConsultaSituacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnConsultar_Click()

Dim retorno As String
    
    Dim status As String
    Dim motivo As String
    Dim nfeProc As String
    Dim cStat As String
    Dim xMotivo As String
    Dim chNFe As String
    Dim dhRegEvento As String
    Dim nProt As String
    Dim digVal As String
    
    retorno = consultarSituacao(txtchNFCe, txttpAmb)
    status = LerDadosJSON(retorno, "status", "", "")
    
    If (status = 100) Then
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

