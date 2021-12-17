VERSION 5.00
Begin VB.Form frmInutilizar 
   Caption         =   "Form1"
   ClientHeight    =   5055
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4485
   LinkTopic       =   "frmInutilizar"
   ScaleHeight     =   5055
   ScaleWidth      =   4485
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnInutilizar 
      Caption         =   "Inutilizar Numeração NFCe"
      Height          =   975
      Left            =   2280
      TabIndex        =   8
      Top             =   1080
      Width           =   1935
   End
   Begin VB.TextBox txtxJust 
      Height          =   1095
      Left            =   120
      TabIndex        =   7
      Text            =   "Justificativa da Inutilizacao"
      Top             =   3480
      Width           =   4215
   End
   Begin VB.TextBox txtNumeroFin 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Text            =   "Numero Final"
      Top             =   3000
      Width           =   1815
   End
   Begin VB.TextBox txtNumeroIni 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Text            =   "Numero Inicial "
      Top             =   2520
      Width           =   1815
   End
   Begin VB.TextBox txtSerie 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Text            =   "Serie"
      Top             =   2040
      Width           =   1815
   End
   Begin VB.TextBox txtCNPJ 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Text            =   "CNPJ Emitente"
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox txtAno 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Text            =   "Ano EX: 21"
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox txttpAmb 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   "Tipo de Ambiente "
      Top             =   600
      Width           =   1815
   End
   Begin VB.TextBox txtcUF 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "Código UF "
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmInutilizar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnInutilizar_Click()

Dim retorno As String
    
    Dim status As String
    Dim motivo As String
    Dim retInutNFe As String
    Dim cStat As String
    Dim xMotivo As String
    Dim nProt As String
    Dim dhRecbto As String
    Dim xml As String
    
    retorno = inutilizar(txtcUF, txttpAmb, txtAno, txtCNPJ, txtSerie, txtNumeroIni, txtNumeroFin, txtxJust)
    status = LerDadosJSON(retorno, "status", "", "")
    
    If (status = 102) Then
        motivo = LerDadosJSON(retorno, "motivo", "", "")
        MsgBox (retorno)
    End If
        
        If (status = -10) Then
           motivo = LerDadosJSON(retorno, "motivo", "", "")
           cStat = LerDadosJSON(retorno, "retInuNFe", "cStat", "")
           xMotivo = LerDadosJSON(retorno, "retInuNFe", "xMotivo", "")
           dhRecbto = LerDadosJSON(retorno, "retInuNFe", "dhRecbto", "")
           xml = LerDadosJSON(retorno, "retInuNFe", "xml", "")
           MsgBox (motivo)
         End If
End Sub
