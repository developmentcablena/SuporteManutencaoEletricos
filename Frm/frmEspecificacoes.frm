VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmEspecificacoes 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Suporte Manutenção - Editar Especificações"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6150
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEspecificacoes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   6150
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdFechar 
      Caption         =   "Fechar"
      Height          =   375
      Left            =   4800
      TabIndex        =   0
      Top             =   1680
      Width           =   1215
   End
   Begin GridEX20.GridEX gexEspecificacoes 
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   2566
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      RecordNavigator =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      HeaderStyle     =   2
      MethodHoldFields=   -1  'True
      AllowDelete     =   -1  'True
      BorderStyle     =   3
      GroupByBoxVisible=   0   'False
      ColumnHeaderHeight=   285
      IntProp1        =   0
      ColumnsCount    =   3
      Column(1)       =   "frmEspecificacoes.frx":0442
      Column(2)       =   "frmEspecificacoes.frx":066E
      Column(3)       =   "frmEspecificacoes.frx":07E2
      FormatStylesCount=   5
      FormatStyle(1)  =   "frmEspecificacoes.frx":099E
      FormatStyle(2)  =   "frmEspecificacoes.frx":0ACA
      FormatStyle(3)  =   "frmEspecificacoes.frx":0B7A
      FormatStyle(4)  =   "frmEspecificacoes.frx":0C2E
      FormatStyle(5)  =   "frmEspecificacoes.frx":0D06
      ImageCount      =   0
      PrinterProperties=   "frmEspecificacoes.frx":0DBE
   End
End
Attribute VB_Name = "frmEspecificacoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strSQL As String
Private rs As ADODB.Recordset

Private Sub cmdFechar_Click()
    Call Unload(Me)
End Sub

Private Sub gexEspecificacoes_AfterUpdate()
    MsgBox "Item atualizado com sucesso!", vbOKOnly + vbInformation, "Suporte Manutenção"
End Sub

Private Sub gexEspecificacoes_BeforeDelete(ByVal Cancel As GridEX20.JSRetBoolean)
    If MsgBox("Tem certeza que deseja excluir este item?", vbYesNo + vbQuestion, "Suporte Manutenção") = vbYes Then
        If fnPesquisarEspecificacoes(gexEspecificacoes.Value(1)) = True Then
            MsgBox "Item já utilizado em OS!" & vbCrLf & "Item só pode ser inativado!", vbOKOnly + vbInformation, "Suporte Manutenção"
            Cancel = True
            Exit Sub
        End If
    Else
        Cancel = True
        Exit Sub
    End If
    MsgBox "Item excluído com sucesso!", vbOKOnly + vbInformation, "Suporte Manutenção"
End Sub

Private Function fnPesquisarEspecificacoes(ByVal vEspecificacaoID As Integer) As Boolean
    fnPesquisarEspecificacoes = False
    strSQL = "SELECT * FROM tb_OS WHERE EspecificacaoID = " & vEspecificacaoID
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        fnPesquisarEspecificacoes = True
    End If
    
    rs.Close
    Set rs = Nothing
End Function


