VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmCaracteristicas 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Suporte Manutenção - Editar Caracteristicas"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6135
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCaracteristicas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   6135
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdFechar 
      Caption         =   "Fechar"
      Height          =   375
      Left            =   4800
      TabIndex        =   0
      Top             =   1680
      Width           =   1215
   End
   Begin GridEX20.GridEX gexCaracteristicas 
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
      Column(1)       =   "frmCaracteristicas.frx":0442
      Column(2)       =   "frmCaracteristicas.frx":063E
      Column(3)       =   "frmCaracteristicas.frx":0796
      FormatStylesCount=   5
      FormatStyle(1)  =   "frmCaracteristicas.frx":0952
      FormatStyle(2)  =   "frmCaracteristicas.frx":0A7E
      FormatStyle(3)  =   "frmCaracteristicas.frx":0B2E
      FormatStyle(4)  =   "frmCaracteristicas.frx":0BE2
      FormatStyle(5)  =   "frmCaracteristicas.frx":0CBA
      ImageCount      =   0
      PrinterProperties=   "frmCaracteristicas.frx":0D72
   End
End
Attribute VB_Name = "frmCaracteristicas"
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

Private Sub gexCaracteristicas_AfterUpdate()
    MsgBox "Item atualizado com sucesso!", vbOKOnly + vbInformation, "Suporte Manutenção"
End Sub

Private Sub gexCaracteristicas_BeforeDelete(ByVal Cancel As GridEX20.JSRetBoolean)
    If MsgBox("Tem certeza que deseja excluir este item?", vbYesNo + vbQuestion, "Suporte Manutenção") = vbYes Then
        If fnPesquisarCaracteristica(gexCaracteristicas.Value(1)) = True Then
            MsgBox "Item já utilizado em OS!" & vbCrLf & "Item só pode ser inativado!", vbOKOnly + vbInformation, "Suporte Manutenção"
            Cancel = True
            Exit Sub
        End If
        
        If fnPesquisarSubItem(gexCaracteristicas.Value(1)) = True Then
            MsgBox "Item possui sub-itens cadastrados!" & vbCrLf & "Item só pode ser inativado!", vbOKOnly + vbCritical, "Suporte Manutenção"
            Cancel = True
            Exit Sub
        End If
    Else
        Cancel = True
        Exit Sub
    End If
    MsgBox "Item excluído com sucesso!", vbOKOnly + vbInformation, "Suporte Manutenção"
End Sub

Private Function fnPesquisarSubItem(ByVal vTipoID As Integer) As Boolean
    fnPesquisarSubItem = False
    strSQL = "SELECT * FROM tb_Especificacoes WHERE TipoID = " & vTipoID
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        fnPesquisarSubItem = True
    End If
    
    rs.Close
    Set rs = Nothing
End Function

Private Function fnPesquisarCaracteristica(ByVal vTipoID As Integer) As Boolean
    fnPesquisarCaracteristica = False
    strSQL = "SELECT * FROM tb_OS WHERE TipoID = " & vTipoID
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        fnPesquisarCaracteristica = True
    End If
    
    rs.Close
    Set rs = Nothing
End Function

