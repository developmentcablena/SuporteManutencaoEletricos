VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmAcompanharOS 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Suporte Manutenção - Acompanhar OS"
   ClientHeight    =   9135
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13470
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAcompanharOS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9135
   ScaleWidth      =   13470
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdDevolver 
      Appearance      =   0  'Flat
      Caption         =   "&Devolver"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      TabIndex        =   10
      Top             =   8640
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancelar 
      Appearance      =   0  'Flat
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3120
      TabIndex        =   9
      Top             =   8640
      Width           =   1335
   End
   Begin VB.CommandButton cmdImprimir 
      Appearance      =   0  'Flat
      Caption         =   "&Imprimir"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4455
      TabIndex        =   8
      Top             =   8640
      Width           =   1335
   End
   Begin VB.CommandButton cmdAceite 
      Appearance      =   0  'Flat
      Caption         =   "&Aceite"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   8640
      Width           =   1335
   End
   Begin VB.TextBox txtOSID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   2985
      MaxLength       =   5
      TabIndex        =   5
      Top             =   420
      Width           =   1485
   End
   Begin VB.CommandButton cmdPesquisar 
      Appearance      =   0  'Flat
      Caption         =   "&Pesquisar"
      Height          =   315
      Left            =   4560
      TabIndex        =   4
      Top             =   405
      Width           =   1095
   End
   Begin VB.ComboBox cboStatus 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   420
      Width           =   2775
   End
   Begin VB.CommandButton cmdFechar 
      Appearance      =   0  'Flat
      Caption         =   "Fechar"
      Height          =   375
      Left            =   12000
      TabIndex        =   2
      Top             =   8640
      Width           =   1335
   End
   Begin GridEX20.GridEX gexOS 
      Height          =   7575
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   13361
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      RecordNavigator =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      HeaderStyle     =   2
      MethodHoldFields=   -1  'True
      AllowEdit       =   0   'False
      BorderStyle     =   3
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   14
      Column(1)       =   "frmAcompanharOS.frx":058A
      Column(2)       =   "frmAcompanharOS.frx":075E
      Column(3)       =   "frmAcompanharOS.frx":0952
      Column(4)       =   "frmAcompanharOS.frx":0B2A
      Column(5)       =   "frmAcompanharOS.frx":0D36
      Column(6)       =   "frmAcompanharOS.frx":0E92
      Column(7)       =   "frmAcompanharOS.frx":1072
      Column(8)       =   "frmAcompanharOS.frx":125A
      Column(9)       =   "frmAcompanharOS.frx":145E
      Column(10)      =   "frmAcompanharOS.frx":1646
      Column(11)      =   "frmAcompanharOS.frx":178A
      Column(12)      =   "frmAcompanharOS.frx":18FE
      Column(13)      =   "frmAcompanharOS.frx":1A96
      Column(14)      =   "frmAcompanharOS.frx":1BDA
      FormatStylesCount=   5
      FormatStyle(1)  =   "frmAcompanharOS.frx":1D5A
      FormatStyle(2)  =   "frmAcompanharOS.frx":1E86
      FormatStyle(3)  =   "frmAcompanharOS.frx":1F36
      FormatStyle(4)  =   "frmAcompanharOS.frx":1FEA
      FormatStyle(5)  =   "frmAcompanharOS.frx":20C2
      ImageCount      =   0
      PrinterProperties=   "frmAcompanharOS.frx":217A
   End
   Begin VB.Label lblOSID 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nº OS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2985
      TabIndex        =   6
      Top             =   180
      Width           =   555
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   195
      Width           =   615
   End
End
Attribute VB_Name = "frmAcompanharOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rs As ADODB.Recordset
Private strSQL As String

Private Sub cboStatus_Click()
    
    Select Case cboStatus.Text
        Case Is = "Em Aberto"
            cmdAceite.Enabled = False
            cmdDevolver.Enabled = False
            cmdCancelar.Enabled = True
            cmdImprimir.Enabled = True
        Case Is = "Urgente"
            cmdAceite.Enabled = False
            cmdDevolver.Enabled = False
            cmdCancelar.Enabled = True
            cmdImprimir.Enabled = True
        Case Is = "Em Atendimento"
            cmdAceite.Enabled = False
            cmdDevolver.Enabled = False
            cmdCancelar.Enabled = False
            cmdImprimir.Enabled = False
        Case Is = "Aguardando Aceite"
            cmdAceite.Enabled = True
            cmdDevolver.Enabled = True
            cmdCancelar.Enabled = False
            cmdImprimir.Enabled = False
        Case Is = "Finalizada"
            cmdAceite.Enabled = False
            cmdDevolver.Enabled = False
            cmdCancelar.Enabled = False
            cmdImprimir.Enabled = False
        Case Is = "Cancelada"
            cmdAceite.Enabled = False
            cmdDevolver.Enabled = False
            cmdCancelar.Enabled = False
            cmdImprimir.Enabled = False
        Case Is = "Não Validada"
            cmdAceite.Enabled = False
            cmdDevolver.Enabled = False
            cmdCancelar.Enabled = False
            cmdImprimir.Enabled = False
    End Select
    
    'Call suListarOS(gintUsuarioID, cboStatus.Text)
    Call suListarOS(gstrDepto, cboStatus.Text)
End Sub

Private Sub cmdAceite_Click()
Dim i As Integer

    gintOSID = 0

    If gexOS.RowCount = 0 Then
        MsgBox "Não há nenhuma OS na lista!", vbOKOnly + vbExclamation, "Suporte Manutenção"
        Exit Sub
    End If

    For i = 1 To gexOS.RowCount
        If gexOS.RowSelected(i) = True Then
            gintOSID = gexOS.Value(1)
            Call frmAceite.Show(vbModal)
            Call cboStatus_Click
        End If
    Next

End Sub

Private Sub cmdCancelar_Click()
Dim i As Integer

    gintOSID = 0
    
    If gexOS.RowCount = 0 Then
        MsgBox "Não há nenhum chamado listado!", vbOKOnly + vbExclamation, "Suporte Manutenção"
        Exit Sub
    End If
    
    For i = 1 To gexOS.RowCount
        If gexOS.RowSelected(i) = True Then
            gintOSID = gexOS.Value(1)
            If MsgBox("Tem certeza que deseja cancelar a OS " & Format(gintOSID, "0000") & "?", vbYesNo + vbQuestion, "Suporte Manutenção") = vbYes Then
                'Call suCancelarOS(gintOSID)
                Call frmCancelarOS.Show(vbModal)
                Call suListarOS(gintUsuarioID, cboStatus.Text)
            End If
        End If
    Next

End Sub

Private Sub cmdDevolver_Click()
Dim i As Integer

    gintOSID = 0

    If gexOS.RowCount = 0 Then
        MsgBox "Não há nenhuma OS na lista!", vbOKOnly + vbExclamation, "Suporte Manutenção"
        Exit Sub
    End If
    
    For i = 1 To gexOS.RowCount
        If gexOS.RowSelected(i) = True Then
            gintOSID = gexOS.Value(1)
            With frmAceite
                .Caption = "Suporte Manutenção - Devolver OS"
            End With
            Call frmAceite.Show(vbModal)
            Call cboStatus_Click
        End If
    Next

End Sub

Private Sub cmdFechar_Click()
    Call Unload(Me)
End Sub

Private Sub cmdImprimir_Click()
Dim i As Integer
    
    gintOSID = 0
    
    If gexOS.RowCount = 0 Then
        MsgBox "Não há nenhum chamado listado!", vbOKOnly + vbExclamation, "Suporte Manutenção"
        Exit Sub
    End If
    
    For i = 1 To gexOS.RowCount
        If gexOS.RowSelected(i) = True Then
            gintOSID = gexOS.Value(1)
            Call acrOS.Show(vbModal)
        End If
    Next
End Sub

Private Sub cmdPesquisar_Click()
    If Len(Trim(cboStatus.Text)) = 0 Then
        MsgBox "Selecione um status!", vbOKOnly + vbExclamation, "Suporte Manutenção"
        cboStatus.SetFocus
        Exit Sub
    End If
    
    If Len(Trim(txtOSID.Text)) > 0 Then
        If IsNumeric(txtOSID.Text) = True Then
            If fnPesquisarOSID(Trim(txtOSID.Text), gstrDepto, cboStatus.Text) = False Then
                MsgBox "Ordem de serviço não localizada!", vbOKOnly + vbExclamation, "Suporte Manutenção"
            End If
        Else
            MsgBox "Digite um valor numérico!", vbOKOnly + vbExclamation, "Suporte Manutenção"
        End If
    Else
        MsgBox "Digite o código da ordem de serviço!", vbOKOnly + vbExclamation, "Suporte Manutenção"
    End If
End Sub

Private Sub gexOS_DblClick()
Dim i As Integer
    
    gintOSID = 0
    gstrTela = ""
    
    If gexOS.RowCount = 0 Then
        MsgBox "Não há nenhum chamado listado!", vbOKOnly + vbExclamation, "Suporte Manutenção"
        Exit Sub
    End If
    
    For i = 1 To gexOS.RowCount
        If gexOS.RowSelected(i) = True Then
            gintOSID = gexOS.Value(1)
            gstrTela = "frmAcompanharOS"
            Call suMostrarOS(gintOSID, cboStatus.Text, gexOS.Value(10))
            Call frmSuporteSistemas.Show(vbModal)
            gstrTela = ""
        End If
    Next
End Sub

Private Sub Form_Load()
    Call suListarStatus
End Sub

Private Sub suCancelarOS(ByVal vOSID As Long)
On Error GoTo Erro

    strSQL = "UPDATE tb_OS SET DataCancelamento = '" & Format(Now, "dd/MM/yyyy HH:mm:ss") & "',Status = 4 WHERE OSID = " & vOSID & ""
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenKeyset, adLockOptimistic
    
    Set rs = Nothing
    MsgBox "OS " & Format(vOSID, "0000") & " cancelada com sucesso!", vbOKOnly + vbInformation, "Suporte Manutenção"
    Exit Sub
    
Erro:
    Set rs = Nothing
    MsgBox "Erro: " & Err.Description, vbOKOnly + vbCritical, "Suporte Manutenção"

End Sub

Private Sub suMostrarOS(ByVal vOSID As Long, ByVal vStatus As String, ByVal vUsuario As String)
    strSQL = "SELECT * FROM vw_Chamados WHERE OSID = " & vOSID
    'strSQL = "SELECT REPLACE(LEFT(Especificacao,CHARINDEX('-',Especificacao)),'-','') AS Especificacao,Divisao,Tipo,DescricaoServico,Prioridade,Prazo,Previsao,Atendente,ReporteTecnico FROM vw_Chamados WHERE OSID = " & vOSID
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        If vStatus = "Em Aberto" And gstrNome = vUsuario Then
            With frmSuporteSistemas
                .cboDivisao.Enabled = True
                .cboDivisao.Text = rs!Divisao
                .cboTipo.Enabled = True
                .cboTipo.Text = rs!Tipo
                .cboLinha.Enabled = True
                .cboLinha.Text = rs!Linha
                .cboEspecificacao.Enabled = True
                .cboEspecificacao.Text = rs!Especificacao
                .txtObservacao.Enabled = True
                .txtObservacao.Text = rs!DescricaoServico
                '.txtReporteTecnico.Text = rs!ReporteTecnico & ""
                .chkPrioridade.Enabled = True
                .chkPrioridade.Value = IIf(rs!Prioridade = True, 1, 0)
                .txtPrazo.Enabled = True
                .txtPrazo.Text = rs!Prazo
                .txtPrevisao.Enabled = False
                .txtPrevisao.Text = rs!Previsao
                If rs!Atendente <> "" Then
                    .cboAtendente.Text = rs!Atendente
                End If
                .lblQtdeCaracteres.Enabled = True
                .cmdCadastrar.Caption = "&Alterar"
                .cmdCadastrar.Enabled = True
                .cmdCancelar.Enabled = True
            End With
        Else
            With frmSuporteSistemas
                .cboDivisao.Enabled = False
                .cboDivisao.Text = rs!Divisao
                .cboTipo.Enabled = False
                .cboTipo.Text = rs!Tipo
                .cboLinha.Enabled = False
                .cboLinha.Text = rs!Linha
                .cboEspecificacao.Enabled = False
                .cboEspecificacao.Text = rs!Especificacao
                .txtObservacao.Enabled = True
                .txtObservacao.Locked = True
                .txtObservacao.Text = rs!DescricaoServico
                .txtReporteTecnico.Enabled = True
                .txtReporteTecnico.Locked = True
                .txtReporteTecnico.Text = rs!ReporteTecnico & ""
                .chkPrioridade.Enabled = False
                .chkPrioridade.Value = IIf(rs!Prioridade = True, 1, 0)
                .txtPrazo.Enabled = False
                .txtPrazo.Text = rs!Prazo
                .txtPrevisao.Enabled = False
                .txtPrevisao.Text = rs!Previsao
                If rs!Atendente <> "" Then
                    .cboAtendente.Enabled = False
                    .cboAtendente.Text = rs!Atendente
                Else
                    .cboAtendente.Enabled = False
                End If
                .lblQtdeCaracteres.Enabled = False
                .cmdCadastrar.Enabled = False
                .cmdCancelar.Enabled = False
            End With
        End If
    End If
    
    rs.Close
    Set rs = Nothing
End Sub

Private Function fnPesquisarOSID(ByVal vOSID As Long, ByVal vDepto As String, ByVal vStatus As String) As Boolean
    fnPesquisarOSID = False
    
    If vStatus = "Em Aberto" Then
        strSQL = "SELECT * FROM vw_Chamados WHERE OSID = " & vOSID & " AND Departamento = '" & vDepto & "' AND Status = 0 ORDER BY OSID"
    ElseIf vStatus = "Em Atendimento" Then
        strSQL = "SELECT * FROM vw_Chamados WHERE OSID = " & vOSID & " AND Departamento = '" & vDepto & "' AND Status = 1 ORDER BY OSID"
    ElseIf vStatus = "Aguardando Aceite" Then
        strSQL = "SELECT * FROM vw_Chamados WHERE OSID = " & vOSID & " AND Departamento = '" & vDepto & "' AND Status = 2 ORDER BY OSID"
    ElseIf vStatus = "Finalizada" Then
        strSQL = "SELECT * FROM vw_Chamados WHERE OSID = " & vOSID & " AND Departamento = '" & vDepto & "' AND Status = 3 ORDER BY OSID"
    ElseIf vStatus = "Cancelada" Then
        strSQL = "SELECT * FROM vw_Chamados WHERE OSID = " & vOSID & " AND Departamento = '" & vDepto & "' AND Status = 4 ORDER BY OSID"
    End If
    
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        gexOS.HoldFields
        Set gexOS.ADORecordset = rs
        fnPesquisarOSID = True
    Else
        gexOS.HoldFields
        Set gexOS.ADORecordset = rs
    End If
    
    rs.Close
    Set rs = Nothing
End Function

Private Sub suListarOS(ByVal vDepto As String, ByVal vStatus As String)
    
    If vStatus = "Em Aberto" Then
        strSQL = "SELECT * FROM vw_Chamados WHERE Departamento = '" & vDepto & "' AND Status = 0 ORDER BY OSID"
    ElseIf vStatus = "Em Atendimento" Then
        strSQL = "SELECT * FROM vw_Chamados WHERE Departamento = '" & vDepto & "' AND Status = 1 ORDER BY OSID"
    ElseIf vStatus = "Aguardando Aceite" Then
        strSQL = "SELECT * FROM vw_Chamados WHERE Departamento = '" & vDepto & "' AND Status = 2 ORDER BY OSID"
    ElseIf vStatus = "Finalizada" Then
        strSQL = "SELECT * FROM vw_Chamados WHERE Departamento = '" & vDepto & "' AND Status = 3 ORDER BY OSID"
    ElseIf vStatus = "Cancelada" Then
        strSQL = "SELECT * FROM vw_Chamados WHERE Departamento = '" & vDepto & "' AND Status = 4 ORDER BY OSID"
    ElseIf vStatus = "Não Validada" Then
        strSQL = "SELECT * FROM vw_Chamados WHERE Departamento = '" & vDepto & "' AND Status = 6 ORDER BY OSID"
    End If
        
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenKeyset, adLockOptimistic
    
    If rs.EOF = False Then
        gexOS.HoldFields
        Set gexOS.ADORecordset = rs
    Else
        gexOS.HoldFields
        Set gexOS.ADORecordset = rs
    End If
        
    Set rs = Nothing
End Sub

'Private Sub suListarOS(ByVal vUsuarioID As Integer, ByVal vStatus As String)
'
'    If vStatus = "Em Aberto" Then
'        strSQL = "SELECT * FROM vw_Chamados WHERE UsuarioID = " & vUsuarioID & " AND Status = 0 ORDER BY OSID"
'    ElseIf vStatus = "Em Atendimento" Then
'        strSQL = "SELECT * FROM vw_Chamados WHERE UsuarioID = " & vUsuarioID & " AND Status = 1 ORDER BY OSID"
'    ElseIf vStatus = "Aguardando Aceite" Then
'        strSQL = "SELECT * FROM vw_Chamados WHERE UsuarioID = " & vUsuarioID & " AND Status = 2 ORDER BY OSID"
'    ElseIf vStatus = "Finalizada" Then
'        strSQL = "SELECT * FROM vw_Chamados WHERE UsuarioID = " & vUsuarioID & " AND Status = 3 ORDER BY OSID"
'    ElseIf vStatus = "Cancelada" Then
'        strSQL = "SELECT * FROM vw_Chamados WHERE UsuarioID = " & vUsuarioID & " AND Status = 4 ORDER BY OSID"
'    ElseIf vStatus = "Não Validada" Then
'        strSQL = "SELECT * FROM vw_Chamados WHERE UsuarioID = " & vUsuarioID & " AND Status = 6 ORDER BY OSID"
'    End If
'
'    Set rs = New ADODB.Recordset
'    rs.Open strSQL, cn, adOpenKeyset, adLockOptimistic
'
'    If rs.EOF = False Then
'        gexOS.HoldFields
'        Set gexOS.ADORecordset = rs
'    Else
'        gexOS.HoldFields
'        Set gexOS.ADORecordset = rs
'    End If
'
'    Set rs = Nothing
'End Sub

Private Sub suListarStatus()
    cboStatus.AddItem "Em Aberto"
    cboStatus.AddItem "Em Atendimento"
    cboStatus.AddItem "Aguardando Aceite"
    cboStatus.AddItem "Finalizada"
    cboStatus.AddItem "Cancelada"
    cboStatus.AddItem "Não Validada"
End Sub

Private Sub txtOSID_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Len(Trim(txtOSID.Text)) > 0 Then
            Call cmdPesquisar_Click
        End If
    End If
End Sub
