VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmChamados 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Suporte Manutenção - Verificar Chamados"
   ClientHeight    =   10365
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16215
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmChamados.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10365
   ScaleWidth      =   16215
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSituacao 
      Appearance      =   0  'Flat
      Caption         =   "&Situação"
      Height          =   375
      Left            =   6000
      TabIndex        =   14
      Top             =   9840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox chk1 
      Appearance      =   0  'Flat
      Caption         =   "Minhas Ordens de Serviço"
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   4320
      TabIndex        =   13
      Top             =   120
      Width           =   2535
   End
   Begin VB.CheckBox chkHistorico 
      Appearance      =   0  'Flat
      Caption         =   "Visualizar Ocorrências"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   8160
      TabIndex        =   12
      Top             =   435
      Width           =   1335
   End
   Begin VB.CommandButton cmdDevolver 
      Appearance      =   0  'Flat
      Caption         =   "&Devolver"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      TabIndex        =   11
      Top             =   9840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox chkComentario 
      Appearance      =   0  'Flat
      Caption         =   "Visualizar Comentário"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   6360
      TabIndex        =   10
      Top             =   435
      Width           =   1455
   End
   Begin VB.CommandButton cmdAtender 
      Appearance      =   0  'Flat
      Caption         =   "&Atender"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   9840
      Width           =   1335
   End
   Begin VB.CommandButton cmdPesquisar 
      Appearance      =   0  'Flat
      Caption         =   "&Pesquisar"
      Height          =   300
      Left            =   4320
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txtOSID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2880
      MaxLength       =   5
      TabIndex        =   1
      Top             =   495
      Width           =   1335
   End
   Begin VB.CommandButton cmdBaixarOS 
      Appearance      =   0  'Flat
      Caption         =   "&Finalizar OS"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4560
      TabIndex        =   7
      Top             =   9840
      Width           =   1335
   End
   Begin VB.CommandButton cmdImprimir 
      Appearance      =   0  'Flat
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   9840
      Width           =   1335
   End
   Begin VB.ComboBox cboStatus 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   480
      Width           =   2655
   End
   Begin VB.CommandButton cmdFechar 
      Appearance      =   0  'Flat
      Caption         =   "Fechar"
      Height          =   375
      Left            =   14760
      TabIndex        =   4
      Top             =   9840
      Width           =   1335
   End
   Begin GridEX20.GridEX gexOS 
      Height          =   8655
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   15975
      _ExtentX        =   28178
      _ExtentY        =   15266
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
      ColumnsCount    =   15
      Column(1)       =   "frmChamados.frx":1CCA
      Column(2)       =   "frmChamados.frx":1EC2
      Column(3)       =   "frmChamados.frx":202E
      Column(4)       =   "frmChamados.frx":21BA
      Column(5)       =   "frmChamados.frx":245E
      Column(6)       =   "frmChamados.frx":2636
      Column(7)       =   "frmChamados.frx":284E
      Column(8)       =   "frmChamados.frx":2A6A
      Column(9)       =   "frmChamados.frx":2BC6
      Column(10)      =   "frmChamados.frx":2DAE
      Column(11)      =   "frmChamados.frx":2FBA
      Column(12)      =   "frmChamados.frx":31BE
      Column(13)      =   "frmChamados.frx":33CE
      Column(14)      =   "frmChamados.frx":353A
      Column(15)      =   "frmChamados.frx":367E
      FormatStylesCount=   5
      FormatStyle(1)  =   "frmChamados.frx":37FE
      FormatStyle(2)  =   "frmChamados.frx":392A
      FormatStyle(3)  =   "frmChamados.frx":39DA
      FormatStyle(4)  =   "frmChamados.frx":3A8E
      FormatStyle(5)  =   "frmChamados.frx":3B66
      ImageCount      =   0
      PrinterProperties=   "frmChamados.frx":3C1E
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
      Left            =   2880
      TabIndex        =   8
      Top             =   255
      Width           =   555
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      TabIndex        =   5
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "frmChamados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rs As ADODB.Recordset
Private strSQL As String
Private strPrevisaoSistemas As String
Private strSituacao As String
Private strDataSituacao As String
Private strDataInicio As String
Private strDataBaixa As String
Private strTempoParada As String
Private strRelatorio As String

Private Sub cboStatus_Click()
    
    Select Case cboStatus.Text
        Case Is = "Em Aberto"
            If gblnSupervisor = True Then
                cmdAtender.Caption = "&Atender(S)"
                cmdAtender.Enabled = True
            Else
                cmdAtender.Caption = "&Atender"
                cmdAtender.Enabled = True
            End If
            cmdDevolver.Enabled = True
            cmdBaixarOS.Enabled = False
            cmdImprimir.Enabled = True
            cmdBaixarOS.Enabled = False
            cmdSituacao.Visible = False
            chkComentario.Caption = "Visualizar Comentário"
            chkComentario.Value = 0
            chkComentario.Enabled = False
            gexOS.ForeColor = vbBlack
            chkHistorico.Enabled = False
            chkHistorico.Value = 0
        Case Is = "Urgente"
            If gblnSupervisor = True Then
                cmdAtender.Caption = "&Atender(S)"
                cmdAtender.Enabled = True
            Else
                cmdAtender.Caption = "&Atender"
                cmdAtender.Enabled = True
            End If
            cmdDevolver.Enabled = True
            cmdBaixarOS.Enabled = False
            cmdImprimir.Enabled = True
            cmdBaixarOS.Enabled = False
            cmdSituacao.Visible = False
            chkComentario.Caption = "Visualizar Comentário"
            chkComentario.Value = 0
            chkComentario.Enabled = False
            gexOS.ForeColor = vbRed
            chkHistorico.Enabled = False
            chkHistorico.Value = 0
        Case Is = "Em Atendimento"
            cmdAtender.Enabled = False
            cmdDevolver.Enabled = False
            cmdBaixarOS.Enabled = True
            cmdImprimir.Enabled = True
            cmdBaixarOS.Enabled = True
            cmdSituacao.Visible = True
            chkComentario.Caption = "Visualizar Comentário"
            chkComentario.Value = 0
            chkComentario.Enabled = False
            gexOS.ForeColor = vbBlack
            chkHistorico.Enabled = False
            chkHistorico.Value = 0
        Case Is = "Aguardando Aceite"
            cmdAtender.Enabled = False
            cmdDevolver.Enabled = False
            cmdBaixarOS.Enabled = True
            cmdImprimir.Enabled = False
            cmdBaixarOS.Enabled = False
            cmdSituacao.Visible = False
            chkComentario.Caption = "Visualizar Comentário"
            chkComentario.Value = 0
            chkComentario.Enabled = False
            gexOS.ForeColor = vbBlack
            chkHistorico.Enabled = False
            chkHistorico.Value = 0
        Case Is = "Finalizada"
            cmdAtender.Enabled = False
            cmdDevolver.Enabled = False
            cmdBaixarOS.Enabled = True
            cmdImprimir.Enabled = True
            cmdBaixarOS.Enabled = False
            cmdSituacao.Visible = False
            chkComentario.Caption = "Visualizar Comentário"
            chkComentario.Value = 0
            chkComentario.Enabled = True
            gexOS.ForeColor = vbBlue
            chkHistorico.Enabled = True
            chkHistorico.Value = 0
        Case Is = "Cancelada"
            cmdAtender.Enabled = False
            cmdDevolver.Enabled = False
            cmdBaixarOS.Enabled = True
            cmdImprimir.Enabled = False
            cmdBaixarOS.Enabled = False
            cmdSituacao.Visible = False
            chkComentario.Caption = "Visualizar Motivo"
            chkComentario.Value = 0
            chkComentario.Enabled = True
            gexOS.ForeColor = vbBlack
            chkHistorico.Enabled = False
            chkHistorico.Value = 0
        Case Is = "Não Validada"
            cmdAtender.Enabled = False
            cmdDevolver.Enabled = False
            cmdBaixarOS.Enabled = True
            cmdImprimir.Enabled = True
            cmdBaixarOS.Enabled = True
            cmdSituacao.Visible = False
            chkComentario.Caption = "Visualizar Motivo"
            chkComentario.Value = 0
            chkComentario.Enabled = True
            gexOS.ForeColor = vbRed
            chkHistorico.Enabled = True
            chkHistorico.Value = 0
    End Select
        
    Call suListarChamados(cboStatus.Text)
End Sub

Private Sub chk1_Click()
    If cboStatus.Text <> "" Then
        Call suListarChamados(cboStatus.Text)
    Else
        MsgBox "Selecione um Status!", vbExclamation + vbOKOnly, "Suporte Manutenção"
    End If
End Sub

Private Sub chkComentario_Click()
    If chkComentario.Value = 1 Then
        chkHistorico.Value = 0
    End If
End Sub

Private Sub chkHistorico_Click()
    If chkHistorico.Value = 1 Then
        chkComentario.Value = 0
    End If
End Sub

Private Sub cmdAtender_Click()
Dim i As Long

    gintOSID = 0
    strPrevisaoSistemas = ""
    
    If gexOS.RowCount = 0 Then
        MsgBox "Não há nenhum chamado listado!", vbOKOnly + vbExclamation, "Suporte Manutenção"
        Exit Sub
    End If
    
    For i = 1 To gexOS.RowCount
        If gexOS.RowSelected(i) = True Then
            gintOSID = gexOS.Value(1)
            
            If MsgBox("Tem certeza que deseja atender a OS " & Format(gintOSID, "0000") & "?", vbYesNo + vbQuestion, "Suporte Manutenção") = vbYes Then
                
                strPrevisaoSistemas = InputBox("Digite a data de previsão para conclusão do serviço!", "Suporte Manutenção", Format(Now, "dd/MM/yy hh:mm"))
                                                
                If Len(strPrevisaoSistemas) < 14 Then
                    MsgBox "Data/ Hora inválida!" & vbCrLf & "O formato deve ser dd/mm/aa hh:mm", vbOKOnly + vbExclamation, "Suporte Manutenção"
                    Exit Sub
                End If
                
                If IsDate(strPrevisaoSistemas) = False Then
                    MsgBox "Data inválida!", vbOKOnly + vbExclamation, "Suporte Manutenção"
                    Exit Sub
                End If
                
                If gexOS.Value(8) <> "" And gexOS.Value(8) <> gstrNome Then
                    If gblnSupervisor = False Then
                        If MsgBox("A OS " & gintOSID & " foi encaminhada para o funcionário " & gexOS.Value(8) & vbCrLf & "Deseja atender assim mesmo?!", vbExclamation + vbYesNo, "Suporte Manutenção") = vbYes Then
                            If fnCadastrarAtendente(gintOSID, gstrNome, CDate(strPrevisaoSistemas)) = True Then
                                Call suListarChamados(cboStatus.Text)
                            End If
                        Else
                            Exit Sub
                        End If
                    Else
                        If fnCadastrarSupervisor(gintOSID, gstrNome, CDate(strPrevisaoSistemas)) = True Then
                            Call suListarChamados(cboStatus.Text)
                        End If
                    End If
                Else
                    If fnCadastrarPrevisao(gintOSID, CDate(strPrevisaoSistemas)) = True Then
                        Call suListarChamados(cboStatus.Text)
                    End If
                End If
            Else
                Exit Sub
            End If
        End If
    Next
    
End Sub

Private Function fnCadastrarSupervisor(ByVal vOSID As Long, ByVal vSupervisor As String, ByVal vPrevisaoSistemas As Variant) As Boolean
On Error GoTo Erro
    
    fnCadastrarSupervisor = False
        
    strSQL = "UPDATE tb_OS SET Supervisor = '" & vSupervisor & "' , PrevisaoSistemas='" & vPrevisaoSistemas & "', Status = 1,DataAtendimentoAtual='" & Now & "' WHERE OSID = " & vOSID & ""
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenKeyset, adLockOptimistic
    
    fnCadastrarSupervisor = True
    Set rs = Nothing
    Exit Function
    
Erro:
    MsgBox "Erro: " & Err.Description, vbOKOnly + vbCritical, "Suporte Manutenção"
    Set rs = Nothing
End Function


Private Function fnCadastrarPrevisao(ByVal vOSID As Long, ByVal vPrevisaoSistemas As Variant) As Boolean
On Error GoTo Erro
    
    fnCadastrarPrevisao = False
        
    strSQL = "UPDATE tb_OS SET PrevisaoSistemas='" & Format(vPrevisaoSistemas, "dd/MM/yyyy HH:mm") & "', Status = 1,DataAtendimentoAtual = '" & Format(Now, "dd/MM/yyyy HH:mm") & "' WHERE OSID = " & vOSID & ""
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenKeyset, adLockOptimistic
    
    fnCadastrarPrevisao = True
    Set rs = Nothing
    Exit Function
    
Erro:
    MsgBox "Erro: " & Err.Description, vbOKOnly + vbCritical, "Suporte Manutenção"
    Set rs = Nothing
End Function

Private Function fnCadastrarAtendente(ByVal vOSID As Long, ByVal vAtendente As String, ByVal vPrevisaoSistemas As Variant) As Boolean
On Error GoTo Erro
    
    fnCadastrarAtendente = False
        
    strSQL = "UPDATE tb_OS SET Atendente = '" & vAtendente & "',PrevisaoSistemas='" & Format(vPrevisaoSistemas, "dd/MM/yyyy HH:mm") & "', Status = 1,DataAtendimentoAtual='" & Format(Now, "dd/MM/yyyy HH:mm") & "' WHERE OSID = " & vOSID & ""
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenKeyset, adLockOptimistic
    
    fnCadastrarAtendente = True
    Set rs = Nothing
    Exit Function
    
Erro:
    MsgBox "Erro: " & Err.Description, vbOKOnly + vbCritical, "Suporte Manutenção"
    Set rs = Nothing
End Function

Private Sub cmdBaixarOS_Click()
Dim i As Long

    gintOSID = 0
    
    If gexOS.RowCount = 0 Then
        MsgBox "Não há nenhum chamado listado!", vbOKOnly + vbExclamation, "Suporte Manutenção"
        Exit Sub
    End If
    
    For i = 1 To gexOS.RowCount
        If gexOS.RowSelected(i) = True Then
            gintOSID = gexOS.Value(1)
            Call frmReporteTecnico.Show(vbModal)
        End If
    Next
    
    Call cboStatus_Click
End Sub

Private Sub cmdDevolver_Click()
Dim i As Long

    gintOSID = 0
    
    If gexOS.RowCount = 0 Then
        MsgBox "Não há nenhum chamado listado!", vbOKOnly + vbExclamation, "Suporte Manutenção"
        Exit Sub
    End If
    
    For i = 1 To gexOS.RowCount
        If gexOS.RowSelected(i) = True Then
            gintOSID = gexOS.Value(1)
            With frmReporteTecnico
                .Caption = "Suporte Manutenção - Devolver OS"
                .lblDataBaixa.Visible = False
                .mskDataBaixa.Visible = False
                .cmdBaixar.Visible = False
                .cmdValidar.Visible = True
            End With
            Call frmReporteTecnico.Show(vbModal)
        End If
    Next
    
End Sub

Private Sub cmdFechar_Click()
    Call Unload(Me)
End Sub

Private Sub cmdImprimir_Click()
Dim i As Long
    
    gintOSID = 0
    gintStatusFinalizada = 0
    
    If gexOS.RowCount = 0 Then
        MsgBox "Não há nenhum chamado listado!", vbOKOnly + vbExclamation, "Suporte Manutenção"
        Exit Sub
    End If
    
    If cboStatus.Text = "Finalizada" Then
        gintStatusFinalizada = 1
    End If
    
    For i = 1 To gexOS.RowCount
        If gexOS.RowSelected(i) = True Then
            gintOSID = gexOS.Value(1)
            Call acrOS.Show(vbModal)
        End If
    Next

End Sub

Private Sub cmdSituacao_Click()
Dim i As Long

    gintOSID = 0
    gstrcboStatus = ""

    If gexOS.RowCount = 0 Then
        MsgBox "Não há nenhuma OS na lista!", vbOKOnly + vbExclamation, "Suporte Técnico"
        Exit Sub
    End If
    
    For i = 1 To gexOS.RowCount
        If gexOS.RowSelected(i) = True Then
            gintOSID = gexOS.Value(1)
            gstrcboStatus = cboStatus.Text
            With frmSituacao
                .txtOSID.Text = Format(gintOSID, "0000")
                .txtOSID.Enabled = False
                .txtDataAtual.Enabled = False
                Call suPesquisarSituacao(gintOSID)
                .txtComentario.Text = strSituacao
                .txtDataAtual.Text = strDataSituacao
                .mskDataInicio.Mask = ""
                .mskDataBaixa.Mask = ""
                .mskTempo.Mask = ""
                .mskDataInicio.Text = strDataInicio
                .mskDataBaixa.Text = strDataBaixa
                .mskTempo.Text = strTempoParada
                If Len(Trim(.txtComentario.Text)) > 0 Then
                    .cmdAtualizar.Caption = "&Novo"
                End If
                If .cmdAtualizar.Caption = "&Adicionar" Or .cmdAtualizar.Caption = "&Cadastrar" Then
                    .mskDataInicio.Text = Format(Now, "dd/MM/yy HH:mm")
                    .mskDataBaixa.Text = Format(Now, "dd/MM/yy HH:mm")
                    .mskTempo.Mask = "##:##:##"
                End If
            End With
            Call frmSituacao.Show(vbModal)
            Call cboStatus_Click
        End If
    Next
End Sub

Private Sub suPesquisarSituacao(ByVal vOSID As Long)
    strSQL = "SELECT Situacao,DataSituacao,DataInicio,DataBaixa,TempoParada,TempoParadaSegundos FROM vw_Chamados WHERE OSID = " & vOSID
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    If rs.EOF = False Then
        strSituacao = rs!Situacao & ""
        strDataSituacao = Format(rs!DataSituacao, "dd/MM/yyyy") & ""
        strDataInicio = Format(rs!DataInicio, "dd/MM/yy HH:mm") & ""
        strDataBaixa = Format(rs!DataBaixa, "dd/MM/yy HH:mm") & ""
        
        Dim tempoParadaEmHoras As String
        
        If IsNull(rs!TempoParadaSegundos) Then
            tempoParadaEmHoras = ""
        Else
            tempoParadaEmHoras = fnTempoParadaEmHoras(rs!TempoParadaSegundos)
        End If
        
        strTempoParada = tempoParadaEmHoras
    End If
    rs.Close
    Set rs = Nothing
End Sub

Private Sub Form_Load()
    Call suListarStatus
End Sub

Private Sub cmdPesquisar_Click()
'    If Len(Trim(cboStatus.Text)) = 0 Then
'        MsgBox "Selecione um status!", vbOKOnly + vbExclamation, "Suporte Manutenção"
'        cboStatus.SetFocus
'        Exit Sub
'    End If
    
    If Len(Trim(txtOSID.Text)) = 0 Then
        MsgBox "Digite o Nº da OS!", vbOKOnly + vbExclamation, "Suporte Manutenção"
        txtOSID.SetFocus
        Exit Sub
    Else
'        If IsNumeric(Trim(txtOSID.Text)) = True Then
'            If fnPesquisarOSID(Trim(txtOSID.Text), cboStatus.Text) = False Then
'                MsgBox "OS não localizada!", vbOKOnly + vbExclamation, "Suporte Manutenção"
'                Exit Sub
'            End If
'        Else
            If IsNumeric(Trim(txtOSID.Text)) = True Then
                If fnPesquisarOSID(Trim(txtOSID.Text), cboStatus.Text) = False Then
                    MsgBox "OS não localizada!", vbOKOnly + vbExclamation, "Suporte Manutenção"
                    Exit Sub
                End If
            Else
    
                MsgBox "Digite um valor numérico!", vbOKOnly + vbExclamation, "Suporte Manutenção"
                txtOSID.Text = ""
                txtOSID.SetFocus
            End If
    End If
End Sub

Private Function fnPesquisarOSID(ByVal vOSID As Long, ByVal vStatus As String) As Boolean
    fnPesquisarOSID = False
    
    Select Case vStatus
        Case Is = "Em Aberto"
            strSQL = "SELECT * FROM vw_Chamados WHERE OSID = " & vOSID & " AND Status = 0 AND Prioridade = 0 ORDER BY OSID"
        Case Is = "Urgente"
            strSQL = "SELECT * FROM vw_Chamados WHERE OSID = " & vOSID & " AND Status = 0 AND Prioridade = 1 ORDER BY OSID"
        Case Is = "Em Atendimento"
            strSQL = "SELECT * FROM vw_Chamados WHERE OSID = " & vOSID & " AND Status = 1 ORDER BY OSID"
        Case Is = "Aguardando Aceite"
            strSQL = "SELECT * FROM vw_Chamados WHERE OSID = " & vOSID & " AND Status = 2 ORDER BY OSID"
        Case Is = "Finalizada"
            strSQL = "SELECT * FROM vw_Chamados WHERE OSID = " & vOSID & " AND Status = 3 ORDER BY OSID"
        Case Is = "Cancelada"
            strSQL = "SELECT * FROM vw_Chamados WHERE OSID = " & vOSID & " AND Status = 4 ORDER BY OSID"
        Case Is = "Não Validada"
            strSQL = "SELECT * FROM vw_Chamados WHERE OSID = " & vOSID & " AND Status = 6 ORDER BY OSID"
        Case Else
            strSQL = "SELECT * FROM vw_Chamados WHERE OSID = " & vOSID & " ORDER BY OSID"
    End Select
    
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenKeyset, adLockOptimistic
    
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

Private Sub suListarChamados(ByVal vStatus As String)
    
    If chk1.Value = 0 Then
        Select Case vStatus
            Case Is = "Em Aberto"
                strSQL = "SELECT * FROM vw_Chamados WHERE Status = 0 AND Prioridade = 0 ORDER BY OSID"
            Case Is = "Urgente"
                strSQL = "SELECT * FROM vw_Chamados WHERE Status = 0 AND Prioridade = 1 ORDER BY OSID"
            Case Is = "Em Atendimento"
                strSQL = "SELECT * FROM vw_Chamados WHERE Status = 1 ORDER BY OSID"
            Case Is = "Aguardando Aceite"
                strSQL = "SELECT * FROM vw_Chamados WHERE Status = 2 ORDER BY OSID"
            Case Is = "Finalizada"
                strSQL = "SELECT * FROM vw_Chamados WHERE Status = 3 ORDER BY OSID"
            Case Is = "Cancelada"
                strSQL = "SELECT * FROM vw_Chamados WHERE Status = 4 ORDER BY OSID"
            Case Is = "Não Validada"
                strSQL = "SELECT * FROM vw_Chamados WHERE Status = 6 ORDER BY OSID"
        End Select
    Else
        Select Case vStatus
            Case Is = "Em Aberto"
                strSQL = "SELECT * FROM vw_Chamados WHERE Atendente = '" & gstrNome & "' AND Status = 0 AND Prioridade = 0 ORDER BY OSID"
            Case Is = "Urgente"
                strSQL = "SELECT * FROM vw_Chamados WHERE Atendente = '" & gstrNome & "' AND Status = 0 AND Prioridade = 1 ORDER BY OSID"
            Case Is = "Em Atendimento"
                strSQL = "SELECT * FROM vw_Chamados WHERE Atendente = '" & gstrNome & "' AND Status = 1 ORDER BY OSID"
            Case Is = "Aguardando Aceite"
                strSQL = "SELECT * FROM vw_Chamados WHERE Atendente = '" & gstrNome & "' AND Status = 2 ORDER BY OSID"
            Case Is = "Finalizada"
                strSQL = "SELECT * FROM vw_Chamados WHERE Atendente = '" & gstrNome & "' AND Status = 3 ORDER BY OSID"
            Case Is = "Cancelada"
                strSQL = "SELECT * FROM vw_Chamados WHERE Atendente = '" & gstrNome & "' AND Status = 4 ORDER BY OSID"
            Case Is = "Não Validada"
                strSQL = "SELECT * FROM vw_Chamados WHERE Atendente = '" & gstrNome & "' AND Status = 6 ORDER BY OSID"
        End Select
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

Private Sub suListarStatus()
    cboStatus.Clear
    cboStatus.AddItem "Em Aberto"
    cboStatus.AddItem "Urgente"
    cboStatus.AddItem "Em Atendimento"
    cboStatus.AddItem "Aguardando Aceite"
    cboStatus.AddItem "Finalizada"
    cboStatus.AddItem "Cancelada"
    cboStatus.AddItem "Não Validada"
End Sub

Private Sub gexOS_DblClick()
Dim i As Long
    
    gintOSID = 0
    gstrTela = ""
    
    If gexOS.RowCount = 0 Then
        MsgBox "Não há nenhum chamado listado!", vbOKOnly + vbExclamation, "Suporte Manutenção"
        Exit Sub
    End If
    
    If chkHistorico.Value = 1 Then
        For i = 1 To gexOS.RowCount
            If gexOS.RowSelected(i) = True Then
                gintOSID = gexOS.Value(1)
                gstrTela = "frmChamados"
                Call suMostrarOcorrencia(gintOSID)
                Call frmOcorrencias.Show(vbModal)
                gstrTela = ""
            End If
        Next
        Exit Sub
    End If
    
    If chkComentario.Value = 0 Then
        For i = 1 To gexOS.RowCount
            If gexOS.RowSelected(i) = True Then
                gintOSID = gexOS.Value(1)
                gstrTela = "frmChamados"
                Call suMostrarOS(gintOSID, cboStatus.Text)
                Call frmSuporteSistemas.Show(vbModal)
                gstrTela = ""
            End If
        Next
    Else
        For i = 1 To gexOS.RowCount
            If gexOS.RowSelected(i) = True Then
                gintOSID = gexOS.Value(1)
                gstrTela = "frmChamados"
                If cboStatus.Text = "Finalizada" Then
                    Call suMostrarComentario(gintOSID)
                    Call frmAceite.Show(vbModal)
                ElseIf cboStatus.Text = "Cancelada" Then
                    Call suMostrarMotivo(gintOSID)
                    Call frmCancelarOS.Show(vbModal)
                ElseIf cboStatus.Text = "Não Validada" Then
                    Call suMostrarNaoValidada(gintOSID)
                    Call frmAceite.Show(vbModal)
                End If
                gstrTela = ""
            End If
        Next
    End If
End Sub

Private Sub suMostrarOcorrencia(ByVal vOSID As Long)
    strSQL = "SELECT * FROM tb_Ocorrencias WHERE OSID = " & vOSID
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        With frmOcorrencias
            .gexOcorrencias.HoldFields
            Set .gexOcorrencias.ADORecordset = rs
        End With
    End If
    
    Set rs = Nothing
End Sub


Private Sub suMostrarNaoValidada(ByVal vOSID As Long)
    strSQL = "SELECT OSID,MotivoOSNaoValidada FROM vw_Chamados WHERE OSID = " & vOSID
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        With frmAceite
            .txtOSID.Text = Format(rs!OSID, "0000")
            .txtComentario.Text = rs!MotivoOSNaoValidada & ""
            .cmdAceite.Enabled = False
            .cmdNaoValidar.Enabled = False
        End With
    End If
    
    rs.Close
    Set rs = Nothing
End Sub

Private Sub suMostrarMotivo(ByVal vOSID As Long)
    strSQL = "SELECT OSID,Comentario FROM vw_Chamados WHERE OSID = " & vOSID
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        With frmCancelarOS
            .txtOSID.Text = Format(rs!OSID, "0000")
            .txtMotivo.Text = rs!Comentario & ""
            .cmdCancelar.Enabled = False
        End With
    End If
    
    rs.Close
    Set rs = Nothing
End Sub

Private Sub suMostrarComentario(ByVal vOSID As Long)
    strSQL = "SELECT OSID,Comentario FROM vw_Chamados WHERE OSID = " & vOSID
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        With frmAceite
            .txtOSID.Text = Format(rs!OSID, "0000")
            .txtComentario.Text = rs!Comentario & ""
            .cmdAceite.Enabled = False
            .cmdNaoValidar.Enabled = False
        End With
    End If
    
    rs.Close
    Set rs = Nothing
End Sub

Private Sub suMostrarOS(ByVal vOSID As Long, ByVal vStatus As String)
    strSQL = "SELECT * FROM vw_Chamados WHERE OSID = " & vOSID
    'strSQL = "SELECT REPLACE(LEFT(Especificacao,CHARINDEX('-',Especificacao)),'-','') AS Especificacao,Divisao,Tipo,DescricaoServico,Prioridade,Prazo,Previsao,Atendente,ReporteTecnico FROM vw_Chamados WHERE OSID = " & vOSID
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        If gblnSupervisor = True And (vStatus = "Em Aberto" Or vStatus = "Em Atendimento" Or vStatus = "Urgente") Then
            With frmSuporteSistemas
                .cboDivisao.Enabled = True
                .cboDivisao.Text = rs!Divisao
                .cboTipo.Enabled = True
                .cboTipo.Text = rs!Tipo
                .cboLinha.Enabled = True
                .cboLinha.Text = rs!Linha
                .cboEspecificacao.Enabled = True
                .cboEspecificacao.Text = rs!Especificacao
                .txtObservacao.Enabled = False
                .txtObservacao.Text = rs!DescricaoServico
                '.txtReporteTecnico.Text = rs!ReporteTecnico & ""
                .chkPrioridade.Enabled = False
                .chkPrioridade.Value = IIf(rs!Prioridade = True, 1, 0)
                .txtPrazo.Enabled = False
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
                If gblnEncaminharOS = False Then
                    If rs!Atendente <> "" Then
                        .cboAtendente.Enabled = False
                        .cboAtendente.Text = rs!Atendente
                    Else
                        .cboAtendente.Enabled = False
                    End If
                    .cmdCadastrar.Enabled = False
                Else
                    If rs!Atendente <> "" Then
                        .cboAtendente.Text = rs!Atendente
                    End If
                    If cboStatus.Text = "Em Aberto" Or cboStatus.Text = "Urgente" Then
                        .cboAtendente.Enabled = True
                        .cmdCadastrar.Caption = "&Alterar"
                        .cmdCadastrar.Enabled = True
                    Else
                        .cboAtendente.Enabled = False
                        .cmdCadastrar.Caption = "&Cadastrar"
                        .cmdCadastrar.Enabled = False
                    End If
                End If
                .lblQtdeCaracteres.Enabled = False
                '.cmdCadastrar.Enabled = False
                .cmdCancelar.Enabled = False
            End With
        End If
    End If
    
    rs.Close
    Set rs = Nothing
End Sub

Private Sub txtOSID_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Len(Trim(txtOSID.Text)) > 0 Then
            Call cmdPesquisar_Click
        End If
    End If
End Sub
