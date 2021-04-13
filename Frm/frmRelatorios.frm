VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRelatorios 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Suporte Manutenção - Relatórios"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10110
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRelatorios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   10110
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboLinha 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4800
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   375
      Width           =   2175
   End
   Begin VB.CheckBox chkLinha 
      Appearance      =   0  'Flat
      Caption         =   "Limpar"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6000
      TabIndex        =   25
      Top             =   120
      Width           =   975
   End
   Begin VB.CheckBox chkStatus 
      Appearance      =   0  'Flat
      Caption         =   "Limpar"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6000
      TabIndex        =   24
      Top             =   750
      Width           =   975
   End
   Begin VB.CheckBox chkUsuario 
      Appearance      =   0  'Flat
      Caption         =   "Limpar"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1440
      TabIndex        =   23
      Top             =   750
      Width           =   975
   End
   Begin VB.CheckBox chkEspec 
      Appearance      =   0  'Flat
      Caption         =   "Limpar"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8880
      TabIndex        =   22
      Top             =   105
      Width           =   975
   End
   Begin VB.CheckBox chkCaract 
      Appearance      =   0  'Flat
      Caption         =   "Limpar"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3720
      TabIndex        =   21
      Top             =   105
      Width           =   975
   End
   Begin VB.CheckBox chkTipo 
      Appearance      =   0  'Flat
      Caption         =   "Limpar"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1440
      TabIndex        =   20
      Top             =   105
      Width           =   975
   End
   Begin VB.ComboBox cboUsuario 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1020
      Width           =   4575
   End
   Begin VB.TextBox txtDataCadAte 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8760
      TabIndex        =   7
      Top             =   1035
      Width           =   1215
   End
   Begin VB.TextBox txtDataCadDe 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7080
      TabIndex        =   6
      Top             =   1035
      Width           =   1215
   End
   Begin VB.ComboBox cboTipo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2520
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   2175
   End
   Begin VB.ComboBox cboEspecificacao 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7080
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   360
      Width           =   2895
   End
   Begin VB.ComboBox cboDivisao 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
   Begin MSComDlg.CommonDialog cdlSalvarArquivo 
      Left            =   8400
      Top             =   7680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Pasta de Trabalho do Microsoft Office Excel (*.xls) | *.xls"
   End
   Begin MSComctlLib.ProgressBar pgbProgresso 
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   5400
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.CommandButton cmdExportar 
      Caption         =   "&Exportar XLS"
      Height          =   375
      Left            =   1800
      TabIndex        =   12
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton cmdGerar 
      Caption         =   "&Gerar"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "Fechar"
      Height          =   375
      Left            =   8520
      TabIndex        =   11
      Top             =   5880
      Width           =   1455
   End
   Begin VB.ComboBox cboStatus 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4800
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1020
      Width           =   2175
   End
   Begin FPSpreadADO.fpSpread fpsRelatorio 
      Height          =   3855
      Left            =   120
      TabIndex        =   9
      Top             =   1440
      Width           =   9855
      _Version        =   458752
      _ExtentX        =   17383
      _ExtentY        =   6800
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxRows         =   65000
      OperationMode   =   1
      SpreadDesigner  =   "frmRelatorios.frx":030A
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Linha"
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
      Left            =   4800
      TabIndex        =   26
      Top             =   120
      Width           =   525
   End
   Begin VB.Label lblUsuario 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Usuário"
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
      TabIndex        =   19
      Top             =   780
      Width           =   750
   End
   Begin VB.Label lblAte 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "até"
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
      Left            =   8385
      TabIndex        =   18
      Top             =   1080
      Width           =   315
   End
   Begin VB.Label lblDataCadastro 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Data Cadastro"
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
      Left            =   7080
      TabIndex        =   17
      Top             =   795
      Width           =   1380
   End
   Begin VB.Label lblTipo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Caract."
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
      Left            =   2520
      TabIndex        =   16
      Top             =   120
      Width           =   690
   End
   Begin VB.Label lblEspecificacao 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Equipamento"
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
      Left            =   7080
      TabIndex        =   15
      Top             =   120
      Width           =   1275
   End
   Begin VB.Label lblDivisao 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo"
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
      TabIndex        =   14
      Top             =   120
      Width           =   420
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
      Left            =   4800
      TabIndex        =   10
      Top             =   780
      Width           =   615
   End
End
Attribute VB_Name = "frmRelatorios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rs As ADODB.Recordset
Private strSQL As String
Private strDataSituacao As String
Private strSituacao As String

Private Sub cboDivisao_Click()
    With fpsRelatorio
        .Reset
    End With
    
    pgbProgresso.Value = 0
'    cboTipo.Clear
'    cboLinha.Clear
'    cboEspecificacao.Clear
    
    If chkCaract.Value = 1 Then
        Exit Sub
    End If
    
'    If Len(Trim(cboDivisao.Text)) > 0 Then
'        Call suListarTipo(cboDivisao.Text)
'    End If
End Sub

Private Sub cboEspecificacao_Click()
    pgbProgresso.Value = 0
    
    With fpsRelatorio
        .Reset
    End With
End Sub

Private Sub cboLinha_Click()
    With fpsRelatorio
        .Reset
    End With
    
    pgbProgresso.Value = 0
    cboEspecificacao.Clear
    
    If chkEspec.Value = 1 Then
        Exit Sub
    End If
    
    If Len(Trim(cboLinha.Text)) > 0 Then
        Call suListarEspecificacao(cboLinha.Text)
    End If

End Sub

Private Function fnLinhaID(ByVal vLinha As String) As Integer
Dim rs1 As New ADODB.Recordset

    strSQL = "SELECT * FROM tb_Linhas WHERE Linha = '" & vLinha & "'"
    rs1.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If rs1.EOF = False Then
        fnLinhaID = rs1!LinhaID
    End If
    
    rs1.Close
    Set rs1 = Nothing
End Function

Private Sub cboStatus_Click()
    With fpsRelatorio
        .Reset
    End With
    
    pgbProgresso.Value = 0
End Sub

Private Sub cboTipo_Click()
    With fpsRelatorio
        .Reset
    End With
    
    pgbProgresso.Value = 0
    'cboLinha.Clear
    
    If chkLinha.Value = 1 Then
        Exit Sub
    End If
    
'    If Len(Trim(cboTipo.Text)) > 0 Then
'        Call suListarLinha(cboDivisao.Text, cboTipo.Text)
'    End If
End Sub

Private Sub cboUsuario_Click()
    
    pgbProgresso.Value = 0
    
    With fpsRelatorio
        .Reset
    End With
End Sub

Private Sub chkCaract_Click()
    If chkCaract.Value = 1 Then
        cboTipo.Clear
    Else
        Call suListarTipo
    End If
End Sub

Private Sub chkEspec_Click()
    If chkEspec.Value = 1 Then
        cboEspecificacao.Clear
    Else
        Call suListarEspecificacao(cboLinha.Text)
    End If
End Sub

Private Sub chkLinha_Click()
    If chkLinha.Value = 1 Then
        cboLinha.Clear
        cboEspecificacao.Clear
    Else
        Call suListarLinha
    End If
End Sub

Private Sub chkStatus_Click()
    If chkStatus.Value = 1 Then
        cboStatus.Clear
    Else
        Call suListarStatus
    End If
End Sub

Private Sub chkTipo_Click()
    If chkTipo.Value = 1 Then
        cboDivisao.Clear
    Else
        Call suListarDivisoes
    End If
End Sub

Private Sub chkUsuario_Click()
    If chkUsuario.Value = 1 Then
        cboUsuario.Clear
    Else
        Call suListarUsuarios
    End If
End Sub

Private Sub cmdExportar_Click()
Dim blnExportar As Boolean
    
    If fpsRelatorio.SheetName = "Relatório OS" Then
        
        cdlSalvarArquivo.FileName = ""
        
        Call cdlSalvarArquivo.ShowSave
        
        If cdlSalvarArquivo.FileName <> "" Then
            fpsRelatorio.Protect = False
            blnExportar = fpsRelatorio.ExportToExcel(cdlSalvarArquivo.FileName, fpsRelatorio.SheetName, "")
            fpsRelatorio.Protect = True
        Else
            Exit Sub
        End If
        
        If blnExportar = True Then
            MsgBox "Arquivo exportado com sucesso!", vbOKOnly + vbInformation, "Suporte Manutenção"
        Else
            MsgBox "Erro ao exportar o arquivo!", vbOKOnly + vbCritical, "Suporte Manutenção"
        End If
    Else
        MsgBox "Nenhum relatório foi gerado!", vbOKOnly + vbExclamation, "Suporte Manutenção"
    End If
    
End Sub

Private Sub cmdFechar_Click()
    Call Unload(Me)
End Sub

Private Sub cmdGerar_Click()
    
    If Len(cboDivisao.Text) = 0 And Len(cboTipo.Text) = 0 And Len(cboLinha.Text) = 0 And Len(cboEspecificacao.Text) = 0 And Len(cboUsuario.Text) = 0 And Len(cboStatus.Text) = 0 And Len(txtDataCadDe.Text) = 0 And Len(txtDataCadAte.Text) = 0 Then
        MsgBox "Nenhum critério foi informado!", vbOKOnly + vbExclamation, "Suporte Manutenção"
        Exit Sub
    End If
    
    If IsDate(txtDataCadDe.Text) = True And IsDate(txtDataCadAte.Text) = True Then
        If CDate(txtDataCadDe.Text) > CDate(txtDataCadAte.Text) Then
            MsgBox "DATA INICIAL não pode ser maior que a DATA FINAL!", vbOKOnly + vbExclamation, "Suporte Manutenção"
            txtDataCadDe.Text = ""
            txtDataCadAte.Text = ""
            txtDataCadDe.SetFocus
            Exit Sub
        End If
    End If
    
    If Len(txtDataCadDe.Text) > 0 Then
        If IsDate(txtDataCadDe.Text) = False Then
            MsgBox "Data inválida!", vbOKOnly + vbExclamation, "Suporte Manutenção"
            txtDataCadDe.Text = ""
            txtDataCadDe.SetFocus
            Exit Sub
        End If
    End If
    
    If Len(txtDataCadAte.Text) > 0 Then
        If IsDate(txtDataCadAte.Text) = False Then
            MsgBox "Data inválida!", vbOKOnly + vbExclamation, "Suporte Manutenção"
            txtDataCadAte.Text = ""
            txtDataCadAte.SetFocus
            Exit Sub
        End If
    End If
    
    Call suGerarRelatorioDetalhado(cboDivisao.Text, cboTipo.Text, cboEspecificacao.Text, Left(cboUsuario.Text, 4), cboStatus.Text, txtDataCadDe.Text, txtDataCadAte.Text, cboLinha.Text)
End Sub

Private Sub Form_Load()
    Call suListarDivisoes
    Call suListarTipo
    Call suListarLinha
    'Call suListarEspecificacao
    Call suListarStatus
    Call suListarUsuarios
End Sub

Private Function fnDivisaoID(ByVal vDivisao As String) As Integer
Dim rs1 As New ADODB.Recordset

    strSQL = "SELECT * FROM tb_Divisao WHERE Divisao = '" & vDivisao & "'"
    rs1.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If rs1.EOF = False Then
        fnDivisaoID = rs1!DivisaoID
    End If
    
    rs1.Close
    Set rs1 = Nothing
End Function

Private Function fnTipoID(ByVal vDivisao As String, ByVal vTipo As String) As Integer
Dim rs1 As New ADODB.Recordset

    strSQL = "SELECT * FROM tb_Tipos WHERE Tipo = '" & vTipo & "' AND DivisaoID = " & fnDivisaoID(vDivisao)
    rs1.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If rs1.EOF = False Then
        fnTipoID = rs1!TipoID
    End If
    
    rs1.Close
    Set rs1 = Nothing
End Function

Private Function fnEspecificacaoID(ByVal vDivisao As String, ByVal vTipo As String, ByVal vEspecificacao As String, ByVal vLinha As String) As Double
Dim rs1 As New ADODB.Recordset

    strSQL = "SELECT * FROM vw_Especificacoes WHERE Especificacao = '" & vEspecificacao & "' AND DivisaoID = " & fnDivisaoID(vDivisao) & " AND TipoID = " & fnTipoID(vDivisao, vTipo) & " AND LinhaID = " & fnLinhaID(vLinha)
    rs1.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If rs1.EOF = False Then
        fnEspecificacaoID = rs1!EspecificacaoID
    End If
    
    rs1.Close
    Set rs1 = Nothing
End Function

Private Sub suGerarRelatorioDetalhado(ByVal vTipoID As String, ByVal vCaractID As String, ByVal vEspecID As String, ByVal vUsuarioID As String, ByVal vStatus As String, ByVal vDataCadDe As Variant, ByVal vDataCadAte As Variant, ByVal vLinhaID As String)
On Error GoTo Erro

Dim blnWhere As Boolean
Dim intStatus As Integer
Dim blnPrioridade As Boolean
Dim strStatus As String
Dim strPrioridade As String

    blnPrioridade = False
    
    strSQL = "SELECT * FROM vw_Chamados (NOLOCK)"
    blnWhere = False
    
    If Len(vTipoID) > 0 Then
        strSQL = strSQL & IIf(blnWhere = True, " AND ", " WHERE ") & "Divisao = '" & vTipoID & "'"
        blnWhere = True
    End If
            
    If Len(vCaractID) > 0 Then
        strSQL = strSQL & IIf(blnWhere = True, " AND ", " WHERE ") & "Tipo = '" & vCaractID & "'"
        blnWhere = True
    End If

    If Len(vLinhaID) > 0 Then
        strSQL = strSQL & IIf(blnWhere = True, " AND ", " WHERE ") & "Linha = '" & vLinhaID & "'"
        blnWhere = True
    End If

    If Len(vEspecID) > 0 Then
        strSQL = strSQL & IIf(blnWhere = True, " AND ", " WHERE ") & "Especificacao = '" & vEspecID & "'"
        blnWhere = True
    End If
    
    If Len(vUsuarioID) > 0 Then
        strSQL = strSQL & IIf(blnWhere = True, " AND ", " WHERE ") & "UsuarioID = " & CInt(vUsuarioID) & ""
        blnWhere = True
    End If
    
    If Len(vStatus) > 0 Then

        Select Case vStatus
            Case Is = "Em Aberto"
                intStatus = 0
            Case Is = "Urgente"
                blnPrioridade = True
            Case Is = "Em Atendimento"
                intStatus = 1
            Case Is = "Aguardando Aceite"
                intStatus = 2
            Case Is = "Finalizada"
                intStatus = 3
            Case Is = "Cancelada"
                intStatus = 4
            Case Is = "Não Validada"
                intStatus = 6
            Case Is = "Não Executada"
                intStatus = 7
        End Select
                    
        If blnPrioridade = True Then
            strSQL = strSQL & IIf(blnWhere = True, " AND ", " WHERE ") & "Prioridade = '" & blnPrioridade & "'"
        Else
            If intStatus <> 7 Then
                strSQL = strSQL & IIf(blnWhere = True, " AND ", " WHERE ") & "Status = " & intStatus & ""
            Else
                strSQL = strSQL & IIf(blnWhere = True, " AND ", " WHERE ") & "Status IN (0,1) "
            End If
        End If
        
        blnWhere = True
    End If
    
    If Len(vDataCadDe) > 0 And Len(vDataCadAte) > 0 Then
        strSQL = strSQL & IIf(blnWhere = True, " AND ", " WHERE ") & "DataCadastro >= '" & vDataCadDe & " 00:00:00' AND DataCadastro <= '" & vDataCadAte & " 23:59:59'"
        blnWhere = True
    ElseIf Len(vDataCadDe) > 0 And Len(vDataCadAte) = 0 Then
        strSQL = strSQL & IIf(blnWhere = True, " AND ", " WHERE ") & "DataCadastro BETWEEN '" & vDataCadDe & " 00:00:00' AND '" & vDataCadDe & " 23:59:59'"
        blnWhere = True
    ElseIf Len(vDataCadDe) = 0 And Len(vDataCadAte) > 0 Then
        strSQL = strSQL & IIf(blnWhere = True, " AND ", " WHERE ") & "DataCadastro <= '" & vDataCadAte & " 23:59:59'"
        blnWhere = True
    End If
    
    strSQL = strSQL & " ORDER BY OSID"
    
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        pgbProgresso.min = 0
        pgbProgresso.Value = 0
        pgbProgresso.Max = IIf(rs.RecordCount = 0, 1, rs.RecordCount)
        Screen.MousePointer = 11
    Else
        MsgBox "Nenhuma OS foi localizada!", vbOKOnly + vbExclamation, "Suporte Manutenção"
        Exit Sub
    End If
    
    With fpsRelatorio
        .Reset
        .Row = 1
        .FontSize = 8
        .FontName = "Verdana"
        .FontBold = True
        .TypeHAlign = TypeHAlignCenter
        .MaxRows = 65000
        
        .Col = 1
        .ColWidth(1) = 15
        .Text = "SOLICITANTE"
        
        .Col = 2
        .ColWidth(2) = 15
        .Text = "DEPTO."
        
        .Col = 3
        .ColWidth(3) = 5
        .Text = "OS"
        
        .Col = 4
        .ColWidth(4) = 15
        .Text = "STATUS"
        
        .Col = 5
        .ColWidth(5) = 15
        .Text = "TIPO"
        
        .Col = 6
        .ColWidth(6) = 15
        .Text = "CARACTERÍSTICA"
        
        .Col = 7
        .ColWidth(7) = 15
        .Text = "LINHA"
        
        .Col = 8
        .ColWidth(8) = 30
        .Text = "EQUIPAMENTO"
        
        .Col = 9
        .ColWidth(9) = 30
        .Text = "OBSERVAÇÃO"
        
        .Col = 10
        .ColWidth(10) = 15
        .Text = "DATA CADASTRO"
        
        .Col = 11
        .ColWidth(11) = 15
        .Text = "NECESSIDADE"
        
        .Col = 12
        .ColWidth(12) = 15
        .Text = "PREV. MANUTENÇÃO"
        
        .Col = 13
        .ColWidth(13) = 15
        .Text = "ATENDENTE"
        
        .Col = 14
        .ColWidth(14) = 15
        .Text = "DATA INICIO"
        
        .Col = 15
        .ColWidth(15) = 15
        .Text = "DATA FINAL."
    
        .Col = 16
        .ColWidth(16) = 15
        .Text = "TEMPO PARADA MAQUINA"
    
        .Col = 17
        .ColWidth(17) = 30
        .Text = "REPORTE TÉCNICO"
    
        .Col = 18
        .ColWidth(18) = 15
        .Text = "DATA ACEITE"
                
        .Col = 19
        .ColWidth(19) = 30
        .Text = "COMENTÁRIO"
    
        .Col = 20
        .ColWidth(20) = 15
        .Text = "DATA CANCEL."

        .Col = 21
        .ColWidth(21) = 15
        .Text = "DATA OS NÃO VALIDADA"

        .Col = 22
        .ColWidth(22) = 30
        .Text = "MOTIVO OS NÃO VALIDADA"
        
        .Col = 23
        .ColWidth(23) = 15
        .Text = "DATA SITUAÇÃO"

        .Col = 24
        .ColWidth(24) = 30
        .Text = "SITUAÇÃO"

        .Col = 25
        .ColWidth(25) = 15
        .Text = "DATA ATENDIMENTO SISTEMA"

        .Col = 26
        .ColWidth(26) = 15
        .Text = "DATA BAIXA SISTEMA"
        
        .Col = 27
        .ColWidth(27) = 15
        .Text = "MÁQUINA PARADA"

        Do While Not rs.EOF
        DoEvents
            .Row = .Row + 1
            
            .Col = 1
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignLeft
            .Text = rs!Nome & ""
            
            .Col = 2
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignLeft
            .Text = rs!Departamento & ""
            
            .Col = 3
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignRight
            .Text = CStr(Format(rs!OSID, "0000")) & ""
            
            Select Case rs!Status
                Case 0
                    strStatus = "Em Aberto"
                Case 1
                    strStatus = "Em Atendimento"
                Case 2
                    strStatus = "Aguardando Aceite"
                Case 3
                    strStatus = "OS Finalizada"
                Case 4
                    strStatus = "OS Cancelada"
                Case 6
                    strStatus = "Não Validada"
            End Select
            
            .Col = 4
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignLeft
            .Text = strStatus & ""
            
            .Col = 5
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignLeft
            .Text = rs!Divisao & ""
            
            .Col = 6
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignLeft
            .Text = rs!Tipo & ""
            
            .Col = 7
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignLeft
            .Text = rs!Linha & ""
            
            .Col = 8
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignLeft
            .Text = rs!Especificacao & ""
            
            .Col = 9
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignLeft
            .Text = rs!DescricaoServico & ""
            
            .Col = 10
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignCenter
            .Text = Format(rs!DataCadastro, "dd/MM/yy HH:mm")
            
            .Col = 11
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignCenter
            .Text = Format(rs!Previsao, "dd/MM/yyyy")
                    
            .Col = 12
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignCenter
            .Text = Format(rs!PrevisaoSistemas, "dd/MM/yy HH:mm")
            
            .Col = 13
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignLeft
            .Text = rs!Atendente & ""
            
            .Col = 14
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignCenter
            .Text = Format(rs!DataInicio, "dd/MM/yy HH:mm")
        
            .Col = 15
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignCenter
            .Text = Format(rs!DataBaixa, "dd/MM/yy HH:mm")
            
            .Col = 16
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignCenter
             
            If IsNull(rs!TempoParadaSegundos) Then
                .Text = Format(rs!tempoParada, "HH:mm:ss") & ""
            Else
                .Text = fnTempoParadaEmHoras(rs!TempoParadaSegundos) & ""
            End If
            
            .Col = 17
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignLeft
            .Text = rs!ReporteTecnico & ""
            
            .Col = 18
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignCenter
            .Text = Format(rs!DataAceite, "dd/MM/yy HH:mm")
            
            .Col = 19
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignLeft
            .Text = rs!Comentario & ""
            
            .Col = 20
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignCenter
            .Text = Format(rs!DataCancelamento, "dd/MM/yy HH:mm")
            
            .Col = 21
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignCenter
            .Text = Format(rs!DataOSNaoValidada, "dd/MM/yy HH:mm")
            
            .Col = 22
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignLeft
            .Text = rs!MotivoOSNaoValidada & ""
            
            Call suLogSituacao(rs!OSID)
            
            .Col = 23
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignCenter
            '.Text = Format(rs!DataSituacao, "dd/MM/yy HH:mm")
            .Text = Format(strDataSituacao, "dd/MM/yy HH:mm")
            
            .Col = 24
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignLeft
            '.Text = rs!Situacao & ""
            .Text = strSituacao & ""
            
            .Col = 25
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignCenter
            .Text = Format(rs!DataAtendimentoAtual, "dd/MM/yy HH:mm")

            .Col = 26
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignCenter
            .Text = Format(rs!DataBaixaAtual, "dd/MM/yy HH:mm")
            
            Select Case rs!Prioridade
                Case False
                    strPrioridade = "Não"
                Case Else
                    strPrioridade = "Sim"
            End Select
            
            .Col = 27
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignLeft
            .Text = strPrioridade & ""

            pgbProgresso.Value = rs.AbsolutePosition
            rs.MoveNext
        Loop
        
        .SheetName = "Relatório OS"
        .OperationMode = OperationModeRead
        
    End With
    
    Set rs = Nothing
    MsgBox "Relatório gerado com sucesso!", vbOKOnly + vbInformation, "Suporte Manutenção"
    Screen.MousePointer = 0
    Exit Sub
    
Erro:
    MsgBox "Erro: " & Err.Description, vbOKOnly + vbCritical, "Suporte Manutenção"
    Set rs = Nothing
    Screen.MousePointer = 0
End Sub

Private Sub suLogSituacao(ByVal vOSID As Long)
Dim rs2 As New ADODB.Recordset
    
    strSituacao = ""
    strDataSituacao = ""
    
    strSQL = "SELECT TOP 1 Situacao,DataSituacao FROM dbo.tb_Log_Situacoes WHERE OSID=" & vOSID & " ORDER BY OSID,DataSituacao DESC"
    rs2.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    If rs2.EOF = False Then
        strSituacao = rs2!Situacao & ""
        strDataSituacao = rs2!DataSituacao & ""
    End If
    Set rs2 = Nothing
End Sub

Private Sub suListarUsuarios()
    strSQL = "SELECT UsuarioID,Nome FROM vw_Usuarios WHERE Inativo = 0 ORDER BY Nome"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    Do While Not rs.EOF
        cboUsuario.AddItem Format(rs!UsuarioID, "0000") & " - " & rs!Nome
        rs.MoveNext
    Loop
    
    Set rs = Nothing
End Sub

Private Sub suListarDivisoes()
    strSQL = "SELECT * FROM tb_Divisao ORDER BY DivisaoID"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    Do While Not rs.EOF
        cboDivisao.AddItem rs!Divisao
        rs.MoveNext
    Loop
    
    Set rs = Nothing
End Sub

Private Sub suListarTipo()
    strSQL = "SELECT DISTINCT(Tipo) FROM vw_Tipos ORDER BY Tipo"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    Do While Not rs.EOF
        cboTipo.AddItem rs!Tipo
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
End Sub

Private Sub suListarLinha()
    strSQL = "SELECT DISTINCT(Linha) FROM vw_Linhas (NOLOCK) ORDER BY Linha"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    Do While Not rs.EOF
        cboLinha.AddItem rs!Linha
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
End Sub

Private Sub suListarEspecificacao(ByVal vLinha As String)
    strSQL = "SELECT * FROM vw_Especificacoes WHERE LinhaID = " & fnLinhaID(vLinha) & ""
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly

    Do While Not rs.EOF
        cboEspecificacao.AddItem rs!Especificacao
        rs.MoveNext
    Loop

    rs.Close
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
    cboStatus.AddItem "Não Executada"
End Sub

Private Sub suGerarRelatorio(ByVal vStatus As String)
On Error GoTo Erro

    Select Case vStatus
        Case Is = "Em Aberto"
            strSQL = "SELECT * FROM vw_Chamados WHERE Status = 0 ORDER BY OSID"
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
    
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        pgbProgresso.min = 0
        pgbProgresso.Value = 0
        pgbProgresso.Max = IIf(rs.RecordCount = 0, 1, rs.RecordCount)
        Screen.MousePointer = 11
    Else
        MsgBox "Nenhuma OS foi localizada!", vbOKOnly + vbExclamation, "Suporte Manutenção"
        Exit Sub
    End If
    
    With fpsRelatorio
        .Reset
        .Row = 1
        .FontSize = 8
        .FontName = "Verdana"
        .FontBold = True
        .TypeHAlign = TypeHAlignCenter
        .MaxRows = 65000
        
        .Col = 1
        .ColWidth(1) = 15
        .Text = "SOLICITANTE"
        
        .Col = 2
        .ColWidth(2) = 5
        .Text = "OS"
        
        .Col = 3
        .ColWidth(3) = 15
        .Text = "TIPO"
        
        .Col = 4
        .ColWidth(4) = 15
        .Text = "CARACTERÍSTICA"
        
        .Col = 5
        .ColWidth(5) = 30
        .Text = "ESPECIFICAÇÃO"
        
        .Col = 6
        .ColWidth(6) = 30
        .Text = "OBSERVAÇÃO"
        
        .Col = 7
        .ColWidth(7) = 15
        .Text = "DATA CADASTRO"
        
        .Col = 8
        .ColWidth(8) = 15
        .Text = "NECESSIDADE"
        
        Select Case vStatus
            Case Is = "Em Atendimento"
                .Col = 9
                .ColWidth(9) = 15
                .Text = "PREV. MANUTENÇÃO"
                
                .Col = 10
                .ColWidth(10) = 15
                .Text = "ATENDENTE"
            Case Is = "Aguardando Aceite"
                .Col = 9
                .ColWidth(9) = 15
                .Text = "PREV. SISTEMAS"
                
                .Col = 10
                .ColWidth(10) = 15
                .Text = "ATENDENTE"
                
                .Col = 11
                .ColWidth(11) = 15
                .Text = "DATA FINALIZAÇÃO"
            Case Is = "Finalizada"
                .Col = 9
                .ColWidth(9) = 15
                .Text = "PREV. SISTEMAS"
                
                .Col = 10
                .ColWidth(10) = 15
                .Text = "ATENDENTE"
                
                .Col = 11
                .ColWidth(11) = 15
                .Text = "DATA FINAL."
            
                .Col = 12
                .ColWidth(12) = 15
                .Text = "DATA ACEITE"
            Case Is = "Cancelada"
                .Col = 9
                .ColWidth(9) = 15
                .Text = "DATA CANCEL."
        End Select
                
        Do While Not rs.EOF
        DoEvents
            .Row = .Row + 1
            .Col = 1
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignLeft
            .Text = rs!Nome & ""
            
            .Col = 2
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignRight
            .Text = CStr(Format(rs!OSID, "0000")) & ""
            
            .Col = 3
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignLeft
            .Text = rs!Divisao & ""
            
            .Col = 4
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignLeft
            .Text = rs!Tipo & ""
            
            .Col = 5
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignLeft
            .Text = rs!Especificacao & ""
            
            .Col = 6
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignLeft
            .Text = rs!DescricaoServico & ""
            
            .Col = 7
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignCenter
            .Text = Format(rs!DataCadastro, "dd/MM/yy HH:mm")
            
            .Col = 8
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignCenter
            .Text = Format(rs!Previsao, "dd/MM/yyyy")
                    
        Select Case vStatus
            Case Is = "Em Atendimento"
                .Col = 9
                .FontSize = 8
                .FontName = "Verdana"
                .TypeHAlign = TypeHAlignCenter
                .Text = Format(rs!PrevisaoSistemas, "dd/MM/yy HH:mm")
                
                .Col = 10
                .FontSize = 8
                .FontName = "Verdana"
                .TypeHAlign = TypeHAlignLeft
                .Text = rs!Atendente & ""
            Case Is = "Aguardando Aceite"
                .Col = 9
                .FontSize = 8
                .FontName = "Verdana"
                .TypeHAlign = TypeHAlignCenter
                .Text = Format(rs!PrevisaoSistemas, "dd/MM/yyyy")
                
                .Col = 10
                .FontSize = 8
                .FontName = "Verdana"
                .TypeHAlign = TypeHAlignLeft
                .Text = rs!Atendente & ""
            
                .Col = 11
                .FontSize = 8
                .FontName = "Verdana"
                .TypeHAlign = TypeHAlignLeft
                .Text = Format(rs!DataBaixa, "dd/MM/yy HH:mm")
            Case Is = "Finalizada"
                .Col = 9
                .FontSize = 8
                .FontName = "Verdana"
                .TypeHAlign = TypeHAlignCenter
                .Text = Format(rs!PrevisaoSistemas, "dd/MM/yy HH:mm")
                
                .Col = 10
                .FontSize = 8
                .FontName = "Verdana"
                .TypeHAlign = TypeHAlignLeft
                .Text = rs!Atendente & ""
            
                .Col = 11
                .FontSize = 8
                .FontName = "Verdana"
                .TypeHAlign = TypeHAlignCenter
                .Text = Format(rs!DataBaixa, "dd/MM/yy HH:mm")
            
                .Col = 12
                .FontSize = 8
                .FontName = "Verdana"
                .TypeHAlign = TypeHAlignCenter
                .Text = Format(rs!DataAceite, "dd/MM/yy HH:mm")
            Case Is = "Cancelada"
                .Col = 9
                .FontSize = 8
                .FontName = "Verdana"
                .TypeHAlign = TypeHAlignCenter
                .Text = Format(rs!DataCancelamento, "dd/MM/yy HH:mm")
        End Select
            
            pgbProgresso.Value = rs.AbsolutePosition
            rs.MoveNext
        Loop
        
        .OperationMode = OperationModeRead
        
    End With
    
    Set rs = Nothing
    MsgBox "Relatório gerado com sucesso!", vbOKOnly + vbInformation, "Suporte Manutenção"
    Screen.MousePointer = 0
    Exit Sub
    
Erro:
    MsgBox "Erro: " & Err.Description, vbOKOnly + vbCritical, "Suporte Manutenção"
    Set rs = Nothing
    Screen.MousePointer = 0
End Sub
