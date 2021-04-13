VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRelatorioAtend 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Filtro por Atendente"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6015
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   6015
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtEspecificacao 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3600
      TabIndex        =   22
      Top             =   2160
      Width           =   2295
   End
   Begin VB.ComboBox cboEspecificacao 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   2160
      Width           =   1935
   End
   Begin VB.TextBox txtLinha 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3600
      TabIndex        =   19
      Top             =   1680
      Width           =   2295
   End
   Begin VB.ComboBox cboLinha 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   1680
      Width           =   1935
   End
   Begin VB.TextBox txtOSDe 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3600
      TabIndex        =   17
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtOSAte 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4800
      TabIndex        =   16
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtDataCadAte 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4800
      TabIndex        =   15
      Top             =   1200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtDataCadDe 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3600
      TabIndex        =   14
      Top             =   1200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdExportar 
      Caption         =   "&Exportar XLS"
      Height          =   375
      Left            =   1680
      TabIndex        =   13
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
      Height          =   375
      Left            =   4440
      TabIndex        =   10
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton cmdGerar 
      Caption         =   "&Gerar"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   3360
      Width           =   1455
   End
   Begin VB.ComboBox cboDataCad 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox txtDataCad 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3600
      TabIndex        =   6
      Top             =   1200
      Width           =   2295
   End
   Begin VB.ComboBox cboOSF 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   720
      Width           =   1935
   End
   Begin VB.TextBox txtOS 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3600
      TabIndex        =   3
      Top             =   720
      Width           =   2295
   End
   Begin VB.ComboBox cboAtendF 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   240
      Width           =   1935
   End
   Begin VB.TextBox txtAtendente 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3600
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
   Begin FPSpreadADO.fpSpread fpsRelatorio 
      Height          =   3255
      Left            =   120
      TabIndex        =   11
      Top             =   4080
      Width           =   5535
      _Version        =   458752
      _ExtentX        =   9763
      _ExtentY        =   5741
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
      SpreadDesigner  =   "frmRelatorioAtend.frx":0000
   End
   Begin MSComctlLib.ProgressBar pgbProgresso 
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   2760
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin MSComDlg.CommonDialog cdlSalvarArquivo 
      Left            =   5880
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Pasta de Trabalho do Microsoft Office Excel (*.xls) | *.xls"
   End
   Begin VB.Label Label2 
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
      Left            =   120
      TabIndex        =   23
      Top             =   2280
      Width           =   1275
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
      Left            =   120
      TabIndex        =   20
      Top             =   1800
      Width           =   525
   End
   Begin VB.Label lblDataCad 
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
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   1380
   End
   Begin VB.Label lblOS 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "OS"
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
      TabIndex        =   4
      Top             =   840
      Width           =   255
   End
   Begin VB.Label lblAtendente 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Atendente"
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
      TabIndex        =   1
      Top             =   360
      Width           =   1005
   End
End
Attribute VB_Name = "frmRelatorioAtend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rs As ADODB.Recordset
Private strSQL As String

Private Sub cboDataCad_Click()
    If cboDataCad.Text = "Entre" Then
        txtDataCad.Visible = False
        txtDataCadDe.Visible = True
        txtDataCadAte.Visible = True
    Else
        txtDataCad.Visible = True
        txtDataCadDe.Visible = False
        txtDataCadAte.Visible = False
    End If
End Sub

Private Sub cboOSF_Click()
    If cboOSF.Text = "Entre" Then
        txtOS.Visible = False
        txtOSDe.Visible = True
        txtOSAte.Visible = True
    Else
        txtOS.Visible = True
        txtOSDe.Visible = False
        txtOSAte.Visible = False
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

    If Trim(txtOS.Text) <> "" Then
        If IsNumeric(Trim(txtOS.Text)) = False Then
            MsgBox "OS inválida!", vbOKOnly + vbExclamation, "Suporte Manutenção"
            txtOS.Text = ""
            txtOS.SetFocus
            Exit Sub
        End If
    End If
    
    If Trim(txtOSDe.Text) <> "" Then
        If IsNumeric(Trim(txtOSDe.Text)) = False Then
            MsgBox "OS inválida!", vbOKOnly + vbExclamation, "Suporte Manutenção"
            txtOSDe.Text = ""
            txtOSDe.SetFocus
            Exit Sub
        End If
    End If
    
    If Trim(txtOSAte.Text) <> "" Then
        If IsNumeric(Trim(txtOSAte.Text)) = False Then
            MsgBox "OS inválida!", vbOKOnly + vbExclamation, "Suporte Manutenção"
            txtOSAte.Text = ""
            txtOSAte.SetFocus
            Exit Sub
        End If
    End If
    
    If Trim(txtDataCad.Text) <> "" Then
        If IsDate(Trim(txtDataCad.Text)) = False Then
            MsgBox "Data inválida!", vbOKOnly + vbExclamation, "Suporte Manutenção"
            txtDataCad.Text = ""
            txtDataCad.SetFocus
            Exit Sub
        End If
    End If
    
    If Trim(txtDataCadDe.Text) <> "" Then
        If IsDate(Trim(txtDataCadDe.Text)) = False Then
            MsgBox "Data inválida!", vbOKOnly + vbExclamation, "Suporte Manutenção"
            txtDataCadDe.Text = ""
            txtDataCadDe.SetFocus
            Exit Sub
        End If
    End If
    
    If Trim(txtDataCadAte.Text) <> "" Then
        If IsDate(Trim(txtDataCadAte.Text)) = False Then
            MsgBox "Data inválida!", vbOKOnly + vbExclamation, "Suporte Manutenção"
            txtDataCadAte.Text = ""
            txtDataCadAte.SetFocus
            Exit Sub
        End If
    End If
    
    Call suGerarRelatorioDetalhado(Trim(txtAtendente.Text), cboAtendF.Text, IIf(Trim(txtOS.Text) = "", 0, Trim(txtOS.Text)), IIf(Trim(txtOSDe.Text) = "", 0, Trim(txtOSDe.Text)), IIf(Trim(txtOSAte.Text) = "", 0, Trim(txtOSAte.Text)), cboOSF.Text, Trim(txtDataCad.Text), Trim(txtDataCadDe.Text), Trim(txtDataCadAte.Text), cboDataCad.Text, Trim(txtLinha.Text), Trim(txtEspecificacao.Text), cboLinha.Text, cboEspecificacao.Text)
End Sub

Private Sub Form_Load()
    Call suListarAtendF
    Call suListarOSF
    Call suListarDataCad
    Call suListarLinha
    Call suListarEspecificacao
End Sub

Private Sub suListarLinha()
    cboLinha.AddItem "Igual"
    cboLinha.AddItem "Iniciado por"
    cboLinha.AddItem "Terminado por"
    cboLinha.AddItem "Contém"
End Sub

Private Sub suListarEspecificacao()
    cboEspecificacao.AddItem "Igual"
    cboEspecificacao.AddItem "Iniciado por"
    cboEspecificacao.AddItem "Terminado por"
    cboEspecificacao.AddItem "Contém"
End Sub

Private Sub suListarAtendF()
    cboAtendF.AddItem "Igual"
    cboAtendF.AddItem "Iniciado por"
    cboAtendF.AddItem "Terminado por"
    cboAtendF.AddItem "Contém"
End Sub

Private Sub suListarOSF()
    cboOSF.AddItem "Igual"
    cboOSF.AddItem "Maior igual"
    cboOSF.AddItem "Menor igual"
    cboOSF.AddItem "Entre"
End Sub

Private Sub suListarDataCad()
    cboDataCad.AddItem "Igual"
    cboDataCad.AddItem "Maior igual"
    cboDataCad.AddItem "Menor igual"
    cboDataCad.AddItem "Entre"
End Sub

Private Sub suGerarRelatorioDetalhado(ByVal vAtend As String, ByVal vAtendF As String, ByVal vOS As Long, ByVal vOSDe As Integer, ByVal vOSAte As Integer, ByVal vOSF As String, ByVal vDataCad As Variant, ByVal vDataCadDe As Variant, ByVal vDataCadAte As Variant, ByVal vDataCadF As String, ByVal vLinha As String, ByVal vEspecificacao As String, ByVal vLinhaF As String, ByVal vEspecificacaoF As String)
On Error GoTo Erro

Dim blnWhere As Boolean
Dim strStatus As String

    strSQL = "SELECT * FROM dbo.vw_Relatorio_Por_Atendente"
    blnWhere = False
    
    If Len(vAtend) > 0 Then
        Select Case vAtendF
            Case "Igual"
                strSQL = strSQL & IIf(blnWhere = True, " AND ", " WHERE ") & "Atendente = '" & vAtend & "'"
                blnWhere = True
            Case "Iniciado por"
                strSQL = strSQL & IIf(blnWhere = True, " AND ", " WHERE ") & "Atendente LIKE '" & vAtend & "%'"
                blnWhere = True
            Case "Terminado por"
                strSQL = strSQL & IIf(blnWhere = True, " AND ", " WHERE ") & "Atendente LIKE '%" & vAtend & "'"
                blnWhere = True
            Case "Contém"
                strSQL = strSQL & IIf(blnWhere = True, " AND ", " WHERE ") & "Atendente LIKE '%" & vAtend & "%'"
                blnWhere = True
        End Select
    End If
    
    If Len(vOS) > 0 Then
        Select Case vOSF
            Case "Igual"
                strSQL = strSQL & IIf(blnWhere = True, " AND ", " WHERE ") & "OSID = " & vOS & ""
                blnWhere = True
            Case "Maior igual"
                strSQL = strSQL & IIf(blnWhere = True, " AND ", " WHERE ") & "OSID >= " & vOS & ""
                blnWhere = True
            Case "Menor igual"
                strSQL = strSQL & IIf(blnWhere = True, " AND ", " WHERE ") & "OSID <= " & vOS & ""
                blnWhere = True
        End Select
    End If
    
    If vOSDe > 0 And vOSAte > 0 Then
        strSQL = strSQL & IIf(blnWhere = True, " AND ", " WHERE ") & "OSID >= " & vOSDe & " AND OSID <= " & vOSAte & ""
        blnWhere = True
    ElseIf vOSDe > 0 And vOSAte = 0 Then
        strSQL = strSQL & IIf(blnWhere = True, " AND ", " WHERE ") & "OSID >= " & vOSDe & ""
        blnWhere = True
    ElseIf vOSDe = 0 And vOSAte > 0 Then
        strSQL = strSQL & IIf(blnWhere = True, " AND ", " WHERE ") & "OSID <= " & vOSAte & ""
        blnWhere = True
    End If
    
    If Len(vDataCad) > 0 Then
        Select Case vDataCadF
            Case "Igual"
                strSQL = strSQL & IIf(blnWhere = True, " AND ", " WHERE ") & "DataCadastro >= '" & vDataCad & " 00:00:00' AND DataCadastro <= '" & vDataCad & " 23:59:59'"
                blnWhere = True
            Case "Maior igual"
                strSQL = strSQL & IIf(blnWhere = True, " AND ", " WHERE ") & "DataCadastro >= '" & vDataCad & " 00:00:00'"
                blnWhere = True
            Case "Menor igual"
                strSQL = strSQL & IIf(blnWhere = True, " AND ", " WHERE ") & "DataCadastro <= '" & vDataCad & " 23:59:59'"
                blnWhere = True
        End Select
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
    
    If Len(vLinha) > 0 Then
        Select Case vLinhaF
            Case "Igual"
                strSQL = strSQL & IIf(blnWhere = True, " AND ", " WHERE ") & "Linha = '" & vLinha & "'"
                blnWhere = True
            Case "Iniciado por"
                strSQL = strSQL & IIf(blnWhere = True, " AND ", " WHERE ") & "Linha LIKE '" & vLinha & "%'"
                blnWhere = True
            Case "Terminado por"
                strSQL = strSQL & IIf(blnWhere = True, " AND ", " WHERE ") & "Linha LIKE '%" & vLinha & "'"
                blnWhere = True
            Case "Contém"
                strSQL = strSQL & IIf(blnWhere = True, " AND ", " WHERE ") & "Linha LIKE '%" & vLinha & "%'"
                blnWhere = True
        End Select
    End If
        
    If Len(vEspecificacao) > 0 Then
        Select Case vEspecificacaoF
            Case "Igual"
                strSQL = strSQL & IIf(blnWhere = True, " AND ", " WHERE ") & "Especificacao = '" & vEspecificacao & "'"
                blnWhere = True
            Case "Iniciado por"
                strSQL = strSQL & IIf(blnWhere = True, " AND ", " WHERE ") & "Especificacao LIKE '" & vEspecificacao & "%'"
                blnWhere = True
            Case "Terminado por"
                strSQL = strSQL & IIf(blnWhere = True, " AND ", " WHERE ") & "Especificacao LIKE '%" & vEspecificacao & "'"
                blnWhere = True
            Case "Contém"
                strSQL = strSQL & IIf(blnWhere = True, " AND ", " WHERE ") & "Especificacao LIKE '%" & vEspecificacao & "%'"
                blnWhere = True
        End Select
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
        .Text = "DEPTO"
        
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
        .Text = "DATA SITUAÇÃO"

        .Col = 19
        .ColWidth(19) = 30
        .Text = "SITUAÇÃO"

        .Col = 20
        .ColWidth(20) = 15
        .Text = "DATA ATENDIMENTO SISTEMA"

        .Col = 21
        .ColWidth(21) = 15
        .Text = "DATA BAIXA SISTEMA"

        Do While Not rs.EOF
        DoEvents
            .Row = .Row + 1
            
            .Col = 1
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignLeft
            .Text = rs!Solicitante & ""
            
            .Col = 2
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignLeft
            .Text = rs!Depto & ""
            
            .Col = 3
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignRight
            .Text = CStr(Format(rs!OSID, "0000")) & ""
            
            .Col = 4
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignLeft
            .Text = rs!Status & ""
            
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
            
            Dim tempoParadaEmHoras As String
            
            If IsNull(rs!TempoParadaSegundos) Then
                .Text = Format(rs!tempoParada, "HH:mm:ss") & ""
            Else
                tempoParadaEmHoras = fnTempoParadaEmHoras(rs!TempoParadaSegundos)
                .Text = tempoParadaEmHoras & ""
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
            .Text = Format(rs!DataSituacao, "dd/MM/yy HH:mm")
            
            .Col = 19
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignLeft
            .Text = rs!Situacao & ""
            
            .Col = 20
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignCenter
            .Text = Format(rs!DataAtendimentoAtual, "dd/MM/yy HH:mm")
            
            .Col = 21
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignCenter
            .Text = Format(rs!DataBaixaAtual, "dd/MM/yy HH:mm")
            
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

