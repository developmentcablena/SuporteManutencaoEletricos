VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmReporteTecnico 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Suporte Manutenção - Reporte"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5910
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReporteTecnico.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   5910
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "&Validar"
      Height          =   375
      Left            =   2400
      TabIndex        =   11
      Top             =   6000
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtOSID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   960
      TabIndex        =   9
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton cmdBaixar 
      Appearance      =   0  'Flat
      Caption         =   "&Finalizar"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   5400
      Width           =   1335
   End
   Begin VB.CommandButton cmdFechar 
      Appearance      =   0  'Flat
      Caption         =   "Fechar"
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      Top             =   5400
      Width           =   1335
   End
   Begin MSMask.MaskEdBox mskDataBaixa 
      Height          =   315
      Left            =   2040
      TabIndex        =   2
      Top             =   4800
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      _Version        =   393216
      ClipMode        =   1
      Appearance      =   0
      MaxLength       =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/## ##:##"
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtReporteTecnico 
      Appearance      =   0  'Flat
      Height          =   2415
      Left            =   120
      MaxLength       =   500
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1080
      Width           =   5655
   End
   Begin VB.CommandButton cmdValidar 
      Appearance      =   0  'Flat
      Caption         =   "&Validar"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   6000
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSMask.MaskEdBox mskDataInicio 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   4800
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      _Version        =   393216
      ClipMode        =   1
      Appearance      =   0
      MaxLength       =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/## ##:##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskTempo 
      Height          =   315
      Left            =   3960
      TabIndex        =   3
      Top             =   4800
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      _Version        =   393216
      ClipMode        =   1
      Appearance      =   0
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##:##:##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskDataBaixaAtual 
      Height          =   315
      Left            =   2040
      TabIndex        =   14
      Top             =   3960
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      _Version        =   393216
      ClipMode        =   1
      Appearance      =   0
      Enabled         =   0   'False
      MaxLength       =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/## ##:##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskDataInicioAtual 
      Height          =   315
      Left            =   120
      TabIndex        =   15
      Top             =   3960
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      _Version        =   393216
      ClipMode        =   1
      Appearance      =   0
      Enabled         =   0   'False
      MaxLength       =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/## ##:##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskCalculo 
      Height          =   315
      Left            =   3960
      TabIndex        =   16
      Top             =   3960
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      _Version        =   393216
      ClipMode        =   1
      Appearance      =   0
      Enabled         =   0   'False
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##:##:##"
      PromptChar      =   "_"
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Data/ Hora Final"
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
      Left            =   2040
      TabIndex        =   19
      Top             =   3720
      Width           =   1635
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Data/ Hora Inicial"
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
      TabIndex        =   18
      Top             =   3720
      Width           =   1770
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Tempo Atend. Maq."
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
      Left            =   3960
      TabIndex        =   17
      Top             =   3720
      Width           =   1860
   End
   Begin VB.Label lblTempo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Tempo Parada Maq."
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
      Left            =   3960
      TabIndex        =   13
      Top             =   4560
      Width           =   1920
   End
   Begin VB.Label lblDataInicio 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Data/ Hora Inicial"
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
      TabIndex        =   12
      Top             =   4560
      Width           =   1770
   End
   Begin VB.Label lblOSID 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nº OS:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   120
      TabIndex        =   8
      Top             =   390
      Width           =   705
   End
   Begin VB.Label lblDataBaixa 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Data/ Hora Final"
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
      Left            =   2040
      TabIndex        =   7
      Top             =   4560
      Width           =   1635
   End
   Begin VB.Label lblReporteTecnico 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Reporte Técnico:"
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
      TabIndex        =   6
      Top             =   840
      Width           =   1635
   End
End
Attribute VB_Name = "frmReporteTecnico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rs As ADODB.Recordset
Private strSQL As String
Private strRelatorio As String

Private strMotivoOSNaoValidada As String
Private strDataOSNaoValidada As String
Private strReporte As String
Private strDataReporte As String

Private Sub cmdBaixar_Click()
    If Len(Trim(txtOSID.Text)) = 0 Then
        MsgBox "Nenhuma OS foi selecionada!", vbOKOnly + vbExclamation, "Suporte Manutenção"
        Exit Sub
    End If
    
    If Len(Trim(txtReporteTecnico.Text)) = 0 Then
        MsgBox "Digite o reporte técnico!", vbOKOnly + vbExclamation, "Suporte Manutenção"
        txtReporteTecnico.SetFocus
        Exit Sub
    End If
    
    If IsDate(mskDataBaixa.Text) = False Then
        MsgBox "Digite a data/ hora de finalização do serviço!", vbOKOnly + vbExclamation, "Suporte Manutenção"
        mskDataBaixa.SetFocus
        Exit Sub
    Else
        If CDate(mskDataBaixa.Text) <= CDate(mskDataInicio.Text) Then
            MsgBox "Data/ hora de finalização não pode ser menor ou igual a data de inicio!", vbOKOnly + vbExclamation, "Suporte Manutenção"
            mskDataBaixa.SetFocus
            Exit Sub
        End If
        If CDate(mskDataBaixa.Text) > Now Then
            MsgBox "Data/ hora de finalização não pode ser maior que a data atual!", vbOKOnly + vbExclamation, "Suporte Manutenção"
            mskDataBaixa.SetFocus
            Exit Sub
        End If
    End If
    
    Me.Enabled = False
    
    Dim tempoParadaEmSegundos As Long
    
    If fnPesquisarBaixa(Trim(txtOSID.Text)) = False Then
        'Call suLogarSituacao(Trim(txtOSID.Text))
        
        tempoParadaEmSegundos = fnTempoParadaEmSegundos(mskTempo.Text)
        
        Call suBaixarOS(gintOSID, Trim(txtReporteTecnico.Text), CDate(mskDataInicio.Text), CDate(mskDataBaixa.Text), fnCapturarEMail(gintOSID), fnCapturarUsuario(gintOSID), fnCapturarEmailAtendente(gintOSID), tempoParadaEmSegundos)
    Else
        Call suInserirReporteOcorrencia(gintOSID, strReporte, strDataReporte, strMotivoOSNaoValidada, strDataOSNaoValidada)
        
        tempoParadaEmSegundos = fnTempoParadaEmSegundos(mskTempo.Text)
        Call suBaixarOS(gintOSID, Trim(txtReporteTecnico.Text), CDate(mskDataInicio.Text), CDate(mskDataBaixa.Text), fnCapturarEMail(gintOSID), fnCapturarUsuario(gintOSID), fnCapturarEmailAtendente(gintOSID), tempoParadaEmSegundos)
    End If
    Me.Enabled = True
    Call Unload(Me)
End Sub

Private Sub suLogarSituacao(ByVal vOSID As Long)
On Error GoTo Erro
Dim rs1 As ADODB.Recordset
Dim strDataInicio As String
Dim strDataBaixa As String
Dim strTempoParada As String
Dim intAtendenteID As Integer

    strSQL = "SELECT Situacao, DataSituacao, DataInicio, DataBaixa, TempoParada, AtendenteID FROM tb_OS WHERE OSID = " & vOSID & ""
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        If IsNull(rs!DataInicio) = True Then
            strDataInicio = ""
        Else
            strDataInicio = rs!DataInicio
            strDataInicio = CDate(strDataInicio)
        End If
        
        If IsNull(rs!DataBaixa) = True Then
            strDataBaixa = ""
        Else
            strDataBaixa = rs!DataBaixa
            strDataBaixa = CDate(strDataBaixa)
        End If
        
        If IsNull(rs!tempoParada) = True Then
            strTempoParada = ""
        Else
            strTempoParada = rs!tempoParada
        End If
        
        If IsNull(rs!AtendenteID) = True Then
            intAtendenteID = 0
        Else
            intAtendenteID = rs!AtendenteID
        End If
        
        strSQL = "INSERT INTO tb_Log_Situacoes (Situacao,DataSituacao,OSID,DataInicio,DataBaixa,TempoParada,AtendenteID) VALUES ('" & rs!Situacao & "','" & rs!DataSituacao & "'," & vOSID & ",'" & strDataInicio & "','" & strDataBaixa & "','" & strTempoParada & "'," & intAtendenteID & ")"
        Set rs1 = New ADODB.Recordset
        rs1.Open strSQL, cn, adOpenKeyset, adLockOptimistic
        
        Set rs1 = Nothing
    End If
    
    rs.Close
    Set rs = Nothing
    Exit Sub
    
Erro:
    rs.Close
    Set rs = Nothing
    MsgBox "Erro " & Err.Number & vbCrLf & Err.Description, vbOKOnly + vbCritical, "Suporte Técnico"

End Sub

Private Sub suInserirReporteOcorrencia(ByVal vOSID As Long, ByVal vReporte As String, ByVal vDataReporte As String, ByVal vMotivo As String, ByVal vDataMotivo As String)
    strSQL = "INSERT INTO tb_Ocorrencias " & _
             "(ReporteTecnico,DataReporte,Ocorrencia,DataOcorrencia,Status,OSID) " & _
             "VALUES ('" & vReporte & "','" & vDataReporte & "','" & vMotivo & "','" & vDataMotivo & "',1," & vOSID & ")"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenKeyset, adLockOptimistic
    
    Set rs = Nothing
End Sub

Private Function fnPesquisarBaixa(ByVal vOSID As Long) As Boolean
    
    strMotivoOSNaoValidada = ""
    strDataOSNaoValidada = ""
    strReporte = ""
    strDataReporte = ""
    fnPesquisarBaixa = False
    
    strSQL = "SELECT * FROM tb_OS WHERE OSID=" & vOSID & " AND DataOSNaoValidada IS NOT NULL"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        strMotivoOSNaoValidada = rs!MotivoOSNaoValidada
        strDataOSNaoValidada = rs!DataOSNaoValidada
        strReporte = rs!ReporteTecnico
        strDataReporte = rs!DataBaixa
        fnPesquisarBaixa = True
    End If

    rs.Close
    Set rs = Nothing
End Function

Private Sub cmdFechar_Click()
    Call Unload(Me)
End Sub

Private Sub Form_Load()
    Dim datInicio As Date
    Dim datFim As Date
    Dim lngResultado As Long
    Dim strDataInicioAtual As String
    
    strDataInicioAtual = ""
    
    With Me
        .txtOSID.Text = Format(gintOSID, "0000")
        .mskDataInicioAtual.Text = Format(fnDataAtendimentoAtual(gintOSID), "dd/MM/yy HH:mm")
        .mskDataBaixaAtual.Text = Format(Now, "dd/MM/yy HH:mm")
        .mskDataInicio.Text = .mskDataInicioAtual.Text 'Format(Now, "dd/MM/yy HH:mm")
        .mskDataBaixa.Text = Format(Now, "dd/MM/yy HH:mm")
        datInicio = .mskDataInicioAtual.Text
        datFim = .mskDataBaixaAtual.Text
        lngResultado = DateDiff("n", datFim, datInicio)
        .mskCalculo.Mask = ""
        .mskCalculo.Text = ConvertLongToTime(lngResultado)
    End With
End Sub

Private Function fnDataAtendimentoAtual(ByVal vOSID As Long) As String
Dim rs1 As ADODB.Recordset

    fnDataAtendimentoAtual = ""
    
    strSQL = "SELECT * FROM vw_Chamados WHERE OSID = " & vOSID & ""
    Set rs1 = New ADODB.Recordset
    rs1.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If rs1.EOF = False Then
        fnDataAtendimentoAtual = rs1!DataAtendimentoAtual & ""
    End If
    
    rs1.Close
End Function

Public Function ConvertLongToTime(ByVal TheTime As Long) As String
   Dim hrs As Long, min As Long, seg As Long  'Declara as variáveis
   Dim sTime As String
   
   hrs = Abs(TheTime) \ 60    'Armazena a qtde de horas
   min = Abs(TheTime) Mod 60  'Armazena a qtde de minutos
   seg = 0
   
   'Completa a hora com os minutos
   sTime = Format$(hrs, "00") & ":" & Format$(min, "00") & ":" & Format$(seg, "00")

   'Retorna o resultado
   ConvertLongToTime = sTime
End Function

Private Function fnCapturarEmailAtendente(ByVal vOSID As Long) As String
Dim rs1 As New ADODB.Recordset
    
    fnCapturarEmailAtendente = ""
    
    strSQL = "SELECT * FROM vw_Chamados WHERE OSID = " & vOSID & ""
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        strSQL = "SELECT * FROM vw_Usuarios WHERE Nome = '" & rs!Atendente & "'"
        rs1.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
        
        If rs1.EOF = False Then
            fnCapturarEmailAtendente = rs1!Email & ""
        End If
        
        rs1.Close
        Set rs1 = Nothing
    End If
    
    rs.Close
    Set rs = Nothing
End Function

Private Function fnCapturarUsuario(ByVal vOSID As Long) As String
    fnCapturarUsuario = ""
    
    strSQL = "SELECT Nome FROM vw_Chamados WHERE OSID = " & vOSID & ""
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        fnCapturarUsuario = rs!Nome & ""
    End If
    
    rs.Close
    Set rs = Nothing
End Function

Private Function fnCapturarEMail(ByVal vOSID As Long) As String
    fnCapturarEMail = ""
    
    strSQL = "SELECT EMail FROM vw_Chamados WHERE OSID = " & vOSID & ""
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        fnCapturarEMail = rs!Email & ""
    End If
    
    rs.Close
    Set rs = Nothing
End Function

Private Sub suBaixarOS(ByVal vOSID As Long, ByVal vReporteTecnico As String, ByVal vDataInicio As Date, ByVal vDataBaixa As Date, ByVal vEMail As String, ByVal vUsuario As String, ByVal vEMailAtendente As String, ByVal vTempoParadaEmSegundos As Long)
On Error GoTo Erro

    strSQL = "UPDATE tb_OS SET ReporteTecnico = '" & vReporteTecnico & "',DataInicio = '" & vDataInicio & "', DataBaixa = '" & vDataBaixa & "', Status = 2, DataBaixaAtual = '" & Now & "', TempoParadaSegundos = " & vTempoParadaEmSegundos & " WHERE OSID = " & vOSID
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenKeyset, adLockOptimistic
    
    If fnEnviarEmail(vOSID, vDataBaixa, vReporteTecnico, vEMail, vUsuario, vEMailAtendente) = True Then
        MsgBox "OS finalizada com sucesso!", vbOKOnly + vbInformation, "Suporte Manutenção"
        Exit Sub
    Else
        MsgBox "Erro ao enviar e-mail!", vbOKOnly + vbCritical, "Suporte Manutenção"
        Exit Sub
    End If
    
Erro:
    MsgBox "Erro: " & Err.Description, vbOKOnly + vbCritical, "Suporte Manutenção"
End Sub

Private Function fnEnviarEmail(ByVal vOSID As Long, ByVal vDataBaixa As Date, ByVal vReporte As String, ByVal vEMail As String, ByVal vUsuario As String, ByVal vEMailAtendente As String) As Boolean
On Error GoTo Erro
Dim poSendMail As vbSendMail.clsSendMail
    
    fnEnviarEmail = False
    Me.Enabled = False
    Screen.MousePointer = 11
    
    DoEvents
    
    Set poSendMail = New vbSendMail.clsSendMail
    poSendMail.SMTPHost = "email-ssl.com.br"
    poSendMail.SMTPPort = "587"
    poSendMail.UseAuthentication = True
'    poSendMail.UserName = "sistema-interno@cablena.com.br"
'    poSendMail.Password = "Adm1nPW#2606"
'    poSendMail.From = "sistema-interno@cablena.com.br"
    poSendMail.UserName = fnContaSMTP
    poSendMail.Password = fnSenhaSMTP
    poSendMail.From = fnContaSMTP
    poSendMail.FromDisplayName = "ADM Suporte Manutenção"
    poSendMail.Recipient = vEMail
    poSendMail.RecipientDisplayName = vUsuario
    poSendMail.CcRecipient = vEMailAtendente
    poSendMail.Subject = "Suporte Manutenção - Finalização da OS " & Format(vOSID, "0000")
    poSendMail.Priority = HIGH_PRIORITY
    Call suRelatorio(vOSID)
    poSendMail.Message = strRelatorio
    poSendMail.Send

    Set poSendMail = Nothing
    fnEnviarEmail = True
    Screen.MousePointer = 0
    Me.Enabled = True
    Exit Function

Erro:
    Screen.MousePointer = 0
    Me.Enabled = True
    MsgBox "Erro: " & Err.Description, vbOKOnly = vbCritical, "Suporte Manutenção"
    Set poSendMail = Nothing

End Function

Private Function fnContaSMTP() As String
    fnContaSMTP = ""
    
    strSQL = "SELECT * FROM dbo.tb_Configuracoes WHERE ID = 1"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        fnContaSMTP = Trim(rs!Valor)
    End If
    
    rs.Close
    Set rs = Nothing
End Function

Private Function fnSenhaSMTP() As String
    fnSenhaSMTP = ""
    
    strSQL = "SELECT * FROM dbo.tb_Configuracoes WHERE ID = 2"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        fnSenhaSMTP = Trim(rs!Valor)
    End If
    
    rs.Close
    Set rs = Nothing
End Function

Private Sub suRelatorio(ByVal vOSID As Long)
    strSQL = "SELECT * FROM vw_Chamados WHERE OSID = " & vOSID & ""
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        strRelatorio = ""
        strRelatorio = String(100, "=") & vbCrLf & vbCrLf
        strRelatorio = strRelatorio & "Nº OS: " & Format(rs!OSID, "0000") & vbCrLf & vbCrLf
        strRelatorio = strRelatorio & "Tipo: " & rs!Divisao & String(10, " ") & "Caract.: " & rs!Tipo & String(10, " ") & "Linha: " & rs!Linha & String(10, " ") & "Especificação: " & rs!Especificacao & vbCrLf & vbCrLf
        strRelatorio = strRelatorio & "Reporte Usuário: " & rs!DescricaoServico & vbCrLf & vbCrLf
        strRelatorio = strRelatorio & "Reporte Técnico: " & rs!ReporteTecnico & vbCrLf & vbCrLf
        strRelatorio = strRelatorio & "Data Finalização: " & Format(rs!DataBaixa, "dd/MM/yy HH:mm") & vbCrLf & vbCrLf & vbCrLf & vbCrLf
        strRelatorio = strRelatorio & "*ATENÇÃO: Favor registrar o ACEITE no sistema." & vbCrLf & vbCrLf & vbCrLf
        strRelatorio = strRelatorio & "**OBSERVAÇÃO: Esta mensagem é gerada automaticamente pelo sistema." & vbCrLf & "POR FAVOR NÃO RESPONDA ESTA MENSAGEM." & vbCrLf & vbCrLf & vbCrLf & "Suporte Manutenção" & vbCrLf & "Cablena do Brasil" & vbCrLf & vbCrLf
        strRelatorio = strRelatorio & String(100, "=")
    End If
    
    rs.Close
    Set rs = Nothing
End Sub

