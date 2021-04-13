VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAceite 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Suporte Manutenção - Aceite Usuário"
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
   Icon            =   "frmAceite.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   5910
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdNaoValidar 
      Caption         =   "&Não Validar"
      Height          =   375
      Left            =   1440
      TabIndex        =   6
      Top             =   5400
      Width           =   1215
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
      Height          =   345
      Left            =   960
      MaxLength       =   5
      TabIndex        =   4
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "Fechar"
      Height          =   375
      Left            =   4560
      TabIndex        =   3
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceite 
      Caption         =   "&Validar"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox txtComentario 
      Appearance      =   0  'Flat
      Height          =   2415
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1050
      Width           =   5655
   End
   Begin MSMask.MaskEdBox mskDataBaixa 
      Height          =   315
      Left            =   2040
      TabIndex        =   7
      Top             =   4770
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
   Begin MSMask.MaskEdBox mskDataInicio 
      Height          =   315
      Left            =   120
      TabIndex        =   8
      Top             =   4770
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
   Begin MSMask.MaskEdBox mskTempo 
      Height          =   315
      Left            =   3960
      TabIndex        =   9
      Top             =   4770
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
   Begin MSMask.MaskEdBox mskDataBaixaAtual 
      Height          =   315
      Left            =   2040
      TabIndex        =   10
      Top             =   3930
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
      TabIndex        =   11
      Top             =   3930
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
      TabIndex        =   12
      Top             =   3930
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
      TabIndex        =   18
      Top             =   4530
      Width           =   1635
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
      TabIndex        =   17
      Top             =   4530
      Width           =   1770
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
      TabIndex        =   16
      Top             =   4530
      Width           =   1920
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
      TabIndex        =   15
      Top             =   3690
      Width           =   1860
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
      TabIndex        =   14
      Top             =   3690
      Width           =   1770
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
      TabIndex        =   13
      Top             =   3690
      Width           =   1635
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
      TabIndex        =   5
      Top             =   360
      Width           =   705
   End
   Begin VB.Label lblComentario 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Comentário/ Motivo:"
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
      Top             =   810
      Width           =   2010
   End
End
Attribute VB_Name = "frmAceite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rs As ADODB.Recordset
Private strSQL As String

Private strOcorrencia As String
Private strDataOcorrencia As String

Private Sub cmdAceite_Click()
    Call suCadastrarAceite(Trim(txtOSID.Text), Trim(txtComentario.Text))
    Call Unload(Me)
End Sub

Private Sub cmdFechar_Click()
    Call Unload(Me)
End Sub

Private Sub cmdNaoValidar_Click()
    If Len(Trim(txtComentario.Text)) > 0 Then
        Call suNaoValidar(Trim(txtOSID.Text), Trim(txtComentario.Text))
        Call Unload(Me)
    Else
        MsgBox "Digite o motivo pela não validação do serviço!", vbOKOnly + vbExclamation, "Suporte Manutenção"
    End If
End Sub

Private Sub Form_Load()
    Dim datInicio As Date
    Dim datFim As Date
    Dim lngResultado As Long
    Dim tempoParadaEmHoras As Long
    
    With Me
        .txtOSID.Text = Format(gintOSID, "0000")
        .mskDataInicioAtual.Mask = ""
        .mskDataInicioAtual.Text = Format(fnDataAtendimentoAtual(gintOSID), "dd/MM/yy HH:mm")
        .mskDataBaixaAtual.Mask = ""
        .mskDataBaixaAtual.Text = Format(fnDataBaixaAtual(gintOSID), "dd/MM/yy HH:mm")
        datInicio = .mskDataInicioAtual.Text
        datFim = .mskDataBaixaAtual.Text
        lngResultado = DateDiff("n", datFim, datInicio)
        .mskCalculo.Mask = ""
        .mskCalculo.Text = ConvertLongToTime(lngResultado)
        .mskDataInicio.Mask = ""
        .mskDataInicio.Text = Format(fnDataInicio(gintOSID), "dd/MM/yy HH:mm")
        .mskDataBaixa.Mask = ""
        .mskDataBaixa.Text = Format(fnDataBaixa(gintOSID), "dd/MM/yy HH:mm")
        .mskTempo.Mask = ""
        '.mskTempo.Text = Format(fnTempo(gintOSID), "HH:mm:ss")
        
        tempoParadaEmHoras = fnTempo(gintOSID)
        .mskTempo.Text = fnTempoParadaEmHoras(tempoParadaEmHoras)
        
        
    End With
End Sub

Private Function fnTempo(ByVal vOSID As Long) As Long
    fnTempo = 0
    
    strSQL = "SELECT TempoParadaSegundos FROM vw_Chamados WHERE OSID = " & vOSID & ""
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        fnTempo = IIf(IsNull(rs!TempoParadaSegundos) = True, 0, rs!TempoParadaSegundos)
    End If
    rs.Close
End Function

Private Function fnDataBaixa(ByVal vOSID As Long) As String
    fnDataBaixa = ""
    
    strSQL = "SELECT * FROM vw_Chamados WHERE OSID = " & vOSID & ""
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        fnDataBaixa = rs!DataBaixa
    End If
    rs.Close
End Function

Private Function fnDataInicio(ByVal vOSID As Long) As String
    fnDataInicio = ""
    
    strSQL = "SELECT * FROM vw_Chamados WHERE OSID = " & vOSID & ""
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        fnDataInicio = rs!DataInicio
    End If
    rs.Close
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

Private Function fnDataBaixaAtual(ByVal vOSID As Long) As String
    fnDataBaixaAtual = ""
    
    strSQL = "SELECT * FROM vw_Chamados WHERE OSID = " & vOSID & ""
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        fnDataBaixaAtual = rs!DataBaixaAtual
    End If
    rs.Close
End Function

Private Function fnDataAtendimentoAtual(ByVal vOSID As Long) As String
    fnDataAtendimentoAtual = ""
    
    strSQL = "SELECT * FROM vw_Chamados WHERE OSID = " & vOSID & ""
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        fnDataAtendimentoAtual = rs!DataAtendimentoAtual
    End If
    rs.Close
End Function

Private Sub suCadastrarAceite(ByVal vOSID As Long, ByVal vComentario As String)
    strSQL = "UPDATE tb_OS SET DataAceite = '" & Now & "',Comentario='" & vComentario & "', Status = 3 " & _
             "WHERE OSID = " & vOSID
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenKeyset, adLockOptimistic
    
    Set rs = Nothing
End Sub

Private Sub suNaoValidar(ByVal vOSID As Long, ByVal vMotivo As String)
Dim strDataOSNaoValidada As String
    
    strDataOSNaoValidada = Now
    
    If fnPesquisarOcorrencia(vOSID) = False Then
        'Call suGravarOcorrencia(vOSID, vMotivo, strDataOSNaoValidada)
        strSQL = "UPDATE tb_OS SET DataOSNaoValidada = '" & strDataOSNaoValidada & "',MotivoOSNaoValidada='" & vMotivo & "', Status = 6 " & _
                 "WHERE OSID = " & vOSID
        Set rs = New ADODB.Recordset
        rs.Open strSQL, cn, adOpenKeyset, adLockOptimistic
        Set rs = Nothing
    Else
        'Call suGravarOcorrencia(vOSID, strOcorrencia, strDataOcorrencia)
        strSQL = "UPDATE tb_OS SET DataOSNaoValidada = '" & Now & "',MotivoOSNaoValidada='" & vMotivo & "', Status = 6 " & _
                 "WHERE OSID = " & vOSID
        Set rs = New ADODB.Recordset
        rs.Open strSQL, cn, adOpenKeyset, adLockOptimistic
        Set rs = Nothing
    End If
End Sub

Private Sub suGravarOcorrencia(ByVal vOSID As Long, ByVal vMotivo As String, ByVal vDataOcorrencia As String)
    strSQL = "INSERT INTO tb_Ocorrencias (Ocorrencia,DataOcorrencia, Status,OSID) " & _
             "VALUES ('" & vMotivo & "','" & CDate(vDataOcorrencia) & "',1," & vOSID & " )"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenKeyset, adLockOptimistic
    
    Set rs = Nothing
End Sub

Private Function fnPesquisarOcorrencia(ByVal vOSID As Long) As Boolean
    fnPesquisarOcorrencia = False
    strOcorrencia = ""
    strDataOcorrencia = ""
    
    strSQL = "SELECT * FROM vw_Chamados WHERE OSID = " & vOSID & " and  DataOSNaoValidada IS NOT NULL"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        strOcorrencia = rs!MotivoOSNaoValidada & ""
        strDataOcorrencia = rs!DataOSNaoValidada & ""
        fnPesquisarOcorrencia = True
    End If
    
    rs.Close
    Set rs = Nothing
End Function
