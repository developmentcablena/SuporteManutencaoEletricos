VERSION 5.00
Begin VB.Form frmSuporteSistemas 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Suporte Manutenção - Ordem de Serviço"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7575
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSuporteSistemas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboLinha 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1200
      Width           =   2295
   End
   Begin VB.ComboBox cboAtendente 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmSuporteSistemas.frx":08CA
      Left            =   3000
      List            =   "frmSuporteSistemas.frx":08CC
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   5955
      Width           =   4455
   End
   Begin VB.TextBox txtReporteTecnico 
      Appearance      =   0  'Flat
      Height          =   1695
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   18
      Top             =   3960
      Width           =   7335
   End
   Begin VB.CheckBox chkPrioridade 
      Appearance      =   0  'Flat
      Caption         =   "Máquina Parada"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1920
      TabIndex        =   17
      Top             =   1635
      Width           =   2175
   End
   Begin VB.CommandButton cmdCancelar 
      Appearance      =   0  'Flat
      Caption         =   "C&ancelar"
      Height          =   375
      Left            =   1560
      TabIndex        =   16
      Top             =   6465
      Width           =   1335
   End
   Begin VB.TextBox txtPrevisao 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      TabIndex        =   15
      Top             =   5955
      Width           =   1335
   End
   Begin VB.TextBox txtPrazo 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   5955
      Width           =   1335
   End
   Begin VB.CommandButton cmdFechar 
      Appearance      =   0  'Flat
      Caption         =   "Fechar"
      Height          =   375
      Left            =   6120
      TabIndex        =   12
      Top             =   6465
      Width           =   1335
   End
   Begin VB.CommandButton cmdCadastrar 
      Appearance      =   0  'Flat
      Caption         =   "&Cadastrar"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   6480
      Width           =   1335
   End
   Begin VB.TextBox txtObservacao 
      Appearance      =   0  'Flat
      Height          =   1695
      Left            =   120
      MaxLength       =   500
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   1920
      Width           =   7335
   End
   Begin VB.ComboBox cboEspecificacao 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2520
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1200
      Width           =   4935
   End
   Begin VB.ComboBox cboTipo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2520
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   4935
   End
   Begin VB.ComboBox cboDivisao 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label lblLinha 
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
      TabIndex        =   23
      Top             =   960
      Width           =   525
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
      Left            =   3000
      TabIndex        =   22
      Top             =   5715
      Width           =   1005
   End
   Begin VB.Label lblLimite 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Limite de caracteres: "
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
      Left            =   4320
      TabIndex        =   21
      Top             =   1680
      Width           =   2130
   End
   Begin VB.Label lblQtdeCaracteres 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "500"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   6480
      TabIndex        =   20
      Top             =   1680
      Width           =   360
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
      TabIndex        =   19
      Top             =   3720
      Width           =   1635
   End
   Begin VB.Label lblPrevisao 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Necessidade"
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
      Left            =   1560
      TabIndex        =   14
      Top             =   5715
      Width           =   1230
   End
   Begin VB.Label lblPrazo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Prazo (dias)"
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
      TabIndex        =   13
      Top             =   5715
      Width           =   1200
   End
   Begin VB.Label lblDescricaoServico 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Reporte Usuário:"
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
      TabIndex        =   11
      Top             =   1680
      Width           =   1635
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
      Left            =   2520
      TabIndex        =   10
      Top             =   960
      Width           =   1275
   End
   Begin VB.Label lblTipo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Característica"
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
      TabIndex        =   9
      Top             =   240
      Width           =   1365
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
      TabIndex        =   8
      Top             =   240
      Width           =   420
   End
End
Attribute VB_Name = "frmSuporteSistemas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rs As ADODB.Recordset
Private strSQL As String
Private strRelatorio As String

Private Sub cboDivisao_Click()
    cboTipo.Clear
    cboLinha.Clear
    cboEspecificacao.Clear

    If Len(Trim(cboDivisao.Text)) > 0 Then
        Call suListarTipo(cboDivisao.Text)
    End If
End Sub

Private Sub cboLinha_Click()
    cboEspecificacao.Clear
    
    If Len(Trim(cboTipo.Text)) > 0 Then
        Call suListarEspecificacao(cboDivisao.Text, cboTipo.Text, cboLinha.Text)
    End If
End Sub

Private Sub cboTipo_Click()
    cboLinha.Clear
    cboEspecificacao.Clear
    
    If Len(Trim(cboTipo.Text)) > 0 Then
        Call suListarLinha(cboDivisao.Text, cboTipo.Text)
    End If
End Sub

Private Sub chkPrioridade_Click()
    If chkPrioridade.Value = 1 Then
        chkPrioridade.ForeColor = &HFF&
    Else
        chkPrioridade.ForeColor = &H80000008
    End If
End Sub

Private Sub cmdCancelar_Click()
    If cmdCadastrar.Caption = "&Cadastrar" Then
        Call suLimparCampos
    Else
        Call Unload(Me)
    End If
End Sub

Private Sub cmdFechar_Click()
    Call Unload(Me)
End Sub

Private Sub Form_Load()
    With Me
        .lblReporteTecnico.Enabled = True
        .txtReporteTecnico.Enabled = False
    End With
    Call suListarDivisoes
    Call suListarAtend
End Sub

Private Sub lblQtdeCaracteres_Change()
    If lblQtdeCaracteres.Caption = 0 Then
        lblQtdeCaracteres.ForeColor = vbRed
    Else
        lblQtdeCaracteres.ForeColor = vbBlue
    End If
End Sub

Private Sub txtObservacao_Change()
    lblQtdeCaracteres.Caption = 500 - Len(txtObservacao.Text)
End Sub

Private Sub txtPrazo_LostFocus()
    If Len(Trim(txtPrazo.Text)) > 0 Then
        If IsNumeric(txtPrazo.Text) = True Then
            If txtPrazo.Text >= 0 Then
                txtPrevisao.Text = Format(Now + Trim(txtPrazo.Text), "dd/MM/yyyy")
            Else
                MsgBox "Prazo não pode ser menor que 0!", vbOKOnly + vbExclamation, "Suporte Manutenção"
                txtPrazo.Text = ""
                txtPrazo.SetFocus
            End If
        Else
            MsgBox "Digite a quantidade de dias!", vbOKOnly + vbExclamation, "Suporte Manutenção"
            txtPrazo.Text = ""
            txtPrazo.SetFocus
        End If
    Else
        MsgBox "Digite um prazo para o serviço!", vbOKOnly + vbExclamation, "Suporte Manutenção"
        txtPrazo.SetFocus
    End If
End Sub

Private Sub cmdCadastrar_Click()
    If Len(Trim(cboDivisao.Text)) = 0 Then
        MsgBox "Selecione uma divisão!", vbOKOnly + vbExclamation, "Suporte Manutenção"
        cboDivisao.SetFocus
        Exit Sub
    End If
    
    If Len(Trim(cboTipo.Text)) = 0 Then
        MsgBox "Selecione um tipo!", vbOKOnly + vbExclamation, "Suporte Manutenção"
        cboTipo.SetFocus
        Exit Sub
    End If
    
    If Len(Trim(cboEspecificacao.Text)) = 0 Then
        MsgBox "Selecione uma especificação!", vbOKOnly + vbExclamation, "Suporte Manutenção"
        cboEspecificacao.SetFocus
        Exit Sub
    End If
    
    If Len(Trim(txtObservacao.Text)) = 0 Then
        MsgBox "Digite a descrição do serviço à ser realizado!", vbOKOnly + vbExclamation, "Suporte Manutenção"
        txtObservacao.SetFocus
        Exit Sub
    End If
    
    If Len(Trim(txtPrazo.Text)) = 0 Then
        MsgBox "Digite um prazo para o serviço!", vbOKOnly + vbExclamation, "Suporte Manutenção"
        txtPrazo.SetFocus
        Exit Sub
    End If
    
    If Len(Trim(cboAtendente.Text)) = 0 Then
        MsgBox "Selecione um atendente!", vbOKOnly + vbExclamation, "Suporte Manutenção"
        cboAtendente.SetFocus
        Exit Sub
    End If
    
    If cmdCadastrar.Caption = "&Cadastrar" Then
        If fnCadastrarOS(cboDivisao.Text, cboTipo.Text, cboEspecificacao.Text, gintUsuarioID, gstrEMail, Trim(txtObservacao.Text), chkPrioridade, Trim(txtPrazo.Text), Trim(txtPrevisao.Text), Trim(cboAtendente.Text), Trim(cboLinha.Text)) = True Then
            MsgBox "Ordem de serviço nº " & Format(gintOSID, "0000") & " gerada com sucesso!", vbOKOnly + vbInformation, "Suporte Manutenção"
            Call suLimparCampos
        End If
    Else
        If fnAlterarOS(cboDivisao.Text, cboTipo.Text, cboEspecificacao.Text, Trim(txtObservacao.Text), chkPrioridade, Trim(txtPrazo.Text), Trim(txtPrevisao.Text), Trim(cboAtendente.Text), Trim(cboLinha.Text)) = True Then
            MsgBox "Ordem de serviço nº " & Format(gintOSID, "0000") & " alterada com sucesso!", vbOKOnly + vbInformation, "Suporte Manutenção"
            Call Unload(Me)
        End If
    End If
End Sub

Private Function fnAlterarOS(ByVal vDivisao As String, ByVal vTipo As String, ByVal vEspecificacao As String, ByVal vDescricao As String, ByRef vPrioridade As CheckBox, ByVal vPrazo As Integer, ByVal vPrevisao As Variant, ByVal vAtendente As String, ByVal vLinha As String) As Boolean
On Error GoTo Erro
    fnAlterarOS = False
    
    strSQL = "UPDATE tb_OS SET DivisaoID = " & fnDivisaoID(vDivisao) & ",TipoID = " & fnTipoID(vDivisao, vTipo) & ",EspecificacaoID = " & fnEspecificacaoID(vDivisao, vTipo, vLinha, vEspecificacao) & ",DescricaoServico = '" & vDescricao & "',Prioridade = " & vPrioridade & ",Prazo = " & vPrazo & ",Previsao = '" & vPrevisao & "',Atendente = '" & vAtendente & "',LinhaID = " & fnLinhaID(vDivisao, vTipo, vLinha) & " WHERE OSID = " & gintOSID & ""
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenKeyset, adLockOptimistic
    
    fnAlterarOS = True
    Set rs = Nothing
    Exit Function
    
Erro:
    MsgBox "Erro: " & Err.Description, vbOKOnly + vbCritical, "Suporte Manutenção"
    Set rs = Nothing
End Function

Private Function fnCadastrarOS(ByVal vDivisao As String, ByVal vTipo As String, ByVal vEspecificacao As String, ByVal vUsuarioID As Integer, ByVal vEMail As String, ByVal vDescricao As String, ByRef vPrioridade As CheckBox, ByVal vPrazo As Integer, ByVal vPrevisao As Variant, ByVal vAtendente As String, ByVal vLinha As String) As Boolean
On Error GoTo Erro
    fnCadastrarOS = False
    gintOSID = 0
    
    strSQL = "SELECT * FROM tb_OS WHERE 1=2"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenKeyset, adLockOptimistic
    
    If rs.EOF = True Then
        rs.AddNew
        rs!DivisaoID = 0 & fnDivisaoID(vDivisao)
        rs!TipoID = 0 & fnTipoID(vDivisao, vTipo)
        rs!EspecificacaoID = 0 & fnEspecificacaoID(vDivisao, vTipo, vLinha, vEspecificacao)
        rs!UsuarioID = vUsuarioID
        rs!Email = vEMail & ""
        rs!DescricaoServico = vDescricao & ""
        rs!Prioridade = vPrioridade
        rs!Prazo = 0 & vPrazo
        rs!Previsao = vPrevisao
        rs!DataCadastro = Format(Now, "dd/MM/yyyy HH:mm:ss")
        rs!Status = 0
        rs!Atendente = vAtendente & ""
        rs!LinhaID = 0 & fnLinhaID(vDivisao, vTipo, vLinha)
        rs.Update
        gintOSID = rs!OSID
        fnCadastrarOS = True
        
        If fnEnviarEmail(gintOSID, gstrEMail, gstrNome, fnEmailAtend(vAtendente)) = False Then
            MsgBox "Erro ao enviar o e-mail!", vbOKOnly + vbCritical, "Suporte Manutenção"
            Exit Function
        End If
    End If
    
    Set rs = Nothing
    Exit Function
    
Erro:
    MsgBox "Erro: " & Err.Description, vbOKOnly + vbCritical, "Suporte Manutenção"
    Set rs = Nothing
End Function

Private Function fnEmailAtend(ByVal vAtendente As String) As String
Dim rs1 As New ADODB.Recordset

    strSQL = "SELECT Email FROM tb_Usuarios WHERE Nome = '" & vAtendente & "'"
    rs1.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If rs1.EOF = False Then
        fnEmailAtend = rs1!Email
    End If
    
    rs1.Close
    Set rs1 = Nothing
End Function

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

Private Function fnLinhaID(ByVal vDivisao As String, ByVal vTipo As String, ByVal vLinha As String) As Integer
Dim rs1 As New ADODB.Recordset

    strSQL = "SELECT * FROM tb_Linhas WHERE Linha = '" & vLinha & "' AND TipoID = " & fnTipoID(vDivisao, vTipo)
    rs1.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If rs1.EOF = False Then
        fnLinhaID = rs1!LinhaID
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

Private Function fnEspecificacaoID(ByVal vDivisao As String, ByVal vTipo As String, ByVal vLinha As String, ByVal vEspecificacao As String) As Double
Dim rs1 As New ADODB.Recordset

    strSQL = "SELECT * FROM vw_Especificacoes WHERE Linha = '" & vLinha & "' AND Especificacao = '" & vEspecificacao & "' AND DivisaoID = " & fnDivisaoID(vDivisao) & " AND TipoID = " & fnTipoID(vDivisao, vTipo)
    rs1.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If rs1.EOF = False Then
        fnEspecificacaoID = rs1!EspecificacaoID
    End If
    
    rs1.Close
    Set rs1 = Nothing
End Function

Private Function fnEnviarEmail(ByVal vOSID As Long, ByVal vEMail As String, ByVal vUsuario As String, ByVal vEmailAtend As String) As Boolean
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
    'poSendMail.UserName = "sistema-interno@cablena.com.br"
    'poSendMail.Password = ""
    'poSendMail.From = "sistema-interno@cablena.com.br"
    poSendMail.UserName = fnContaSMTP
    poSendMail.Password = fnSenhaSMTP
    poSendMail.From = fnContaSMTP
    poSendMail.FromDisplayName = vUsuario
    poSendMail.Recipient = "suporte_man_eletricos@cablenadobrasil.com.br"
    poSendMail.RecipientDisplayName = "ADM Suporte Manutenção"
    'poSendMail.CcRecipient = "arbarbeiro@cablena.com.br"
    poSendMail.Subject = "Suporte Manutenção - Nova OS " & Format(vOSID, "0000")
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
        strRelatorio = strRelatorio & "Atendente: " & rs!Atendente & vbCrLf & vbCrLf
        strRelatorio = strRelatorio & "Data Cadastro: " & Format(rs!DataCadastro, "dd/MM/yy HH:mm") & vbCrLf & vbCrLf
        strRelatorio = strRelatorio & String(100, "=")
    End If
    
    rs.Close
    Set rs = Nothing
End Sub

Private Sub suLimparCampos()
    cboDivisao.Clear
    cboAtendente.Clear
    Call suListarDivisoes
    Call suListarAtend
    Call cboDivisao_Click
    txtObservacao.Text = ""
    chkPrioridade.Value = 0
    txtPrazo.Text = ""
    txtPrevisao.Text = ""
End Sub

Private Sub suListarDivisoes()
    If gstrTela = "frmChamados" Or gstrTela = "frmAcompanharOS" Then
        strSQL = "SELECT * FROM tb_Divisao ORDER BY DivisaoID"
    Else
        strSQL = "SELECT * FROM tb_Divisao WHERE Inativo = 0 ORDER BY DivisaoID"
    End If
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    Do While Not rs.EOF
        cboDivisao.AddItem rs!Divisao
        rs.MoveNext
    Loop
    
    Set rs = Nothing
End Sub

Private Sub suListarTipo(ByVal vDivisao As String)
    If gstrTela = "frmChamados" Or gstrTela = "frmAcompanharOS" Then
        strSQL = "SELECT * FROM vw_Tipos WHERE Divisao = '" & vDivisao & "'"
    Else
        strSQL = "SELECT * FROM vw_Tipos WHERE Divisao = '" & vDivisao & "' AND Inativo = 0"
    End If
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    Do While Not rs.EOF
        cboTipo.AddItem rs!Tipo
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
End Sub

Private Sub suListarLinha(ByVal vDivisao As String, ByVal vTipo As String)
    If gstrTela = "frmChamados" Or gstrTela = "frmAcompanharOS" Then
        strSQL = "SELECT * FROM vw_Linhas WHERE TipoID = " & fnTipoID(vDivisao, vTipo) & ""
    Else
        strSQL = "SELECT * FROM vw_Linhas WHERE TipoID = " & fnTipoID(vDivisao, vTipo) & " AND Inativo = 0"
    End If
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    Do While Not rs.EOF
        cboLinha.AddItem rs!Linha
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
End Sub

Private Sub suListarEspecificacao(ByVal vDivisao As String, ByVal vTipo As String, ByVal vLinha As String)
    If gstrTela = "frmChamados" Or gstrTela = "frmAcompanharOS" Then
        strSQL = "SELECT * FROM vw_Especificacoes WHERE LinhaID = " & fnLinhaID(vDivisao, vTipo, vLinha) & ""
        'strSQL = "SELECT DISTINCT(REPLACE(LEFT(Especificacao,CHARINDEX('-',Especificacao)),'-','')) AS Especificacao FROM vw_Especificacoes WHERE TipoID = " & fnTipoID(vDivisao, vTipo) & ""
    Else
        strSQL = "SELECT * FROM vw_Especificacoes WHERE LinhaID = " & fnLinhaID(vDivisao, vTipo, vLinha) & " AND Inativo = 0"
        'strSQL = "SELECT DISTINCT(REPLACE(LEFT(Especificacao,CHARINDEX('-',Especificacao)),'-','')) AS Especificacao FROM vw_Especificacoes WHERE TipoID = " & fnTipoID(vDivisao, vTipo) & " AND Inativo = 0"
    End If
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    Do While Not rs.EOF
        cboEspecificacao.AddItem rs!Especificacao
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
End Sub

Private Sub suListarAtend()
    If gstrTela = "frmChamados" Or gstrTela = "frmAcompanharOS" Then
        strSQL = "SELECT * FROM tb_Usuarios WHERE Departamento IN ('Manutenção','Ferramentaria') ORDER BY Nome"
    Else
        strSQL = "SELECT * FROM tb_Usuarios WHERE Departamento IN ('Manutenção','Ferramentaria') AND Inativo = 0 ORDER BY Nome"
    End If
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    Do While Not rs.EOF
        cboAtendente.AddItem rs!Nome
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
End Sub
