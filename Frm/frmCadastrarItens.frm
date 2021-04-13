VERSION 5.00
Begin VB.Form frmCadastrarItens 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Suporte Manutenção - Cadastrar Itens"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5295
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCadastrarItens.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   5295
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAtualEspec 
      Enabled         =   0   'False
      Height          =   255
      Left            =   4770
      Picture         =   "frmCadastrarItens.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   825
      Width           =   375
   End
   Begin VB.CommandButton cmdAtualCaract 
      Enabled         =   0   'False
      Height          =   255
      Left            =   4770
      Picture         =   "frmCadastrarItens.frx":0784
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   105
      Width           =   375
   End
   Begin VB.CommandButton cmdAtualTipo 
      Enabled         =   0   'False
      Height          =   255
      Left            =   2010
      Picture         =   "frmCadastrarItens.frx":0AC6
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   105
      Width           =   375
   End
   Begin VB.CommandButton cmdCadastrar 
      Caption         =   "&Cadastrar"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "Fechar"
      Height          =   375
      Left            =   3960
      TabIndex        =   6
      Top             =   1560
      Width           =   1215
   End
   Begin VB.ComboBox cboEspecificacao 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   1080
      Width           =   5055
   End
   Begin VB.ComboBox cboTipo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2520
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   360
      Width           =   2655
   End
   Begin VB.ComboBox cboDivisao 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label lblEspecificacao 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Especificação"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1140
   End
   Begin VB.Label lblTipo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Característica"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2520
      TabIndex        =   3
      Top             =   120
      Width           =   1185
   End
   Begin VB.Label lblDivisao 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Tipo"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   360
   End
End
Attribute VB_Name = "frmCadastrarItens"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rs As ADODB.Recordset
Private strSQL As String
Private intDivisaoID As Integer
Private intTipoID As Integer
Private intEspecificacaoID As Integer

Private Sub cboDivisao_Change()
    If Len(Trim(cboDivisao.Text)) = 0 Then
        cmdAtualTipo.Enabled = False
        cboTipo.Clear
        cboEspecificacao.Clear
        Call cboTipo_Change
        Call cboEspecificacao_Change
    End If
End Sub

Private Sub cboDivisao_Click()
    cboTipo.Clear
    cboEspecificacao.Clear

    If Len(Trim(cboDivisao.Text)) > 0 Then
        Call suListarTipo(cboDivisao.Text)
        cmdAtualTipo.Enabled = True
    End If
End Sub

Private Sub cboEspecificacao_Change()
    If Len(Trim(cboEspecificacao.Text)) = 0 Then
        cmdAtualEspec.Enabled = False
    End If
End Sub

Private Sub cboEspecificacao_Click()
    If Len(Trim(cboEspecificacao.Text)) > 0 Then
        cmdAtualEspec.Enabled = True
    End If
End Sub

Private Sub cboTipo_Change()
    If Len(Trim(cboTipo.Text)) = 0 Then
        cmdAtualCaract.Enabled = False
        cboEspecificacao.Clear
        Call cboEspecificacao_Change
    End If
End Sub

Private Sub cboTipo_Click()
    cboEspecificacao.Clear
    
    If Len(Trim(cboTipo.Text)) > 0 Then
        Call suListarEspecificacao(cboDivisao.Text, cboTipo.Text)
        cmdAtualCaract.Enabled = True
    End If
End Sub

Private Sub cmdAtualCaract_Click()
    If Len(Trim(cboDivisao.Text)) > 0 And Len(Trim(cboTipo.Text)) > 0 Then
        With frmCaracteristicas
            strSQL = "SELECT * FROM tb_Tipos WHERE Tipo = '" & Trim(cboTipo.Text) & "' AND DivisaoID=" & fnDivisaoID(cboDivisao.Text)
            Set rs = New ADODB.Recordset
            rs.Open strSQL, cn, adOpenKeyset, adLockOptimistic
            
            If rs.EOF = False Then
                .gexCaracteristicas.HoldFields
                Set .gexCaracteristicas.ADORecordset = rs
            End If
            
            Set rs = Nothing
        End With
        Call frmCaracteristicas.Show(vbModal)
    End If
End Sub

Private Sub cmdAtualEspec_Click()
    If Len(Trim(cboDivisao.Text)) > 0 And Len(Trim(cboTipo.Text)) > 0 And Len(Trim(cboEspecificacao.Text)) > 0 Then
        With frmEspecificacoes
            strSQL = "SELECT * FROM tb_Especificacoes WHERE Especificacao = '" & Trim(cboEspecificacao.Text) & "' AND TipoID=" & fnTipoID(cboDivisao.Text, cboTipo.Text)
            Set rs = New ADODB.Recordset
            rs.Open strSQL, cn, adOpenKeyset, adLockOptimistic
            
            If rs.EOF = False Then
                .gexEspecificacoes.HoldFields
                Set .gexEspecificacoes.ADORecordset = rs
            End If
            
            Set rs = Nothing
        End With
        Call frmEspecificacoes.Show(vbModal)
    End If
End Sub

Private Sub cmdAtualTipo_Click()
    If Len(Trim(cboDivisao.Text)) > 0 Then
        With frmTipos
            strSQL = "SELECT * FROM tb_Divisao WHERE Divisao = '" & cboDivisao.Text & "'"
            Set rs = New ADODB.Recordset
            rs.Open strSQL, cn, adOpenKeyset, adLockOptimistic
            
            If rs.EOF = False Then
                .gexTipos.HoldFields
                Set .gexTipos.ADORecordset = rs
            End If
            
            Set rs = Nothing
        End With
        Call frmTipos.Show(vbModal)
    End If
End Sub

Private Sub cmdCadastrar_Click()
Dim i As Integer
Dim blnAchou As Boolean
    
    i = 0
    blnAchou = False
    
    If cboDivisao.Text = "" And cboTipo.Text = "" And cboEspecificacao.Text = "" Then
        MsgBox "Nenhum campo foi preenchido!", vbOKOnly + vbExclamation, "Suporte Manutenção"
        Exit Sub
    End If
    
    If Len(Trim(cboDivisao.Text)) > 0 And Len(Trim(cboTipo.Text)) = 0 And Len(Trim(cboEspecificacao.Text)) = 0 Then
        For i = 0 To cboDivisao.ListCount
            If cboDivisao.Text = cboDivisao.List(i) Then
                blnAchou = True
            End If
        Next
        If blnAchou = False Then
            Call suCadastrarDivisao(Trim(cboDivisao.Text))
        Else
            MsgBox "Tipo já cadastrado!", vbOKOnly + vbExclamation, "Suporte Manutenção"
            Exit Sub
        End If
    ElseIf Len(Trim(cboDivisao.Text)) > 0 And Len(Trim(cboTipo.Text)) > 0 And Len(Trim(cboEspecificacao.Text)) = 0 Then
        For i = 0 To cboTipo.ListCount
            If cboTipo.Text = cboTipo.List(i) Then
                blnAchou = True
            End If
        Next
        If blnAchou = False Then
            Call suCadastrarTipo(Trim(cboDivisao.Text), Trim(cboTipo.Text))
        Else
            MsgBox "Característica já cadastrada!", vbOKOnly + vbExclamation, "Suporte Manutenção"
            Exit Sub
        End If
    ElseIf Len(Trim(cboDivisao.Text)) > 0 And Len(Trim(cboTipo.Text)) > 0 And Len(Trim(cboEspecificacao.Text)) > 0 Then
        For i = 0 To cboEspecificacao.ListCount
            If cboEspecificacao.Text = cboEspecificacao.List(i) Then
                blnAchou = True
            End If
        Next
        If blnAchou = False Then
            Call suCadastrarEspecificacao(Trim(cboDivisao.Text), Trim(cboTipo.Text), Trim(cboEspecificacao.Text))
        Else
            MsgBox "Especificação já cadastrada!", vbOKOnly + vbExclamation, "Suporte Manutenção"
            Exit Sub
        End If
    End If
    
End Sub

Private Sub cmdFechar_Click()
    Call Unload(Me)
End Sub

Private Sub Form_Load()
    Call suListarDivisoes
End Sub

Private Sub suCadastrarEspecificacao(ByVal vDivisao As String, ByVal vTipo As String, ByVal vEspecificacao As String)
Dim intTipoID As Integer
Dim rs1 As ADODB.Recordset

    strSQL = "SELECT * FROM tb_Tipos WHERE DivisaoID = " & fnDivisaoID(vDivisao) & " AND TipoID= " & fnTipoID(vDivisao, vTipo) & ""
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        intTipoID = rs!TipoID
        
        strSQL = "SELECT * FROM tb_Especificacoes WHERE 1=2"
        Set rs1 = New ADODB.Recordset
        rs1.Open strSQL, cn, adOpenKeyset, adLockOptimistic
        
        If rs1.EOF = True Then
            rs1.AddNew
            rs1!Especificacao = vEspecificacao & ""
            rs1!TipoID = intTipoID
            rs1.Update
        End If
        
        rs1.Close
        Set rs1 = Nothing
    End If
    
    rs.Close
    Set rs = Nothing
End Sub

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

Private Sub suCadastrarTipo(ByVal vDivisao As String, ByVal vTipo As String)
    strSQL = "SELECT * FROM tb_Tipos WHERE 1=2"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenKeyset, adLockOptimistic
    
    If rs.EOF = True Then
        rs.AddNew
        rs!Tipo = vTipo & ""
        rs!DivisaoID = fnDivisaoID(vDivisao)
        rs.Update
    End If
    
    rs.Close
    Set rs = Nothing
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

Private Sub suCadastrarDivisao(ByVal vDivisao As String)
    strSQL = "SELECT * FROM tb_Divisao WHERE 1=2"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenKeyset, adLockOptimistic
    
    If rs.EOF = True Then
        rs.AddNew
        rs!Divisao = vDivisao & ""
        rs.Update
    End If
    
    rs.Close
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

Private Sub suListarTipo(ByVal vDivisao As String)
    strSQL = "SELECT * FROM vw_Tipos WHERE Divisao = '" & vDivisao & "'"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    Do While Not rs.EOF
        cboTipo.AddItem rs!Tipo
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
End Sub

Private Sub suListarEspecificacao(ByVal vDivisao As String, ByVal vTipo As String)
    strSQL = "SELECT * FROM vw_Especificacoes WHERE TipoID = " & fnTipoID(vDivisao, vTipo) & ""
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    Do While Not rs.EOF
        cboEspecificacao.AddItem rs!Especificacao
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
End Sub

