VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} acrOS 
   Caption         =   "Suporte Manutenção - Ordem de Serviço"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   Icon            =   "acrOS.dsx":0000
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   21167
   _ExtentY        =   19420
   SectionData     =   "acrOS.dsx":076A
End
Attribute VB_Name = "acrOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rs As ADODB.Recordset
Private strSQL As String

Private Sub suGerarOS(ByVal vOSID As Long)
    strSQL = "SELECT * FROM vw_Chamados WHERE OSID = " & vOSID
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        If gintStatusFinalizada = 0 Then
            With acrOS
                .fldOSID = Format(rs!OSID, "0000")
                .fldDataCadastro.Text = Format(rs!DataCadastro, "dd/MM/yyyy HH:mm")
                .fldDataPrevisao.Text = Format(rs!Previsao, "dd/MM/yyyy")
                .fldSolicitante.Text = rs!Nome
                .fldDepartamento.Text = rs!Departamento
                .fldPrioridade.Text = IIf(rs!Prioridade = True, "Sim", "Não")
                .fldDivisao.Text = rs!Divisao
                .fldTipo.Text = rs!Tipo
                .fldEspecificacao.Text = rs!Linha
                .fldEspec.Text = rs!Especificacao
                .fldDescricao.Text = rs!DescricaoServico
                .fldAtendente.Text = rs!Atendente
            End With
        Else
            With acrOS
                .fldOSID = Format(rs!OSID, "0000")
                .fldDataCadastro.Text = Format(rs!DataCadastro, "dd/MM/yyyy HH:mm")
                .fldDataPrevisao.Text = Format(rs!Previsao, "dd/MM/yyyy")
                .fldSolicitante.Text = rs!Nome
                .fldDepartamento.Text = rs!Departamento
                .fldPrioridade.Text = IIf(rs!Prioridade = True, "Sim", "Não")
                .fldDivisao.Text = rs!Divisao
                .fldTipo.Text = rs!Tipo
                .fldEspecificacao.Text = rs!Linha
                .fldEspec.Text = rs!Especificacao
                .fldDescricao.Text = rs!DescricaoServico
                .fldAtendente.Text = rs!Atendente
                .fldReporte.Text = rs!ReporteTecnico
                .lblData.Caption = Format(rs!DataInicio, "dd/MM/yyyy HH:mm") & ""
                .Label7.Caption = Format(rs!DataBaixa, "dd/MM/yyyy HH:mm") & ""
                
                If IsNull(rs!TempoParadaSegundos) Then
                    Dim tempoParada As Variant
                    Dim hora As String
                    Dim minuto As String
                    Dim extrairSegundo As Variant
                    Dim segundo As String
                    
                    tempoParada = Split(rs!tempoParada, ":")
                    hora = tempoParada(0)
                    minuto = tempoParada(1)
                    extrairSegundo = Split(tempoParada(2), ".")
                    segundo = extrairSegundo(0)
                    
                    .Label9.Caption = hora & ":" & minuto & ":" & segundo & ""
                Else
                    .Label9.Caption = fnTempoParadaEmHoras(rs!TempoParadaSegundos)
                End If
            End With
        End If
    End If
    
    rs.Close
    Set rs = Nothing
End Sub

Private Sub ActiveReport_Activate()
    Call suGerarOS(gintOSID)
End Sub
