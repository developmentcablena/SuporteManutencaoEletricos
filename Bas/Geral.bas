Attribute VB_Name = "Geral"
Option Explicit

'Strings da conexão
Public cn As ADODB.Connection
Public strBD As String

'Strings do usuário
Global gintUsuarioID As Integer
Global gstrNome As String
Global gstrDepto As String
Global gstrEMail As String
Global gstrSenha As String
Global gstrTela As String
Global gblnEncaminharOS As Boolean
Global gblnSupervisor As Boolean

'Strings de OS
Global gintOSID As Long

'Variáveis para criação de usuário
Global gintNovoUsuarioID As Integer

'Variáveis do frmChamados
'Combo status
Global gstrcboStatus As String
Global gblnTelaChamados As Boolean
Global gintStatusFinalizada As Integer

Public Sub ConectarBD()
    'Conexão com a Telecom
    'strBD = "PROVIDER=SQLOLEDB;SERVER=192.168.0.249\SQLExpress;DATABASE=SuporteManutencao;UID=cablena_user;PWD=C@bl3na;"
    'strBD = "PROVIDER=SQLOLEDB;SERVER=192.168.0.7;DATABASE=suportemanutencao_teste;UID=helpdesk;PWD=H3lpD3sk;"
    'Conexão com a Elétricos
    strBD = "PROVIDER=SQLOLEDB;SERVER=196.200.80.10;DATABASE=SuporteManutencao;UID=cablena_user;PWD=C@bl3na;"
    Set cn = New ADODB.Connection
    cn.CursorLocation = adUseClient
    cn.Open (strBD)
End Sub

Public Function CodDec(ByVal vText As String) As String
On Error Resume Next
Dim strAux As String
Dim intCount As Long
Dim strText As String
    
    strText = vText
    strAux = ""
    For intCount = 1 To Len(strText)
        strAux = Chr((Asc(Mid(strText, intCount, 1)) Xor 255)) & strAux
    Next
    CodDec = strAux
End Function

Public Function fnTempoParadaEmHoras(ByVal vTempo As Long) As String
    Dim tempoParadaEmHoras As String
    Dim extrairHora As Variant
    Dim extrairMinuto As Variant
    Dim extrairSegundo As Variant
    Dim hora As String
    Dim minuto As String
    Dim segundo As String
    
    extrairHora = Split((vTempo / 3600), ",")
    hora = IIf(Len(extrairHora(0)) = 1, "0" & extrairHora(0), extrairHora(0))
    
    extrairMinuto = Split(((vTempo Mod 3600) / 60), ",")
    minuto = IIf(Len(extrairMinuto(0)) = 1, "0" & extrairMinuto(0), extrairMinuto(0))
    
    extrairSegundo = Split((vTempo Mod 60), ",")
    segundo = IIf(Len(extrairSegundo(0)) = 1, "0" + extrairSegundo(0), extrairSegundo(0))
    
    tempoParadaEmHoras = hora + ":" + minuto + ":" + segundo
    
    fnTempoParadaEmHoras = tempoParadaEmHoras
End Function

Public Function fnTempoParadaEmSegundos(ByVal vTempo As String) As Long
    Dim extrairHoraMinutoSegundo As Variant
    Dim hora As Long
    Dim minuto As Long
    Dim segundo As Long
    Dim tempoParadaEmSegundos As Long
    
    extrairHoraMinutoSegundo = Split(vTempo, ":")
    hora = extrairHoraMinutoSegundo(0) * 3600
    minuto = extrairHoraMinutoSegundo(1) * 60
    segundo = extrairHoraMinutoSegundo(2)
    
    tempoParadaEmSegundos = (hora + minuto + segundo)
        
    fnTempoParadaEmSegundos = tempoParadaEmSegundos
End Function
