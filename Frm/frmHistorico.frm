VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmHistorico 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Suporte Manutenção - Histórico Situação OS"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12855
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
   ScaleHeight     =   6975
   ScaleWidth      =   12855
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdFechar 
      Appearance      =   0  'Flat
      Caption         =   "Fechar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11400
      TabIndex        =   1
      Top             =   6480
      Width           =   1335
   End
   Begin GridEX20.GridEX gexHistorico 
      Height          =   6375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12855
      _ExtentX        =   22675
      _ExtentY        =   11245
      Version         =   "2.0"
      AllowRowSizing  =   -1  'True
      RecordNavigator =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      HeaderStyle     =   2
      MethodHoldFields=   -1  'True
      AllowColumnDrag =   0   'False
      AllowEdit       =   0   'False
      BorderStyle     =   3
      GroupByBoxVisible=   0   'False
      RowHeaders      =   -1  'True
      HeaderFontName  =   "MS Sans Serif"
      FontName        =   "MS Sans Serif"
      ColumnHeaderHeight=   285
      IntProp1        =   0
      ColumnsCount    =   6
      Column(1)       =   "frmHistorico.frx":0000
      Column(2)       =   "frmHistorico.frx":01A0
      Column(3)       =   "frmHistorico.frx":033C
      Column(4)       =   "frmHistorico.frx":04DC
      Column(5)       =   "frmHistorico.frx":066C
      Column(6)       =   "frmHistorico.frx":0800
      FormatStylesCount=   5
      FormatStyle(1)  =   "frmHistorico.frx":095C
      FormatStyle(2)  =   "frmHistorico.frx":0A88
      FormatStyle(3)  =   "frmHistorico.frx":0B38
      FormatStyle(4)  =   "frmHistorico.frx":0BEC
      FormatStyle(5)  =   "frmHistorico.frx":0CC4
      ImageCount      =   0
      PrinterProperties=   "frmHistorico.frx":0D7C
   End
End
Attribute VB_Name = "frmHistorico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdFechar_Click()
    Call Unload(Me)
End Sub

