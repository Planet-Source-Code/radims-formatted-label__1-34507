VERSION 5.00
Begin VB.Form fMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FormattedLabel"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   4290
   StartUpPosition =   2  'CenterScreen
   Begin Demo.ucFormattedLabel ucFormattedLabel1 
      Height          =   1890
      Left            =   165
      Top             =   165
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   3334
      ForeColor       =   -2147483630
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   $"fMain.frx":0000
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   $"fMain.frx":0176
      Height          =   1725
      Left            =   180
      TabIndex        =   0
      Top             =   2265
      Width           =   3900
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub txtCaption_Change()
    
    Me.ucFormattedLabel1.Caption = Me.txtCaption
    
End Sub
