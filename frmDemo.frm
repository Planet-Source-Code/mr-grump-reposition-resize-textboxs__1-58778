VERSION 5.00
Begin VB.Form frmDemo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Demo"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   2010
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show"
      Height          =   405
      Left            =   270
      TabIndex        =   2
      Top             =   960
      Width           =   1305
   End
   Begin VB.OptionButton optReposition 
      Caption         =   "Reposition"
      Height          =   195
      Left            =   300
      TabIndex        =   1
      Top             =   240
      Value           =   -1  'True
      Width           =   1425
   End
   Begin VB.OptionButton optResize 
      Caption         =   "Resize"
      Height          =   195
      Left            =   300
      TabIndex        =   0
      Top             =   570
      Width           =   1005
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdShow_Click()
   Form1.Show vbModal
End Sub


