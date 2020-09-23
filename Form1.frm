VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2310
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10545
   LinkTopic       =   "Form1"
   ScaleHeight     =   2310
   ScaleWidth      =   10545
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   8040
      TabIndex        =   5
      Text            =   "Text6"
      Top             =   960
      Width           =   2415
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   6024
      TabIndex        =   4
      Text            =   "Text5"
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   4830
      TabIndex        =   3
      Text            =   "Text4"
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2592
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1176
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label1 
      Height          =   585
      Left            =   1815
      TabIndex        =   6
      Top             =   1620
      Width           =   6480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sgSize As Single
Dim sgPercent(1 To 6) As Single
Dim sgSpace As Single

Private Sub Form_Load()
   If frmDemo.optResize.Value Then
   
    
      ' get all the space used by the textboxes (the original size)
      sgSize = Text1.Width + Text2.Width + Text3.Width + Text4.Width + Text5.Width + Text6.Width
      
      ' calculate the percent of each textbox space from the total space
      ' this is used to keep the original proportions when resizing
      sgPercent(1) = Text1.Width / sgSize
      sgPercent(2) = Text2.Width / sgSize
      sgPercent(3) = Text3.Width / sgSize
      sgPercent(4) = Text4.Width / sgSize
      sgPercent(5) = Text5.Width / sgSize
      sgPercent(6) = Text6.Width / sgSize
      
      ' get the size of the space between 2 textboxes (first 2) - this will be fix
      sgSpace = Text2.Left - Text1.Left - Text1.Width
   End If
End Sub

Private Sub Form_Resize()
   If frmDemo.optReposition.Value Then
      RepositionBoxes
   Else
      ResizeBoxes
   End If
End Sub

' this function will reposition the textboxes without changing their width
Private Sub RepositionBoxes()
On Error Resume Next
Dim sgPosition As Single

Label1.Caption = "Textboxes are repositioned without changing their widths.  Click the forms Maximize button to see the effect.  You can also resize the form by dragging the left or right sides.  This works well for applications that may be installed at different screen resolutions."

   ' get all the space used by the textboxes
   sgSize = Text1.Width + Text2.Width + Text3.Width + Text4.Width + Text5.Width + Text6.Width
   
   ' there are 6 textboxes so between them there are 5 spaces
   ' to get the size of a space just divide the remaining space (substracting from the form available width
   ' the total space of the textboxes) by 5
   sgSpace = ((Me.ScaleWidth - 200) - sgSize) / 5 ' -200 is to put a small space between the last text box and the edge of the form.  Make it twice the size of the value used for sgPosition (i.e. 100) below to keep the spaces equal.
   
   ' in case it is not enough space just use 0 between textboxes
   If sgSpace < 0 Then sgSpace = 0
   
   ' start form 0 and position the first textbox
   sgPosition = 100 '100 is used instead of zero to put a small space between the first text box and the edge of the form.  Make it half of the value used for sgSpace (i.e. 200) above to make the spaces equal.
   Text1.Left = sgPosition
   
   ' get the new position = the old one + a space + the width of the previous textbox
   sgPosition = sgPosition + sgSpace + Text1.Width
   Text2.Left = sgPosition
   
   ' get the new position = the old one + a space + the width of the previous textbox
   sgPosition = sgPosition + sgSpace + Text2.Width
   Text3.Left = sgPosition
   
   ' get the new position = the old one + a space + the width of the previous textbox
   sgPosition = sgPosition + sgSpace + Text3.Width
   Text4.Left = sgPosition
   
   ' get the new position = the old one + a space + the width of the previous textbox
   sgPosition = sgPosition + sgSpace + Text4.Width
   Text5.Left = sgPosition
   
   ' get the new position = the old one + a space + the width of the previous textbox
   sgPosition = sgPosition + sgSpace + Text5.Width
   Text6.Left = sgPosition
End Sub

' this function will resize the textboxes without resizing the gaps
Private Sub ResizeBoxes()
On Error Resume Next
Dim sgTotalSpace As Single
Dim sgPosition As Single

Label1.Caption = "Textboxes will be resized and textbox spacing remains the same.  Click the forms Maximize button to see the effect.  You can also resize the form by dragging the left or right edge.  This works well for applications that may be installed at different screen resolutions."

   ' get all the space used by the textboxes
   sgSize = Text1.Width + Text2.Width + Text3.Width + Text4.Width + Text5.Width + Text6.Width
       
   ' there are 6 textboxes so between them there are 5 spaces
   ' get the total free space = the available form width - the space used by the textboxes now - the space used by the gaps
   sgTotalSpace = (Me.ScaleWidth - 200) - sgSize - 5 * sgSpace
    
   ' start form 0 and position the first textbox
   sgPosition = 100
   Text1.Left = sgPosition
   ' resize the textbox adding the corresponding amount of space from the total free space (to keep the original proportions)
   Text1.Width = Text1.Width + sgPercent(1) * sgTotalSpace
   
   ' get the new position = the old one + a space + the width of the previous textbox
   sgPosition = sgPosition + sgSpace + Text1.Width
   Text2.Left = sgPosition
   ' resize the textbox adding the corresponding amount of space from the total free space (to keep the original proportions)
   Text2.Width = Text2.Width + sgPercent(2) * sgTotalSpace
   
   ' get the new position = the old one + a space + the width of the previous textbox
   sgPosition = sgPosition + sgSpace + Text2.Width
   Text3.Left = sgPosition
   ' resize the textbox adding the corresponding amount of space from the total free space (to keep the original proportions)
   Text3.Width = Text3.Width + sgPercent(3) * sgTotalSpace
   
   ' get the new position = the old one + a space + the width of the previous textbox
   sgPosition = sgPosition + sgSpace + Text3.Width
   Text4.Left = sgPosition
   ' resize the textbox adding the corresponding amount of space from the total free space (to keep the original proportions)
   Text4.Width = Text4.Width + sgPercent(4) * sgTotalSpace
   
   ' get the new position = the old one + a space + the width of the previous textbox
   sgPosition = sgPosition + sgSpace + Text4.Width
   Text5.Left = sgPosition
   ' resize the textbox adding the corresponding amount of space from the total free space (to keep the original proportions)
   Text5.Width = Text5.Width + sgPercent(5) * sgTotalSpace
   
   ' get the new position = the old one + a space + the width of the previous textbox
   sgPosition = sgPosition + sgSpace + Text5.Width
   Text6.Left = sgPosition
   ' resize the textbox adding the corresponding amount of space from the total free space (to keep the original proportions)
   Text6.Width = Text6.Width + sgPercent(6) * sgTotalSpace
End Sub
