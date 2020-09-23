VERSION 5.00
Begin VB.Form frmSize 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Custom Size"
   ClientHeight    =   1425
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtMax 
      Height          =   285
      Left            =   1920
      TabIndex        =   7
      Text            =   "0"
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox txtMin 
      Height          =   285
      Left            =   4080
      TabIndex        =   5
      Text            =   "0"
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox txtNum 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Text            =   "0"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "and a maximum value of"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "items, a minimum value of"
      Height          =   255
      Left            =   2160
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Create a list with"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmSize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This code is the property of Guillaume Couture-Levesque
'Do not use this code in any other program without the
'permission of the author (me(Guillaume Couture-Levesque))

Option Explicit

Private Sub CancelButton_Click()
    'reset values
    txtNum.Text = 0
    txtMin.Text = 0
    txtMax.Text = 0
    frmSize.Visible = False
End Sub

Private Sub OKButton_Click()
    'check for max and min
    If Val(txtMin.Text) > Val(txtMax.Text) Then
        MsgBox "Min must be greater than max!", vbInformation, "List Generation Error!"
        Exit Sub
    End If
    
    'check for num
    If Val(txtNum.Text) <= 1 Then
        MsgBox "You must have more than 1 items!", vbInformation, "List Generation Error!"
        Exit Sub
    End If
    
    'check for size
    If Val(txtNum.Text) > 5000 Then
        MsgBox "You must have 5000 items or less!", vbInformation, "List Generation Error!"
        Exit Sub
    End If
    
    'all is well
    frmSize.Visible = False
End Sub
