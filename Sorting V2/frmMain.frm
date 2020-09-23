VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Sorting Examples"
   ClientHeight    =   7965
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4110
   LinkTopic       =   "Form1"
   ScaleHeight     =   7965
   ScaleWidth      =   4110
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optSort 
      Caption         =   "Option1"
      Height          =   255
      Index           =   12
      Left            =   240
      TabIndex        =   41
      Top             =   7560
      Width           =   255
   End
   Begin VB.OptionButton optSort 
      Caption         =   "Option1"
      Height          =   255
      Index           =   11
      Left            =   2160
      TabIndex        =   39
      Top             =   7200
      Width           =   255
   End
   Begin VB.CheckBox chkTime 
      Caption         =   "Check1"
      Height          =   255
      Left            =   1560
      TabIndex        =   37
      Top             =   3480
      Width           =   255
   End
   Begin VB.OptionButton optSort 
      Caption         =   "Option1"
      Height          =   255
      Index           =   10
      Left            =   240
      TabIndex        =   35
      Top             =   7200
      Width           =   255
   End
   Begin VB.OptionButton optSort 
      Caption         =   "Option1"
      Height          =   255
      Index           =   9
      Left            =   240
      TabIndex        =   34
      Top             =   6840
      Width           =   255
   End
   Begin VB.OptionButton optSort 
      Caption         =   "Option1"
      Height          =   255
      Index           =   8
      Left            =   2160
      TabIndex        =   31
      Top             =   6480
      Width           =   255
   End
   Begin VB.OptionButton optSort 
      Caption         =   "Option1"
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   29
      Top             =   6480
      Width           =   255
   End
   Begin VB.OptionButton optSort 
      Caption         =   "Option1"
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   26
      Top             =   5760
      Width           =   255
   End
   Begin VB.OptionButton optSort 
      Caption         =   "Option1"
      Height          =   255
      Index           =   5
      Left            =   2160
      TabIndex        =   24
      Top             =   5400
      Width           =   255
   End
   Begin VB.OptionButton optSort 
      Caption         =   "Option1"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   22
      Top             =   5400
      Width           =   255
   End
   Begin VB.CommandButton cmdClrSort 
      Caption         =   "Clear Sorted"
      Height          =   375
      Left            =   2880
      TabIndex        =   19
      Top             =   3000
      Width           =   1095
   End
   Begin VB.OptionButton optSort 
      Caption         =   "Option1"
      Height          =   255
      Index           =   3
      Left            =   2160
      TabIndex        =   16
      Top             =   4680
      Width           =   255
   End
   Begin VB.OptionButton optSort 
      Caption         =   "Option1"
      Height          =   255
      Index           =   2
      Left            =   2160
      TabIndex        =   15
      Top             =   4320
      Width           =   255
   End
   Begin VB.OptionButton optSort 
      Caption         =   "Option1"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   12
      Top             =   4680
      Width           =   255
   End
   Begin VB.OptionButton optSort 
      Caption         =   "Option1"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   11
      Top             =   4320
      Value           =   -1  'True
      Width           =   255
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear All"
      Height          =   375
      Left            =   1560
      TabIndex        =   10
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "Sort"
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin VB.ListBox lstSort 
      Height          =   1815
      Left            =   2760
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.ListBox lstList 
      Height          =   1815
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label20 
      Caption         =   "Quick Sort with Bubble Sort"
      Height          =   255
      Left            =   600
      TabIndex        =   42
      Top             =   7560
      Width           =   2055
   End
   Begin VB.Label Label19 
      Caption         =   "Binary Insertion Sort"
      Height          =   255
      Left            =   2520
      TabIndex        =   40
      Top             =   7200
      Width           =   1575
   End
   Begin VB.Label Label18 
      Caption         =   "Timed Sort"
      Height          =   255
      Left            =   1920
      TabIndex        =   38
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label17 
      Caption         =   "Bucket Sort"
      Height          =   255
      Left            =   600
      TabIndex        =   36
      Top             =   7200
      Width           =   975
   End
   Begin VB.Label Label16 
      Caption         =   "Odd-Even Transposition Sort"
      Height          =   255
      Left            =   600
      TabIndex        =   33
      Top             =   6840
      Width           =   2175
   End
   Begin VB.Label Label15 
      Caption         =   "Shaker Sort"
      Height          =   255
      Left            =   2520
      TabIndex        =   32
      Top             =   6480
      Width           =   975
   End
   Begin VB.Label Label14 
      Caption         =   "Radix Sort"
      Height          =   255
      Left            =   600
      TabIndex        =   30
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Caption         =   "Other Sorts"
      Height          =   255
      Left            =   1080
      TabIndex        =   28
      Top             =   6120
      Width           =   1335
   End
   Begin VB.Label Label12 
      Caption         =   "Quick Sort"
      Height          =   255
      Left            =   600
      TabIndex        =   27
      Top             =   5760
      Width           =   975
   End
   Begin VB.Label Label11 
      Caption         =   "Merge Sort"
      Height          =   255
      Left            =   2520
      TabIndex        =   25
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "Heap Sort"
      Height          =   255
      Left            =   600
      TabIndex        =   23
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "O(n log n) Sorts"
      Height          =   255
      Left            =   1200
      TabIndex        =   21
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "O(n^2) Sorts"
      Height          =   255
      Left            =   1440
      TabIndex        =   20
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "Shell Sort"
      Height          =   255
      Left            =   2520
      TabIndex        =   18
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Bubble Sort"
      Height          =   255
      Left            =   2520
      TabIndex        =   17
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4080
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Label Label5 
      Caption         =   "Selection Sort"
      Height          =   255
      Left            =   600
      TabIndex        =   14
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Insertion Sort"
      Height          =   255
      Left            =   600
      TabIndex        =   13
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label lblItems 
      Caption         =   "0 items in list"
      Height          =   615
      Left            =   1560
      TabIndex        =   8
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Sorted List"
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Entry"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "List"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuGen 
      Caption         =   "Generate List"
      Begin VB.Menu mnuRand 
         Caption         =   "Random"
         Begin VB.Menu mnuRnd10 
            Caption         =   "10 Items"
         End
         Begin VB.Menu mnuRnd500 
            Caption         =   "500 Items"
         End
         Begin VB.Menu mnuRnd5k 
            Caption         =   "5000 Items"
         End
         Begin VB.Menu mnuRndCust 
            Caption         =   "Custom Size..."
         End
      End
      Begin VB.Menu mnuRev 
         Caption         =   "Reverse Order"
         Begin VB.Menu mnuRev10 
            Caption         =   "10 Items"
         End
         Begin VB.Menu mnuRev500 
            Caption         =   "500 Items"
         End
         Begin VB.Menu mnuRev5k 
            Caption         =   "5000 Items"
         End
         Begin VB.Menu mnuRevCust 
            Caption         =   "Custom Size..."
         End
      End
      Begin VB.Menu mnuPre 
         Caption         =   "Pre-Sorted"
         Begin VB.Menu mnuPre10 
            Caption         =   "10 Items"
         End
         Begin VB.Menu mnuPre500 
            Caption         =   "500 Items"
         End
         Begin VB.Menu mnuPre5k 
            Caption         =   "5000 Items"
         End
         Begin VB.Menu mnuPreCust 
            Caption         =   "Custom Size..."
         End
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This code is the property of Guillaume Couture-Levesque
'Do not use this code in any other program without the
'permission of the author (me(Guillaume Couture-Levesque))

Option Base 0
Option Explicit

Private Declare Function GetTickCount Lib "kernel32.dll" () As Long

Const Max_Items = 5000
Dim list(Max_Items) As Integer
Dim temp(Max_Items + 1) As Integer
Dim items As Integer

Private Sub cmdAdd_Click()
    Dim entry As Integer
    
    'get entry
    entry = Val(txtEntry.Text)
    
    'check for overflow
    If items = 50 Then
        MsgBox "Maximum number of items reached!", vbInformation, "Entry Error!"
        Exit Sub
    End If
    
    'everything is good
    items = items + 1
    lstList.AddItem (entry)
    lblItems = items & " items in list"
    
    'return to entry
    txtEntry.SetFocus
End Sub

Private Sub cmdClear_Click()
    'clear everything out
    items = 0
    InitList
    lstList.Clear
    lstSort.Clear
    txtEntry.Text = ""
    lblItems.Caption = items & " items in list"
End Sub

Private Sub cmdClrSort_Click()
    'clear out the sorted list and the sorted listbox
    InitList
    lstSort.Clear
End Sub

Private Sub cmdRemove_Click()
    'check for valid selection
    If lstList.Text = "" Then
        MsgBox "You must select an item to remove!", vbInformation, "Selection Error!"
        Exit Sub
    End If
    
    'remove the entry
    items = items - 1
    lstList.RemoveItem (lstList.ListIndex)
    lblItems = items & " items in list"
End Sub

Private Sub cmdSort_Click()
    Dim i As Integer
    Dim start As Long
    Dim finish As Long
    
    'make sure there's enough items
    If items <= 1 Then
        MsgBox "Not enough items to sort!", vbInformation, "Array Formation Error!"
        Exit Sub
    End If
    
    'make the list
    FormList
    
    'start timer
    If chkTime.Value Then
        start = GetTickCount
    End If
    
    'sort the list
    If optSort(0).Value Then
        InsertionSort list(), items
    End If
    If optSort(1).Value Then
        SelectionSort list(), items
    End If
    If optSort(2).Value Then
        BubbleSort list(), items
    End If
    If optSort(3).Value Then
        ShellSort list(), items
    End If
    If optSort(4).Value Then
        HeapSort list(), items
    End If
    If optSort(5).Value Then
        MergeSort list(), temp(), items
    End If
    If optSort(6).Value Then
        QuickSort list(), items
    End If
    If optSort(7).Value Then
        RadixSort list(), temp(), items
    End If
    If optSort(8).Value Then
        ShakerSort list(), items
    End If
    If optSort(9).Value Then
        OETSort list(), items
    End If
    If optSort(10).Value Then
        'check for above 0
        For i = 0 To items - 1
            If list(i) <= 0 Then
                MsgBox "Bucket sort is only for numbers greater than 0!", vbInformation, "Sort Error!"
                Exit Sub
            End If
        Next i
        'check for lower than n
        For i = 0 To items - 1
            If list(i) > items Then
                MsgBox "Bucket sort requires that every number be smaller than or equal to the number of items!", vbInformation, "Sort Error!"
                Exit Sub
            End If
        Next i
        BucketSort list(), items, temp()
    End If
    If optSort(11).Value Then
        BinaryInsertionSort list(), items
    End If
    If optSort(12).Value Then
        BubbleQuickSort list(), items
    End If
    
    'stop timer and show result
    If chkTime.Value Then
        finish = GetTickCount
        MsgBox "Time taken: " & (finish - start) & " milliseconds (thousandths of a second).", vbOKOnly, "Time Result"
    End If
    
    'display the list
    DisplayList
End Sub

Private Sub FormList()
    Dim i As Integer
    
    'get the value from the list and add it to the array
    For i = 0 To (items - 1)
        list(i) = Val(lstList.list(i))
    Next i
End Sub

Private Sub DisplayList()
    Dim i As Integer
    
    'clear the sorted list
    lstSort.Clear
    
    'add all of the values to the sorted list
    For i = 0 To (items - 1)
        lstSort.AddItem (list(i))
    Next i
End Sub

Private Sub Form_Load()
    'init form
    Me.Show
    DoEvents
    
    'init variables
    items = 0
    InitList
    
    Debug.Print ""
    Debug.Print "Start of New Run"
End Sub

Private Sub InitList()
    Dim i As Integer
    
    'zero out the array
    For i = 0 To (Max_Items - 1)
        list(i) = 0
    Next i
    
    'zero out the temporary array
    For i = 0 To Max_Items
        temp(i) = 0
    Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim Form As Form
    
    'proper multiple unload
    For Each Form In Forms
        Unload Form
        Set Form = Nothing
    Next Form
End Sub

Private Sub mnuAbout_Click()
    'show the about box
    MsgBox "This is the an exhibition of several sorting methods. " & vbCrLf & _
    "There are several more sorting methods that will be added eventually. " & vbCrLf & _
    "The Odd-Even Transposition Sort (OETS) although intended for parallel " & vbCrLf & _
    "processing, is only going to run in 1 thread on 1 processor, and has " & vbCrLf & _
    "been rewritten accordingly." & vbCrLf & vbCrLf & "Written by Guillaume Couture-Levesque" & _
    vbCrLf & "July 22nd 2003", vbInformation, "About This Program"
End Sub

Private Sub mnuExit_Click()
    'exit program
    Unload Me
End Sub

Private Sub mnuPre10_Click()
    Dim i As Integer
    
    'generate numbers
    GenerateSorted list(), 1, 10, 10
    
    'update listbox
    lstList.Clear
    For i = 0 To 9
        lstList.AddItem list(i)
    Next i
    items = lstList.ListCount
    lblItems.Caption = items & " items in list"
End Sub

Private Sub mnuPre500_Click()
    Dim i As Integer
    
    'generate numbers
    GenerateSorted list(), 1, 500, 500
    
    'update listbox
    lstList.Clear
    For i = 0 To 499
        lstList.AddItem list(i)
    Next i
    items = lstList.ListCount
    lblItems.Caption = items & " items in list"
End Sub

Private Sub mnuPre5k_Click()
    Dim i As Integer
    
    'generate numbers
    GenerateSorted list(), 1, 5000, 5000
    
    'update listbox
    lstList.Clear
    For i = 0 To 4999
        lstList.AddItem list(i)
    Next i
    items = lstList.ListCount
    lblItems.Caption = items & " items in list"
End Sub

Private Sub mnuPreCust_Click()
    Dim size As Integer
    Dim min As Integer
    Dim max As Integer
    Dim i As Integer
    
    'show dialog
    frmSize.Visible = True
    Do While frmSize.Visible
        DoEvents
    Loop
    
    'generate list
    size = Val(frmSize.txtNum.Text)
    If (size = 0) Then
        Exit Sub
    End If
    min = Val(frmSize.txtMin.Text)
    max = Val(frmSize.txtMax.Text)
    GenerateSorted list(), min, max, size
    
    'update listbox
    lstList.Clear
    For i = 0 To (size - 1)
        lstList.AddItem list(i)
    Next i
    items = lstList.ListCount
    lblItems.Caption = items & " items in list"
End Sub

Private Sub mnuRev10_Click()
    Dim i As Integer
    
    'generate numbers
    GenerateReverse list(), 1, 10, 10
    
    'update listbox
    lstList.Clear
    For i = 0 To 9
        lstList.AddItem list(i)
    Next i
    items = lstList.ListCount
    lblItems.Caption = items & " items in list"
End Sub

Private Sub mnuRev500_Click()
    Dim i As Integer
    
    'generate numbers
    GenerateReverse list(), 1, 500, 500
    
    'update listbox
    lstList.Clear
    For i = 0 To 499
        lstList.AddItem list(i)
    Next i
    items = lstList.ListCount
    lblItems.Caption = items & " items in list"
End Sub

Private Sub mnuRev5k_Click()
    Dim i As Integer
    
    'generate numbers
    GenerateReverse list(), 1, 5000, 5000
    
    'update listbox
    lstList.Clear
    For i = 0 To 4999
        lstList.AddItem list(i)
    Next i
    items = lstList.ListCount
    lblItems.Caption = items & " items in list"
End Sub

Private Sub mnuRevCust_Click()
    Dim size As Integer
    Dim min As Integer
    Dim max As Integer
    Dim i As Integer
    
    'show dialog
    frmSize.Visible = True
    Do While frmSize.Visible
        DoEvents
    Loop
    
    'generate list
    size = Val(frmSize.txtNum.Text)
    If (size = 0) Then
        Exit Sub
    End If
    min = Val(frmSize.txtMin.Text)
    max = Val(frmSize.txtMax.Text)
    GenerateReverse list(), min, max, size
    
    'update listbox
    lstList.Clear
    For i = 0 To (size - 1)
        lstList.AddItem list(i)
    Next i
    items = lstList.ListCount
    lblItems.Caption = items & " items in list"
End Sub

Private Sub mnuRnd10_Click()
    Dim i As Integer
    
    'generate numbers
    GenerateRandom list(), 1, 10, 10
    
    'update listbox
    lstList.Clear
    For i = 0 To 9
        lstList.AddItem list(i)
    Next i
    items = lstList.ListCount
    lblItems.Caption = items & " items in list"
End Sub

Private Sub mnuRnd500_Click()
    Dim i As Integer
    
    'generate numbers
    GenerateRandom list(), 1, 500, 500
    
    'update listbox
    lstList.Clear
    For i = 0 To 499
        lstList.AddItem list(i)
    Next i
    items = lstList.ListCount
    lblItems.Caption = items & " items in list"
End Sub

Private Sub mnuRnd5k_Click()
    Dim i As Integer
    
    'generate numbers
    GenerateRandom list(), 1, 5000, 5000
    
    'update listbox
    lstList.Clear
    For i = 0 To 4999
        lstList.AddItem list(i)
    Next i
    items = lstList.ListCount
    lblItems.Caption = items & " items in list"
End Sub

Private Sub mnuRndCust_Click()
    Dim size As Integer
    Dim min As Integer
    Dim max As Integer
    Dim i As Integer
    
    'show dialog
    frmSize.Visible = True
    Do While frmSize.Visible
        DoEvents
    Loop
    
    'generate list
    size = Val(frmSize.txtNum.Text)
    If (size = 0) Then
        Exit Sub
    End If
    min = Val(frmSize.txtMin.Text)
    max = Val(frmSize.txtMax.Text)
    GenerateRandom list(), min, max, size
    
    'update listbox
    lstList.Clear
    For i = 0 To (size - 1)
        lstList.AddItem list(i)
    Next i
    items = lstList.ListCount
    lblItems.Caption = items & " items in list"
End Sub

Private Sub txtEntry_KeyPress(KeyAscii As Integer)
    'check for enter
    If KeyAscii = 13 Then
        cmdAdd_Click
    End If
End Sub
