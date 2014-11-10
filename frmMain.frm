VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "mschrt20.ocx"
Begin VB.Form frmMain 
   Caption         =   "Sorting Algorithms"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8685
   LinkTopic       =   "Form1"
   ScaleHeight     =   10215
   ScaleWidth      =   8685
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtResults 
      Height          =   285
      Left            =   5040
      TabIndex        =   20
      Top             =   5400
      Width           =   3015
   End
   Begin MSChart20Lib.MSChart MSChart 
      Height          =   3975
      Left            =   240
      OleObjectBlob   =   "frmMain.frx":0000
      TabIndex        =   18
      Top             =   6120
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.TextBox txtArray 
      Height          =   285
      Left            =   480
      TabIndex        =   16
      Text            =   "1000"
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   6960
      TabIndex        =   15
      Top             =   9720
      Width           =   1455
   End
   Begin VB.CheckBox Check 
      Caption         =   "Show Graphic and explanation"
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   14
      Top             =   2400
      Width           =   2775
   End
   Begin VB.CheckBox Check 
      Caption         =   "Binary Search"
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   13
      Top             =   2040
      Width           =   2655
   End
   Begin VB.OptionButton OptSort 
      Caption         =   "Quick Sort"
      Height          =   255
      Index           =   8
      Left            =   5040
      TabIndex        =   12
      Top             =   3960
      Width           =   2655
   End
   Begin VB.OptionButton OptSort 
      Caption         =   "Merge Sort"
      Height          =   255
      Index           =   7
      Left            =   5040
      TabIndex        =   11
      Top             =   3600
      Width           =   2895
   End
   Begin VB.OptionButton OptSort 
      Caption         =   "Jump Sort"
      Height          =   255
      Index           =   6
      Left            =   5040
      TabIndex        =   10
      Top             =   3240
      Width           =   2775
   End
   Begin VB.OptionButton OptSort 
      Caption         =   "Insertion Sort"
      Height          =   255
      Index           =   5
      Left            =   5040
      TabIndex        =   9
      Top             =   2880
      Width           =   2775
   End
   Begin VB.OptionButton OptSort 
      Caption         =   "Heap Sort"
      Height          =   255
      Index           =   4
      Left            =   480
      TabIndex        =   8
      Top             =   4440
      Width           =   2655
   End
   Begin VB.OptionButton OptSort 
      Caption         =   "Gnome Sort"
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   7
      Top             =   4080
      Width           =   2535
   End
   Begin VB.OptionButton OptSort 
      Caption         =   "Comb Sort"
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   6
      Top             =   3720
      Width           =   2535
   End
   Begin VB.OptionButton OptSort 
      Caption         =   "Cocktail Sort"
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   5
      Top             =   3360
      Width           =   2535
   End
   Begin VB.OptionButton OptSort 
      Caption         =   "Bubble Sort"
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   4
      Top             =   3000
      Value           =   -1  'True
      Width           =   2535
   End
   Begin VB.CheckBox chkTimeSort 
      Caption         =   "Include sort in time"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   1680
      Width           =   2535
   End
   Begin VB.TextBox txtTime 
      Height          =   285
      Left            =   480
      TabIndex        =   2
      Top             =   5370
      Width           =   3015
   End
   Begin VB.CommandButton cmdSorted 
      Caption         =   "Run Sorted"
      Height          =   735
      Left            =   480
      TabIndex        =   1
      Top             =   240
      Width           =   3255
   End
   Begin VB.CommandButton cmdUnsort 
      Caption         =   "Run Unsorted"
      Height          =   735
      Left            =   5040
      TabIndex        =   0
      Top             =   240
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   "Calc Result"
      Height          =   255
      Left            =   5040
      TabIndex        =   22
      Top             =   4920
      Width           =   2655
   End
   Begin VB.Label Label 
      Caption         =   "Time Taken"
      Height          =   255
      Left            =   480
      TabIndex        =   21
      Top             =   4920
      Width           =   2655
   End
   Begin VB.Label lblExp 
      Height          =   3375
      Left            =   5160
      TabIndex        =   19
      Top             =   6120
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "Array Size"
      Height          =   255
      Left            =   1560
      TabIndex        =   17
      Top             =   1200
      Width           =   1695
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private dtStart As Long
Private dtEnd As Long
Private srtArray() As Integer

Private Sub cmdUnsort_Click()
    Dim arraySize As Integer
    
    arraySize = txtArray.Text
    
    ReDim j(0 To arraySize) As Integer
    Dim b As Integer
    Dim l As Integer
    Dim count As Long
    Dim Item As Integer
    Dim sw As New StopWatch
    MSChart.Visible = False
    
    b = 0
    count = 0
    
    While b < txtArray.Text
        j(b) = Round(Math.Rnd * 256)
        b = b + 1
    Wend
    
    sw.start

    For l = 0 To 500
        For Each Item In j
            If Item < 50 Then
                count = count + 1
            End If
        Next
    Next
    
    txtTime.Text = sw.ElapsedMilliseconds
    txtResults.Text = count
End Sub

Private Sub cmdSorted_Click()
    Dim arraySize As Integer
    
    arraySize = txtArray.Text

    ReDim srtArray(0 To arraySize) As Integer
    Dim c As Integer
    Dim l As Integer
    Dim X As Integer
    Dim count As Long
    Dim itm As Integer
    Dim sw As New StopWatch
    Dim showGraph As Boolean
    showGraph = False
    
    c = 0
    count = 0
    
    While c <= txtArray.Text
        srtArray(c) = Round(Math.Rnd * 256)
        c = c + 1
    Wend
    
    If Check(1).Value = 1 Then
        MSChart.ChartData = srtArray
        For X = 1 To arraySize + 1
            MSChart.Plot.SeriesCollection(X).DataPoints(-1).Brush.FillColor.Set 0, 0, 0
        Next
        MSChart.Visible = True
        showGraph = True
    End If
        
    If chkTimeSort.Value = 1 Then
        sw.start
    End If
    
    If showGraph Then
        If OptSort(0).Value Then
            lblExp.Caption = "Bubble Sort - steps through each element and swaps with the next one if required, keeps iterating until no swaps are left"
            Call GrphBubbleSort(srtArray)
        ElseIf OptSort(1).Value Then
            lblExp.Caption = "Cocktail Sort - similar to bubble search but forward and backward through a collection"
            Call GrphCocktailSort(srtArray)
        ElseIf OptSort(2).Value Then
            lblExp.Caption = "CombSort - similar to bubble sort but compares values from one end to the other intially, and swaps them if necessary slowly making the distance it compares smaller by x / 1.3 each iteration. Is a true bubble sort at the end But fairly effecient by this point"
            Call GrphCombSort(srtArray)
        ElseIf OptSort(3).Value Then
            lblExp.Caption = "Gnome Sort - swaps the first out of order elements and checks backwards to see if a swap should be made before going forward again"
            Call GrphGnomeSort(srtArray)
        ElseIf OptSort(4).Value Then
            lblExp.Caption = "Heap sort - Uses two arrays"
            Call GrphHeapSort(srtArray)
        ElseIf OptSort(5).Value Then
            lblExp.Caption = "Insertion sort - every iteration of an insertion sort removes an element from the input data, inserting it at the correct position in the already sorted list"
            Call GrphInsertionSort(srtArray)
        ElseIf OptSort(6).Value Then
            lblExp.Caption = "Similar to Bubble sort as it uses nested loops - but moves elements very far initially - iteratively reducing the distance values are moved"
            Call GrphJumpSort(srtArray)
        ElseIf OptSort(7).Value Then
            lblExp.Caption = "Merge sort "
            Call GrphMergeSort(srtArray)
        ElseIf OptSort(8).Value Then
            lblExp.Caption = "QuickSort - To partition an array, a pivot element is first randomly selected, and then compared against every other element. All smaller elements are moved before the pivot, and all larger elements are moved after. The lesser and greater sublists are then recursively processed until the entire list is sorted. This can be done efficiently in linear time and in-place."
            Call GrphQuickSort(srtArray)
        End If
    Else
        If OptSort(0).Value Then
            Call BubbleSort(srtArray)
        ElseIf OptSort(1).Value Then
            Call CocktailSort(srtArray)
        ElseIf OptSort(2).Value Then
            Call CombSort(srtArray)
        ElseIf OptSort(3).Value Then
            Call GnomeSort(srtArray)
        ElseIf OptSort(4).Value Then
            Call HeapSort(srtArray)
        ElseIf OptSort(5).Value Then
            Call InsertionSort(srtArray)
        ElseIf OptSort(6).Value Then
            Call JumpSort(srtArray)
        ElseIf OptSort(7).Value Then
            Call MergeSort(srtArray)
        ElseIf OptSort(8).Value Then
            Call QuickSort(srtArray)
        End If
    End If
    
    If chkTimeSort.Value = 0 Then
        sw.start
    End If
    
    If Check(0).Value = 1 Then
        For l = 0 To 500
            count = count + BinarySearch(srtArray, 50)
        Next
        
    Else
        For l = 0 To 500
            For Each itm In srtArray
                If itm < 50 Then
                    count = count + 1
                End If
            Next
        Next
    End If
    
    txtTime.Text = sw.ElapsedMilliseconds
    txtResults.Text = count

End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

