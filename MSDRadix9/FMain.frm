VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FMain 
   AutoRedraw      =   -1  'True
   Caption         =   "String Sort Benchmark"
   ClientHeight    =   8730
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7710
   LinkTopic       =   "Form1"
   ScaleHeight     =   8730
   ScaleWidth      =   7710
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraBenchmarksShown 
      Caption         =   "Show Benchmarks"
      Height          =   855
      Left            =   120
      TabIndex        =   27
      Top             =   7320
      Width           =   7455
      Begin VB.CheckBox chkUlliRadix 
         Caption         =   "Ulli's Radix"
         Height          =   255
         Left            =   4920
         TabIndex        =   34
         Top             =   240
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox chkUlliQuickie 
         Caption         =   "Ulli's Quickie"
         Height          =   255
         Left            =   3360
         TabIndex        =   33
         Top             =   480
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox chkStrSwapStable 
         Caption         =   "StrSwapStable"
         Height          =   255
         Left            =   3360
         TabIndex        =   32
         Top             =   240
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox chkTwoArray 
         Caption         =   "Two Array"
         Height          =   255
         Left            =   1920
         TabIndex        =   31
         Top             =   480
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox chkAmericanFlag 
         Caption         =   "American Flag"
         Height          =   255
         Left            =   1920
         TabIndex        =   30
         Top             =   240
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox chkTriQuickSortNR 
         Caption         =   "Tri-Quick Sort NR"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   480
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox chkTriQuicksort 
         Caption         =   "Tri-Quick Sort"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Value           =   1  'Checked
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdExportResults 
      Caption         =   "Export Results"
      Height          =   375
      Left            =   5040
      TabIndex        =   26
      Top             =   8280
      Width           =   1215
   End
   Begin VB.TextBox txtTrials 
      Height          =   285
      Left            =   6720
      TabIndex        =   24
      Top             =   6120
      Width           =   855
   End
   Begin MSComDlg.CommonDialog cmndlg 
      Left            =   7200
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.OptionButton optFileBenchmark 
      Caption         =   "File Benchmark"
      Height          =   255
      Left            =   3360
      TabIndex        =   15
      Top             =   8280
      Width           =   1455
   End
   Begin VB.OptionButton optLineBenchmark 
      Caption         =   "Line Benchmark"
      Height          =   255
      Left            =   1800
      TabIndex        =   14
      Top             =   8280
      Width           =   1815
   End
   Begin VB.OptionButton optSimpleBenchmark 
      Caption         =   "Simple Benchmark"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   8280
      Value           =   -1  'True
      Width           =   1815
   End
   Begin VB.Frame fraSimpleBenchmark 
      Caption         =   "Simple Benchmark"
      Height          =   1335
      Left            =   120
      TabIndex        =   12
      Top             =   5880
      Width           =   2175
      Begin VB.TextBox txtCount 
         Height          =   285
         Left            =   1200
         TabIndex        =   21
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtStringLengthSimple 
         Height          =   285
         Left            =   1200
         TabIndex        =   20
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtDeviateSimple 
         Height          =   285
         Left            =   1200
         TabIndex        =   17
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblCount 
         Caption         =   "Count:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblStringLengthSimple 
         Caption         =   "String Length:"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblDeviateSimple 
         Caption         =   "Deviate by:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   1095
      End
   End
   Begin VB.Frame fraLineBenchmark 
      Caption         =   "Line Benchmark"
      Height          =   1335
      Left            =   2400
      TabIndex        =   2
      Top             =   5880
      Width           =   4215
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1060
         Left            =   120
         ScaleHeight     =   1065
         ScaleWidth      =   3975
         TabIndex        =   3
         Top             =   240
         Width           =   3975
         Begin VB.TextBox txtDeviateLine 
            Height          =   285
            Left            =   3120
            TabIndex        =   19
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox txtIncrements 
            Height          =   285
            Left            =   1080
            TabIndex        =   11
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox txtFinalCount 
            Height          =   285
            Left            =   3120
            TabIndex        =   8
            Top             =   0
            Width           =   855
         End
         Begin VB.TextBox txtStringLengthLine 
            Height          =   285
            Left            =   1080
            TabIndex        =   5
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox txtInitialCount 
            Height          =   285
            Left            =   1080
            TabIndex        =   4
            Top             =   0
            Width           =   855
         End
         Begin VB.Label lblDeviateLine 
            Caption         =   "Deviate by:"
            Height          =   255
            Left            =   2040
            TabIndex        =   18
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label lblIncrements 
            Caption         =   "Increments:"
            Height          =   255
            Left            =   0
            TabIndex        =   10
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label lblFinalCount 
            Caption         =   "Final Count:"
            Height          =   255
            Left            =   2040
            TabIndex        =   9
            Top             =   0
            Width           =   975
         End
         Begin VB.Label lblStringLength 
            Caption         =   "String Length:"
            Height          =   255
            Left            =   0
            TabIndex        =   7
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label lblInitialCount 
            Caption         =   "Initial Count:"
            Height          =   255
            Left            =   0
            TabIndex        =   6
            Top             =   0
            Width           =   975
         End
      End
   End
   Begin VB.CommandButton cmdBenchmark 
      Caption         =   "Benchmark"
      Height          =   375
      Left            =   6360
      TabIndex        =   1
      Top             =   8280
      Width           =   1215
   End
   Begin MSChart20Lib.MSChart chtBenchmark 
      Height          =   5775
      Left            =   0
      OleObjectBlob   =   "FMain.frx":0000
      TabIndex        =   0
      Top             =   0
      Width           =   7695
   End
   Begin VB.Label lblTrials 
      Caption         =   "Trials:"
      Height          =   255
      Left            =   6720
      TabIndex        =   25
      Top             =   5880
      Width           =   495
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************************************
'
'    Module:            FMain [FMain.frm]
'    Author:            Ralph Eastwood
'    Email:             tcmreastwood@ntlworld.com
'    Creation:          Sunday 24, December 2006, 7:17:25 AM
'    Last Modification: Thursday 4, January 2007, 6:19:43 PM
'    Purpose:           Benchmark String Sorting Routines
'    Notes:
'    Update History:    - Created (24/12/2006)
'
'***********************************************************************************************

Option Explicit

Private m_cTimer As New CTiming
Private m_cMSDRadix As New CMSDRadixSort
Private m_cTriQuickSort As New CTriQuickSort
Private m_cQuickie As New clsQuickie
Private WithEvents m_cUlliRadixSort As cSort
Attribute m_cUlliRadixSort.VB_VarHelpID = -1

Private m_lUlliRadixOutputCounter As Long
Private m_asUlliRadixOutput() As String
Private m_asString() As String
Private m_asStringHold() As String
Private m_arBenchmarkData() As Double
Private m_lInitialCount As Long
Private m_lFinalCount As Long
Private m_lStringLength As Long
Private m_lStringDeviateLength As Long
Private m_lIncrements As Long
        

Private Sub chkAmericanFlag_Click()
    If GetBenchmarkCount > 1 Then
        ReDim arArray(GetBenchmarkCount - 1)
        chtBenchmark.ChartData = arArray
    End If
    chtBenchmark.RowCount = 1
    chtBenchmark.Row = 1
    chtBenchmark.RowLabel = vbNullString
    LabelColumns
End Sub

Private Sub chkStrSwapStable_Click()
    If GetBenchmarkCount > 1 Then
        ReDim arArray(GetBenchmarkCount - 1)
        chtBenchmark.ChartData = arArray
    End If
    chtBenchmark.RowCount = 1
    chtBenchmark.Row = 1
    chtBenchmark.RowLabel = vbNullString
    LabelColumns
End Sub

Private Sub chkTriQuicksort_Click()
    If GetBenchmarkCount > 1 Then
        ReDim arArray(GetBenchmarkCount - 1)
        chtBenchmark.ChartData = arArray
    End If
    chtBenchmark.RowCount = 1
    chtBenchmark.Row = 1
    chtBenchmark.RowLabel = vbNullString
    LabelColumns
End Sub

Private Sub chkTriQuickSortNR_Click()
    If GetBenchmarkCount > 1 Then
        ReDim arArray(GetBenchmarkCount - 1)
        chtBenchmark.ChartData = arArray
    End If
    chtBenchmark.RowCount = 1
    chtBenchmark.Row = 1
    chtBenchmark.RowLabel = vbNullString
    LabelColumns
End Sub

Private Sub chkTwoArray_Click()
    If GetBenchmarkCount > 1 Then
        ReDim arArray(GetBenchmarkCount - 1)
        chtBenchmark.ChartData = arArray
    End If
    chtBenchmark.RowCount = 1
    chtBenchmark.Row = 1
    chtBenchmark.RowLabel = vbNullString
    LabelColumns
End Sub

Private Sub chkUlliQuickie_Click()
    If GetBenchmarkCount > 1 Then
        ReDim arArray(GetBenchmarkCount - 1)
        chtBenchmark.ChartData = arArray
    End If
    chtBenchmark.RowCount = 1
    chtBenchmark.Row = 1
    chtBenchmark.RowLabel = vbNullString
    LabelColumns
End Sub

Private Sub chkUlliRadix_Click()
    If GetBenchmarkCount > 1 Then
        ReDim arArray(GetBenchmarkCount - 1)
        chtBenchmark.ChartData = arArray
    End If
    chtBenchmark.RowCount = 1
    chtBenchmark.Row = 1
    chtBenchmark.RowLabel = vbNullString
    LabelColumns
End Sub

Private Sub cmdBenchmark_Click()
  
    Dim lStringCount As Long
    Dim lStringLength As Long
    Dim lStringDeviateLength As Long
    Dim arBenchmarkData() As Double
    Dim rHigh As Double
    Dim rLow As Double
    Dim rTime As Double
    
    Dim lInitialCount As Long
    Dim lFinalCount As Long
    Dim lIncrements As Long
    Dim i As Long, j As Long, k As Long, l As Long
    Dim lTrials As Long
    
    Dim strLine As String
    Dim lFileNo As Long
    
    lTrials = Val(txtTrials.Text)
    
    If optSimpleBenchmark.Value Then
    
        lStringCount = Val(txtCount.Text)
        lStringLength = Val(txtStringLengthSimple.Text)
        lStringDeviateLength = Val(txtDeviateSimple.Text)
        
        ReDim arBenchmarkData(GetBenchmarkCount - 1)
        
        ' Generate the random strings
        GenerateRandomStrings lStringLength, lStringCount, m_asStringHold, lStringDeviateLength
        ReDim m_asUlliRadixOutput(UBound(m_asStringHold))
        
        ' Warm Cache
        m_asString = m_asStringHold
        m_cMSDRadix.MSD_TA_RadixSort m_asString, lStringCount
        
        ' Begin Benchmark
        For i = 0 To UBound(arBenchmarkData)
            arBenchmarkData(i) = 999999
        Next i
        k = -1
        
        If chkTriQuicksort.Value Then
            FMain.Caption = "String Sort Benchmark - TriQuickSortString"
            DoEvents
            k = k + 1
            For i = 0 To lTrials - 1
                m_asString = m_asStringHold
                m_cTimer.Reset
                m_cTriQuickSort.TriQuickSortString m_asString, 0, lStringCount - 1
                rTime = m_cTimer.Elapsed
                If arBenchmarkData(k) > rTime Then arBenchmarkData(k) = rTime
            Next i
            If CheckSort(m_asString, lStringCount, False) >= 0 Then MsgBox "TriQuickSort Failed!"
        End If
        
        If chkTriQuickSortNR.Value Then
            FMain.Caption = "String Sort Benchmark - TriQSortString"
            DoEvents
            k = k + 1
            For i = 0 To lTrials - 1
                m_asString = m_asStringHold
                m_cTimer.Reset
                m_cTriQuickSort.TriQSortString m_asString, 0, lStringCount - 1
                rTime = m_cTimer.Elapsed
                If arBenchmarkData(k) > rTime Then arBenchmarkData(k) = rTime
            Next i
            If CheckSort(m_asString, lStringCount, False) >= 0 Then MsgBox "TriQSort NR Failed!"
        End If
        
        If chkAmericanFlag.Value Then
            FMain.Caption = "String Sort Benchmark - MSD_AF_RadixSort"
            DoEvents
            k = k + 1
            For i = 0 To lTrials - 1
                m_asString = m_asStringHold
                m_cTimer.Reset
                m_cMSDRadix.MSD_AF_RadixSort m_asString, lStringCount
                rTime = m_cTimer.Elapsed
                If arBenchmarkData(k) > rTime Then arBenchmarkData(k) = rTime
            Next i
            If CheckSort(m_asString, lStringCount, False) >= 0 Then MsgBox "AF Radix Failed!"
        End If
        
        If chkTwoArray.Value Then
            FMain.Caption = "String Sort Benchmark - MSD_TA_RadixSort"
            DoEvents
            k = k + 1
            For i = 0 To lTrials - 1
                m_asString = m_asStringHold
                m_cTimer.Reset
                m_cMSDRadix.MSD_TA_RadixSort m_asString, lStringCount
                rTime = m_cTimer.Elapsed
                If arBenchmarkData(k) > rTime Then arBenchmarkData(k) = rTime
            Next i
            If CheckSort(m_asString, lStringCount, False) >= 0 Then MsgBox "TA Radix Failed!"
        End If
        
        If chkStrSwapStable.Value Then
            FMain.Caption = "String Sort Benchmark - strSwapStable"
            DoEvents
            k = k + 1
            For i = 0 To lTrials - 1
                m_asString = m_asStringHold
                m_cTimer.Reset
                mStableQuick.strSwapStabletemp m_asString, 0, lStringCount, vbBinaryCompare, dAscending
                rTime = m_cTimer.Elapsed
                If arBenchmarkData(k) > rTime Then arBenchmarkData(k) = rTime
            Next i
            If CheckSort(m_asString, lStringCount, False) >= 0 Then MsgBox "strSwapStable Failed!"
        End If
        
        If chkUlliQuickie.Value Then
            FMain.Caption = "String Sort Benchmark - Ulli's Quickie"
            DoEvents
            k = k + 1
            For i = 0 To lTrials - 1
                m_asString = m_asStringHold
                m_cTimer.Reset
                m_cQuickie.SuperQuickie m_asString, 0, lStringCount - 1
                rTime = m_cTimer.Elapsed
                If arBenchmarkData(k) > rTime Then arBenchmarkData(k) = rTime
            Next i
            If CheckSort(m_asString, lStringCount, False) >= 0 Then MsgBox "Ulli's Quickie Failed!"
        End If
        
        If chkUlliRadix.Value Then
            FMain.Caption = "String Sort Benchmark - Ulli's Radix Sort"
            DoEvents
            k = k + 1
            With m_cUlliRadixSort
                .LowBound = 1
                .HighBound = lStringCount
                .KeyPosition = 1
                .KeySize = Val(txtStringLengthSimple.Text) '256 'We have to tell Ulli's implementation what is the max length of the sort strings
                .PartialKeys = LessFullKeys
                .SortDirection = Ascending
                .RightToLeft = False
                For i = 0 To lTrials - 1
                    m_lUlliRadixOutputCounter = 0
                    'm_asString = m_asStringHold
                    'Ulli's sort doesn't support 0-based arrays :S
                    ReDim m_asString(1 To UBound(m_asStringHold) + 1)
                    For j = 0 To UBound(m_asStringHold)
                        m_asString(j + 1) = m_asStringHold(j)
                    Next j
                    m_cTimer.Reset
                    m_cUlliRadixSort.SortTable m_asString
                    rTime = m_cTimer.Elapsed
                    If arBenchmarkData(k) > rTime Then arBenchmarkData(k) = rTime
                Next i
            
            End With
            If CheckSort(m_asUlliRadixOutput, lStringCount, False) >= 0 Then MsgBox "Ulli's Radix Failed!"
        End If
        
        ' Get range of data
        rLow = 9999999
        For i = 0 To GetBenchmarkCount - 1
            If arBenchmarkData(i) > rHigh Then rHigh = arBenchmarkData(i)
            If arBenchmarkData(i) < rLow Then rLow = arBenchmarkData(i)
        Next i
        
        ' Display on chart
        With chtBenchmark
            .ChartData = arBenchmarkData
            .chartType = VtChChartType2dBar
            .RowCount = 1
            .RowLabel = vbNullString
            LabelColumns
            .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = rHigh
            .Plot.Axis(VtChAxisIdX).CategoryScale.Auto = True
            .Plot.Axis(VtChAxisIdX).AxisTitle = "Sorting Algorithm"
        End With
    
    ElseIf optLineBenchmark.Value Then
        
        lInitialCount = Val(txtInitialCount.Text)
        lFinalCount = Val(txtFinalCount.Text)
        lStringLength = Val(txtStringLengthSimple.Text)
        lStringDeviateLength = Val(txtDeviateSimple.Text)
        lIncrements = Val(txtIncrements.Text)
        m_lInitialCount = lInitialCount
        m_lFinalCount = lFinalCount
        m_lStringLength = lStringLength
        m_lStringDeviateLength = lStringDeviateLength
        m_lIncrements = lIncrements
        
        rLow = 9999999
        
        ' Generate the random strings

        ReDim arBenchmarkData((lFinalCount - lInitialCount) \ lIncrements, GetBenchmarkCount - 1)
                
        ' Generate random strings
        GenerateRandomStrings lStringLength, lFinalCount, m_asStringHold, lStringDeviateLength
            
        ' Do the benchmark
        For i = 0 To lTrials - 1
            j = 0
            
            For lStringCount = lInitialCount To lFinalCount Step lIncrements
                
                If i = 0 Then
                    For k = 0 To UBound(arBenchmarkData, 2)
                        arBenchmarkData(j, k) = 999999
                    Next k
                End If
            
                l = -1
            
                FMain.Caption = "String Sort Benchmark - Trial (" & i & "/" & lTrials - 1 & ") Count = (" & lStringCount & "/" & lFinalCount & ")"
                
                If chkTriQuicksort.Value Then
                    l = l + 1
                    m_asString = m_asStringHold
                    m_cTimer.Reset
                    m_cTriQuickSort.TriQuickSortString m_asString, 0, lStringCount - 1
                    rTime = m_cTimer.Elapsed
                    If arBenchmarkData(j, l) > rTime Then arBenchmarkData(j, l) = rTime
                End If
                
                If chkTriQuickSortNR.Value Then
                    l = l + 1
                    m_asString = m_asStringHold
                    m_cTimer.Reset
                    m_cTriQuickSort.TriQSortString m_asString, 0, lStringCount - 1
                    rTime = m_cTimer.Elapsed
                    If arBenchmarkData(j, l) > rTime Then arBenchmarkData(j, l) = rTime
                End If
                
                If chkAmericanFlag.Value Then
                    l = l + 1
                    m_asString = m_asStringHold
                    m_cTimer.Reset
                    m_cMSDRadix.MSD_AF_RadixSort m_asString, lStringCount
                    rTime = m_cTimer.Elapsed
                    If arBenchmarkData(j, l) > rTime Then arBenchmarkData(j, l) = rTime
                End If
                
                If chkTwoArray.Value Then
                    l = l + 1
                    m_asString = m_asStringHold
                    m_cTimer.Reset
                    m_cMSDRadix.MSD_TA_RadixSort m_asString, lStringCount
                    rTime = m_cTimer.Elapsed
                    If arBenchmarkData(j, l) > rTime Then arBenchmarkData(j, l) = rTime
                End If
                
                If chkStrSwapStable.Value Then
                    l = l + 1
                    m_asString = m_asStringHold
                    m_cTimer.Reset
                    mStableQuick.strSwapStabletemp m_asString, 0, lStringCount, vbBinaryCompare, dAscending
                    rTime = m_cTimer.Elapsed
                    If arBenchmarkData(j, l) > rTime Then arBenchmarkData(j, l) = rTime
                End If
                
                If chkUlliQuickie.Value Then
                    l = l + 1
                    m_asString = m_asStringHold
                    m_cTimer.Reset
                    m_cQuickie.SuperQuickie m_asString, 0, lStringCount - 1
                    rTime = m_cTimer.Elapsed
                    If arBenchmarkData(j, l) > rTime Then arBenchmarkData(j, l) = rTime
                End If
                
                If chkUlliRadix.Value Then
                    l = l + 1
                    m_lUlliRadixOutputCounter = 0

                    With m_cUlliRadixSort
                        .LowBound = 1
                        .HighBound = lStringCount
                        .KeyPosition = 1
                        .KeySize = Val(txtStringLengthLine.Text) '256
                        .PartialKeys = LessFullKeys
                        .SortDirection = Ascending
                        
                        'm_asString = m_asStringHold
                        'Ulli's sort doesn't support 0-based arrays :S
                        ReDim m_asUlliRadixOutput(UBound(m_asStringHold))
                        ReDim m_asString(1 To UBound(m_asStringHold) + 1)
                        For k = 0 To UBound(m_asStringHold)
                            m_asString(k + 1) = m_asStringHold(k)
                        Next k
                        m_cTimer.Reset
                        .SortTable m_asString
                        rTime = m_cTimer.Elapsed
                        If arBenchmarkData(j, l) > rTime Then arBenchmarkData(j, l) = rTime
                        
                    End With

                End If
                
                j = j + 1
                
            Next lStringCount
        Next i
                      
        ' Get Data range
        For i = 0 To GetBenchmarkCount - 1
            j = 0
            For lStringCount = lInitialCount To lFinalCount Step lIncrements
                If arBenchmarkData(j, i) > rHigh Then rHigh = arBenchmarkData(j, i)
                If arBenchmarkData(j, i) < rLow Then rLow = arBenchmarkData(j, i)
                j = j + 1
            Next lStringCount
        Next i
                    
        FMain.Caption = "String Sort Benchmark - Count = (" & lFinalCount & "/" & lFinalCount & ")"
                        
                            
        ' Display on chart
        With chtBenchmark
            .ChartData = arBenchmarkData
            .chartType = VtChChartType2dLine
            LabelColumns
            .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = rHigh
            .Plot.Axis(VtChAxisIdX).CategoryScale.Auto = False
            .Plot.Axis(VtChAxisIdX).CategoryScale.DivisionsPerLabel = ((lFinalCount - lInitialCount) \ lIncrements) \ 10
            .Plot.Axis(VtChAxisIdX).CategoryScale.DivisionsPerTick = ((lFinalCount - lInitialCount) \ lIncrements) \ 10
            .Plot.Axis(VtChAxisIdX).AxisTitle = "Number of Elements"

            lStringCount = lInitialCount
            For i = 0 To .RowCount - 1
                .Row = i + 1
                .RowLabel = Str$(lStringCount)
                lStringCount = lStringCount + lIncrements
            Next i
        End With
    ElseIf optFileBenchmark.Value Then
    
        If LenB(Dir$(App.Path & "\Dictionaries", vbDirectory)) = 0 Then MkDir App.Path & "\Dictionaries"
        If LenB(Dir$(App.Path & "\Output", vbDirectory)) = 0 Then MkDir App.Path & "\Output"
    
        ' Load a dictionary file
        With cmndlg
            .DefaultExt = "dic"
            .DialogTitle = "Load a Dictionary File"
            .Filter = "Dictionary Files (*.dic)|*.dic|Text Files (*.txt)|*.txt"
            .InitDir = App.Path & "\Dictionaries"
            .ShowOpen
        End With

        ' If the file exists
        If (LenB(cmndlg.FileName) > 0) And (LenB(Dir$(cmndlg.FileName)) > 0) Then
            
            i = 0
            lFileNo = FreeFile
            ' Load the file data
            Open cmndlg.FileName For Input As #lFileNo
            Do Until EOF(lFileNo)
                Line Input #lFileNo, strLine
                ReDim Preserve m_asStringHold(i)
                m_asStringHold(i) = strLine
                If (i Mod 1000) = 0 Then FMain.Caption = "String Sort Benchmark - Loading " & cmndlg.FileTitle & "(" & i & ")"
                i = i + 1
            Loop
            FMain.Caption = "String Sort Benchmark - Loading " & cmndlg.FileTitle & "(" & (i - 1) & ")"
            Close #lFileNo
            
            lStringCount = i
            
            ReDim m_asUlliRadixOutput(UBound(m_asStringHold))
            
            ' Randomise the dictionary list
            FMain.Caption = "String Sort Benchmark - Randomising..."
            'RandomiseList m_asString, m_asStringHold
            Dim lLBound As Long, lUBound As Long
            Dim r As Long, sTmp As String
            
            lUBound = UBound(m_asStringHold)
            
            For i = 0 To lUBound
                r = Int((lUBound + 1) * Rnd)
                sTmp = m_asStringHold(i)
                m_asStringHold(i) = m_asStringHold(r)
                m_asStringHold(r) = sTmp
                If (i Mod 1000) = 0 Then FMain.Caption = "String Sort Benchark - Randomising..." & "(" & i & "/" & (lUBound) & ")"
            Next i
            FMain.Caption = "String Sort Benchmark - Randomising..." & "(" & (lUBound) & "/" & (lUBound) & ")"
                
            ' Start benchmark sorting the dictionary
            
            ' Begin Benchmark
            ReDim arBenchmarkData(GetBenchmarkCount - 1)
            For i = 0 To UBound(arBenchmarkData)
                arBenchmarkData(i) = 999999
            Next i
            
            l = -1
            
            ' Benchmark triqsort
            If chkTriQuicksort.Value Then
                l = l + 1
                For i = 0 To lTrials - 1
                    FMain.Caption = "String Sort Benchmark - TriQuickSort (" & i & "/" & lTrials - 1 & ")"
                    DoEvents
                    m_asString = m_asStringHold
                    m_cTimer.Reset
                    m_cTriQuickSort.TriQuickSortString m_asString, 0, lStringCount - 1
                    rTime = m_cTimer.Elapsed
                    If arBenchmarkData(l) > rTime Then arBenchmarkData(l) = rTime
                Next i
                If CheckSort(m_asString, lStringCount, False) >= 0 Then MsgBox "TriQuickSort Failed!"
            
                ' Output the triqsorted dictionary
                lFileNo = FreeFile
                Open App.Path & "\Output\" & "QS_" & cmndlg.FileTitle For Output As #lFileNo
                If Not (Not m_asString) Then
                    For i = 0 To UBound(m_asString)
                        Print #lFileNo, m_asString(i)
                        If (i Mod 1000) = 0 Then FMain.Caption = "String Sort Benchmark - " & "\Output\" & "QS_" & cmndlg.FileTitle & "(" & i & "/" & UBound(m_asString) & ")"
                    Next i
                    FMain.Caption = "String Sort Benchmark - " & "\Output\" & "QS_" & cmndlg.FileTitle & "(" & i & "/" & UBound(m_asString) & ")"
                End If
                Close #lFileNo
            End If
            
            ' Benchmark triqsort
            If chkTriQuickSortNR.Value Then
                l = l + 1
                For i = 0 To lTrials - 1
                    FMain.Caption = "String Sort Benchmark - TriQSort NR (" & i & "/" & lTrials - 1 & ")"
                    DoEvents
                    m_asString = m_asStringHold
                    m_cTimer.Reset
                    m_cTriQuickSort.TriQSortString m_asString, 0, lStringCount - 1
                    rTime = m_cTimer.Elapsed
                    If arBenchmarkData(l) > rTime Then arBenchmarkData(l) = rTime
                Next i
                If CheckSort(m_asString, lStringCount, False) >= 0 Then MsgBox "TriQSort NR Failed!"
                
                ' Output the triqsorted dictionary
                lFileNo = FreeFile
                Open App.Path & "\Output\" & "QN_" & cmndlg.FileTitle For Output As #lFileNo
                If Not (Not m_asString) Then
                    For i = 0 To UBound(m_asString)
                        Print #lFileNo, m_asString(i)
                        If (i Mod 1000) = 0 Then FMain.Caption = "String Sort Benchmark - " & "\Output\" & "QN_" & cmndlg.FileTitle & "(" & i & "/" & UBound(m_asString) & ")"
                    Next i
                    FMain.Caption = "String Sort Benchmark - " & "\Output\" & "QS_" & cmndlg.FileTitle & "(" & i & "/" & UBound(m_asString) & ")"
                End If
                Close #lFileNo
            End If
            
            ' Benchmark MSD_AF_RadixSort
            If chkAmericanFlag.Value Then
                l = l + 1
                For i = 0 To lTrials - 1
                    FMain.Caption = "String Sort Benchmark - MSD AF Radix (" & i & "/" & lTrials - 1 & ")"
                    DoEvents
                    m_asString = m_asStringHold
                    m_cTimer.Reset
                    m_cMSDRadix.MSD_AF_RadixSort m_asString, lStringCount
                    rTime = m_cTimer.Elapsed
                    If arBenchmarkData(l) > rTime Then arBenchmarkData(l) = rTime
                Next i
                If CheckSort(m_asString, lStringCount, False) >= 0 Then MsgBox "AF Radix Failed!"
                
                ' Output the af sorted dictionary
                lFileNo = FreeFile
                Open App.Path & "\Output\" & "AF_" & cmndlg.FileTitle For Output As #lFileNo
                If Not (Not m_asString) Then
                    For i = 0 To UBound(m_asString)
                        Print #lFileNo, m_asString(i)
                        If (i Mod 1000) = 0 Then FMain.Caption = "String Sort Benchmark - " & "\Output\" & "AF_" & cmndlg.FileTitle & "(" & i & "/" & UBound(m_asString) & ")"
                    Next i
                    FMain.Caption = "String Sort Benchmark - " & "\Output\" & "AF_" & cmndlg.FileTitle & "(" & i & "/" & UBound(m_asString) & ")"
                End If
                Close #lFileNo
            End If
            
            ' Benchmark MSD_TA_RadixSort
            If chkTwoArray.Value Then
                l = l + 1
                For i = 0 To lTrials - 1
                    FMain.Caption = "String Sort Benchmark - MSD TA Radix (" & i & "/" & lTrials - 1 & ")"
                    DoEvents
                    m_asString = m_asStringHold
                    m_cTimer.Reset
                    m_cMSDRadix.MSD_TA_RadixSort m_asString, lStringCount
                    rTime = m_cTimer.Elapsed
                    If arBenchmarkData(l) > rTime Then arBenchmarkData(l) = rTime
                Next i
                If CheckSort(m_asString, lStringCount, False) >= 0 Then MsgBox "TA Radix Failed!"
                
                ' Output the ta sorted dictionary
                lFileNo = FreeFile
                Open App.Path & "\Output\" & "TA_" & cmndlg.FileTitle For Output As #lFileNo
                If Not (Not m_asString) Then
                    For i = 0 To UBound(m_asString)
                        Print #lFileNo, m_asString(i)
                        If (i Mod 1000) = 0 Then FMain.Caption = "String Sort Benchmark - " & "\Output\" & "TA_" & cmndlg.FileTitle & "(" & i & "/" & UBound(m_asString) & ")"
                    Next i
                    FMain.Caption = "String Sort Benchmark - " & "\Output\" & "TA_" & cmndlg.FileTitle & "(" & i & "/" & UBound(m_asString) & ")"
                End If
                Close #lFileNo
            End If
            
            ' Benchmark strSwapStabletemp
            If chkStrSwapStable.Value Then
                l = l + 1
                For i = 0 To lTrials - 1
                    FMain.Caption = "String Sort Benchmark - strSwapStable (" & i & "/" & lTrials - 1 & ")"
                    DoEvents
                    m_asString = m_asStringHold
                    m_cTimer.Reset
                    mStableQuick.strSwapStabletemp m_asString, 0, lStringCount, vbBinaryCompare, dAscending
                    rTime = m_cTimer.Elapsed
                    If arBenchmarkData(l) > rTime Then arBenchmarkData(l) = rTime
                Next i
                If CheckSort(m_asString, lStringCount, False) >= 0 Then MsgBox "strSwapStable Failed!"
                
                ' Output the ss sorted dictionary
                lFileNo = FreeFile
                Open App.Path & "\Output\" & "SS_" & cmndlg.FileTitle For Output As #lFileNo
                If Not (Not m_asString) Then
                    For i = 0 To UBound(m_asString)
                        Print #lFileNo, m_asString(i)
                        If (i Mod 1000) = 0 Then FMain.Caption = "String Sort Benchmark - " & "\Output\" & "SS_" & cmndlg.FileTitle & "(" & i & "/" & UBound(m_asString) & ")"
                    Next i
                    FMain.Caption = "String Sort Benchmark - " & "\Output\" & "SS_" & cmndlg.FileTitle & "(" & i & "/" & UBound(m_asString) & ")"
                End If
                Close #lFileNo
            End If
            
            ' Benchmark Ulli's Quickie
            If chkUlliQuickie.Value Then
                l = l + 1
                For i = 0 To lTrials - 1
                    FMain.Caption = "String Sort Benchmark - Ulli's Quickie (" & i & "/" & lTrials - 1 & ")"
                    DoEvents
                    m_asString = m_asStringHold
                    m_cTimer.Reset
                    m_cQuickie.SuperQuickie m_asString, 0, lStringCount - 1
                    rTime = m_cTimer.Elapsed
                    If arBenchmarkData(l) > rTime Then arBenchmarkData(l) = rTime
                Next i
                If CheckSort(m_asString, lStringCount, False) >= 0 Then MsgBox "Ulli's Quickie Failed!"
                
                ' Output the uq sorted dictionary
                lFileNo = FreeFile
                Open App.Path & "\Output\" & "UQ_" & cmndlg.FileTitle For Output As #lFileNo
                If Not (Not m_asString) Then
                    For i = 0 To UBound(m_asString)
                        Print #lFileNo, m_asString(i)
                        If (i Mod 1000) = 0 Then FMain.Caption = "String Sort Benchmark - " & "\Output\" & "UQ_" & cmndlg.FileTitle & "(" & i & "/" & UBound(m_asString) & ")"
                    Next i
                    FMain.Caption = "String Sort Benchmark - " & "\Output\" & "UQ_" & cmndlg.FileTitle & "(" & i & "/" & UBound(m_asString) & ")"
                End If
                Close #lFileNo
            End If
            
            ' Benchmark Ulli's MSD Radix
            If chkUlliRadix.Value Then
                l = l + 1
                For i = 0 To lTrials - 1
                    FMain.Caption = "String Sort Benchmark - Ulli's Radix Sort (" & i & "/" & lTrials - 1 & ")"
                    DoEvents
                    m_lUlliRadixOutputCounter = 0
                    
                    'm_asString = m_asStringHold
                    'Ulli's sort doesn't support 0-based arrays :S
                    ReDim m_asString(1 To UBound(m_asStringHold) + 1)
                    For j = 0 To UBound(m_asStringHold)
                        m_asString(j + 1) = m_asStringHold(j)
                    Next j
                    
                    With m_cUlliRadixSort
                        .LowBound = 1
                        .HighBound = lStringCount
                        .KeyPosition = 1
                        .KeySize = 256
                        .PartialKeys = LessFullKeys
                        .SortDirection = Ascending
                        
                        m_cTimer.Reset
                        .SortTable m_asString
                        rTime = m_cTimer.Elapsed
                        If arBenchmarkData(l) > rTime Then arBenchmarkData(l) = rTime
                        
                    End With
                Next i
                Dim lFail As Long
                lFail = CheckSort(m_asUlliRadixOutput, lStringCount, False)
                If lFail >= 0 Then MsgBox "Ulli's Radix Sort Failed at Index " & lFail
                
                ' Output the uq sorted dictionary
                lFileNo = FreeFile
                Open App.Path & "\Output\" & "UR_" & cmndlg.FileTitle For Output As #lFileNo
                If Not (Not m_asUlliRadixOutput) Then
                    For i = 0 To UBound(m_asUlliRadixOutput)
                        Print #lFileNo, m_asUlliRadixOutput(i)
                        If (i Mod 1000) = 0 Then FMain.Caption = "String Sort Benchmark - " & "\Output\" & "UR_" & cmndlg.FileTitle & "(" & i & "/" & UBound(m_asUlliRadixOutput) & ")"
                    Next i
                    FMain.Caption = "String Sort Benchmark - " & "\Output\" & "UR_" & cmndlg.FileTitle & "(" & i & "/" & UBound(m_asUlliRadixOutput) & ")"
                End If
                Close #lFileNo
            End If
            
            ' Get range of data
            rLow = 9999999
            For i = 0 To GetBenchmarkCount - 1
                If arBenchmarkData(i) > rHigh Then rHigh = arBenchmarkData(i)
                If arBenchmarkData(i) < rLow Then rLow = arBenchmarkData(i)
            Next i
                
            ' Display on chart
            With chtBenchmark
                .ChartData = arBenchmarkData
                .chartType = VtChChartType2dBar
                .RowCount = 1
                .RowLabel = vbNullString
                LabelColumns
                .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = rHigh
                .Plot.Axis(VtChAxisIdX).CategoryScale.Auto = True
                .Plot.Axis(VtChAxisIdX).AxisTitle = "Sorting Algorithm"
            End With
        
                
        End If

    End If

    m_arBenchmarkData = arBenchmarkData

    FMain.Caption = "String Sort Benchmark"
End Sub

Private Sub cmdExportResults_Click()

    Dim lf As Long
    Dim i As Long, j As Long, k As Long
    Dim strData As String
    
    ' Export the results to file
    With cmndlg
        .FileName = "Untitled.csv"
        .DefaultExt = "csv"
        .ShowSave
        
        lf = FreeFile
        Open .FileName For Output As #lf
        
        If optSimpleBenchmark.Value Or optFileBenchmark.Value Then
            strData = strData & """Elements""," & txtCount.Text & vbNewLine
            strData = strData & """Length"",""" & txtStringLengthSimple.Text & " ± " & txtDeviateSimple.Text & """" & vbNewLine
            strData = strData & """Benchmark"",""TimeTaken (ms)""" & vbNewLine
            For i = 1 To chtBenchmark.ColumnCount
                chtBenchmark.Column = i
                strData = strData & """" & chtBenchmark.ColumnLabel & """," & m_arBenchmarkData(i - 1) & vbNewLine
            Next i
            
            Print #lf, strData
                
        ElseIf optLineBenchmark.Value Then

            strData = """String Length: " & m_lStringLength & """,""Deviate Length: " & m_lStringDeviateLength & """" & vbNewLine & _
                      """Number Of Strings"","
                      
            For i = 1 To chtBenchmark.ColumnCount - 1
                chtBenchmark.Column = i
                strData = strData & """" & chtBenchmark.ColumnLabel & ""","
            Next i
            chtBenchmark.Column = i
            strData = strData & """" & chtBenchmark.ColumnLabel & """"
            Print #lf, strData
            
            k = m_lInitialCount
            For i = 0 To UBound(m_arBenchmarkData, 1)
                strData = k & ","
                For j = 0 To UBound(m_arBenchmarkData, 2) - 1
                    strData = strData & m_arBenchmarkData(i, j) & ","
                Next j
                k = k + m_lIncrements
                Print #lf, strData & m_arBenchmarkData(i, j)
            Next i

        End If
        
        Close #lf
        
    End With

    

End Sub

Private Sub Form_Load()
    
    Set m_cUlliRadixSort = New cSort
    
    ' Disable the Line Frame
    optSimpleBenchmark_Click
   
    ' Initialise the stack
    m_cTriQuickSort.InitTriQSortStack
   
    ' Initial Values
    
    ' Simple
    txtCount.Text = "10000"
    txtDeviateSimple.Text = "0"
    txtStringLengthSimple = "30"
    
    ' Line
    txtInitialCount.Text = "10"
    txtFinalCount.Text = "1000"
    txtIncrements.Text = "10"
    txtDeviateLine.Text = "0"
    txtStringLengthLine.Text = "30"
    
    txtTrials.Text = "5"
End Sub

Private Sub EnableLineFrame(ByVal fEnable As Boolean)
    fraLineBenchmark.Enabled = fEnable
    lblInitialCount.Enabled = fEnable
    lblFinalCount.Enabled = fEnable
    lblIncrements.Enabled = fEnable
    lblStringLength.Enabled = fEnable
    txtInitialCount.Enabled = fEnable
    txtIncrements.Enabled = fEnable
    txtFinalCount.Enabled = fEnable
    lblDeviateLine.Enabled = fEnable
    txtDeviateLine.Enabled = fEnable
    txtStringLengthLine.Enabled = fEnable
End Sub
Private Sub EnableSimpleFrame(ByVal fEnable As Boolean)
    fraSimpleBenchmark.Enabled = fEnable
    lblCount.Enabled = fEnable
    lblStringLengthSimple.Enabled = fEnable
    txtCount.Enabled = fEnable
    lblStringLengthSimple.Enabled = fEnable
    lblDeviateSimple.Enabled = fEnable
    txtDeviateSimple.Enabled = fEnable
    txtStringLengthSimple.Enabled = fEnable
End Sub

Private Sub m_cUlliRadixSort_NextPointer(ByVal SortId As Long, ByVal Pointer As Long, Cancel As Boolean)
    m_asUlliRadixOutput(m_lUlliRadixOutputCounter) = m_asString(Pointer)
    m_lUlliRadixOutputCounter = m_lUlliRadixOutputCounter + 1
End Sub
Private Sub optFileBenchmark_Click()
    EnableSimpleFrame optSimpleBenchmark.Value
    EnableLineFrame optLineBenchmark.Value
    
    ' Setup the Graph
    Dim arArrays(6) As Single
    With chtBenchmark
        .SeriesType = VtChSeriesType2dBar
        .ChartData = arArrays
        .RowCount = 1
        .Row = 1
        .RowLabel = vbNullString
        LabelColumns
        .Legend.Location.LocationType = VtChLocationTypeBottom
        .Plot.Axis(VtChAxisIdY).AxisTitle.Text = "Time taken (ms)"
        .AllowSelections = False
    End With
End Sub

Private Sub optLineBenchmark_Click()
    EnableSimpleFrame optSimpleBenchmark.Value
    EnableLineFrame optLineBenchmark.Value
    
    ' Setup the Graph
    Dim arArrays(6) As Single
    With chtBenchmark
        .SeriesType = VtChSeriesType2dLine
        .ChartData = arArrays
        .RowCount = 1
        .Row = 1
        .RowLabel = vbNullString
        LabelColumns
        .Legend.Location.LocationType = VtChLocationTypeBottom
        .Plot.Axis(VtChAxisIdY).AxisTitle.Text = "Time taken (ms)"
        .AllowSelections = False
    End With
End Sub

Private Sub optSimpleBenchmark_Click()
    EnableSimpleFrame optSimpleBenchmark.Value
    EnableLineFrame optLineBenchmark.Value
        
    ' Setup the Graph
    Dim arArrays(6) As Single
    With chtBenchmark
        .SeriesType = VtChSeriesType2dBar
        .ChartData = arArrays
        .RowCount = 1
        .Row = 1
        .RowLabel = vbNullString
        LabelColumns
        .Legend.Location.LocationType = VtChLocationTypeBottom
        .Plot.Axis(VtChAxisIdY).AxisTitle.Text = "Time taken (ms)"
        .AllowSelections = False
    End With
End Sub


Private Sub txtDeviateLine_LostFocus()
    If Val(txtDeviateLine.Text) > Val(txtStringLengthLine.Text) Then
        txtDeviateLine.Text = txtStringLengthLine.Text
    End If
End Sub

Private Sub txtDeviateSimple_LostFocus()
    If Val(txtDeviateSimple.Text) > Val(txtStringLengthSimple.Text) Then
        txtDeviateSimple.Text = txtStringLengthSimple.Text
    End If
End Sub

Private Sub LabelColumns()
    Dim i As Long
    With chtBenchmark
        .ColumnCount = GetBenchmarkCount
        If chkTriQuicksort.Value Then
            i = i + 1
            .Column = i
            .ColumnLabel = "Tri-Quick Sort"
        End If
        If chkTriQuickSortNR.Value Then
            i = i + 1
            .Column = i
            .ColumnLabel = "Tri-Quick Sort NR"
        End If
        If chkAmericanFlag.Value Then
            i = i + 1
            .Column = i
            .ColumnLabel = "American Flag"
        End If
        If chkTwoArray.Value Then
            i = i + 1
            .Column = i
            .ColumnLabel = "Two Array"
        End If
        If chkStrSwapStable.Value Then
            i = i + 1
            .Column = i
            .ColumnLabel = "strSwapStable"
        End If
        If chkUlliQuickie.Value Then
            i = i + 1
            .Column = i
            .ColumnLabel = "Ulli's Quickie"
        End If
        If chkUlliRadix.Value Then
            i = i + 1
            .Column = i
            .ColumnLabel = "Ulli's Radix Sort"
        End If
    End With
End Sub

Private Function GetBenchmarkCount() As Long
    GetBenchmarkCount = chkTriQuicksort.Value + chkTriQuickSortNR.Value + chkAmericanFlag.Value + chkTwoArray.Value + chkStrSwapStable.Value + chkUlliQuickie.Value + chkUlliRadix.Value
End Function
