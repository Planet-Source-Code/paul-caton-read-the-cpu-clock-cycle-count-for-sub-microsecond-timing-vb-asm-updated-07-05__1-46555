VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmMain 
   Caption         =   "CPU Clock Cycle Counter"
   ClientHeight    =   2895
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9090
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   2895
   ScaleWidth      =   9090
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lv 
      Height          =   1920
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   3387
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Test"
         Object.Width           =   5318
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Cycles"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Time (uS)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Prev Cycles"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Prev Time (uS)"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "&Test"
      Height          =   390
      Left            =   4635
      TabIndex        =   1
      Top             =   2190
      Width           =   1440
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
' Simple demonstration of the cCpuClk class being used to read the CPU clock
' cycle count for high resolution timing.
'
' Pentium class processors include a 64 bit register that increments from
' power-on at the CPU clock frequency. With a 2GHz processor, you have in
' effect a 2GHz clock. The cCpuClk class allows the user to retrieve the 64 bit
' CPU clock cycle count into a passed currency parameter. The class can be used
' as a basis for sub-microsecond benchmarking and delay timing. Note that with
' the extreme resolution provided, multitasking and the state of the cpu caches
' will show in the results. Thanks to David Fritts and Robert Rayment for the vtable trick.
'
' Paul Caton
' Paul_Caton@hotmail.com
'
Option Explicit

'QueryPerformance API declares, used to calculate the cpu frequency
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

'For XP manifests
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long

Private Sub Form_Initialize()
  Call InitCommonControls
End Sub

Private Sub Form_Load()
  With lv.ListItems
    Call .Add(, , "QueryPerformance freq/resolution")
    Call .Add(, , "Cpu clock frequency/resolution")
    Call .Add(, , "CpuClk 1st call")
    Call .Add(, , "CpuClk 2nd call")
    Call .Add(, , "QueryPerformanceCounter 1st call")
    Call .Add(, , "QueryPerformanceCounter 2nd call")
    Call .Add(, , "For 1000 Longs multiplied")
    Call .Add(, , "For 1000 Integers multiplied")
  End With
  Call Show
  Call DoTest
End Sub

Private Sub Form_Resize()
  Const STD_SPC As Long = 120
  
  On Error Resume Next
    Call lv.Move(STD_SPC, STD_SPC, ScaleWidth - STD_SPC - STD_SPC, ScaleHeight - STD_SPC - cmdTest.Height - STD_SPC - STD_SPC)
    Call cmdTest.Move(ScaleWidth - STD_SPC - cmdTest.Width, lv.Top + lv.Height + STD_SPC)
  On Error GoTo 0
End Sub

Private Sub cmdTest_Click()
  DoTest
End Sub

Private Sub DoTest()
  Const IMULT   As Integer = 7                              'Integer multiplier
  Const LMULT   As Long = 7                                 'Long multiplier
  Const CUR2INT As Long = 10000                             'Factor to convert Currency representation to int
  Const USEC    As Long = 1000000                           'Microsecond divisor
  Dim Clk       As cCpuClk                                  'Cpu clock class
  Dim li        As Long                                     'Long for loop test
  Dim lj        As Long                                     'Long for loop test
  Dim ii        As Integer                                  'Integer for loop test
  Dim ij        As Integer                                  'Integer for loop test
  Dim c1        As Currency                                 'Cpu clock cycle count 1
  Dim c2        As Currency                                 'Cpu clock cycle count 2
  Dim cCPU      As Currency                                 'Cpu frequency for time calculations
  Dim cCycles   As Currency                                 'QueryPerformance frequency
  Dim cOver1    As Currency                                 'CpuClk call overhead, 1st call
  Dim cOver2    As Currency                                 'CpuClk call overhead, 2nd call
  Dim cQpc1     As Currency                                 'QueryPerformanceCounter overhead, 1st call
  Dim cQpc2     As Currency                                 'QueryPerformanceCounter overhead, 2nd call
  Dim cLong     As Currency                                 'Long loop results
  Dim cInteger  As Currency                                 'Integer loop results
  Dim cNow      As Currency                                 'QueryPerformanceCounter time now
  Dim cStart    As Currency                                 'QueryPerformanceCounter start time
  Dim cStop     As Currency                                 'QueryPerformanceCounter stop time
  Dim Factor    As Double                                   'QueryPerformanceCounter overrun factor

'--Setup-----------------------------------------------------
  Set Clk = New cCpuClk                                     'Create the CpuClk instance
  Screen.MousePointer = vbHourglass
  cmdTest.Enabled = False
  DoEvents
  
'--Calculate call overhead-----------------------------------
  Call Clk.CpuClk(c1)
  Call Clk.CpuClk(c2)
  cOver1 = c2 - c1
  
  'Do it again to ensure caching isn't a factor
  Call Clk.CpuClk(c1)
  Call Clk.CpuClk(c2)
  cOver2 = c2 - c1
  
  'Measure QueryPerformanceCounter overhead. Just for information
  Call Clk.CpuClk(c1)
    Call QueryPerformanceCounter(cNow)
  Call Clk.CpuClk(c2)
  cQpc1 = c2 - c1 - cOver2
  
  'Do it again to show the differences related to caching
  Call Clk.CpuClk(c1)
    Call QueryPerformanceCounter(cNow)
  Call Clk.CpuClk(c2)
  cQpc2 = c2 - c1 - cOver2
  
'--Calculate cpu speed---------------------------------------
  Call QueryPerformanceFrequency(cCycles)                   'Get the QueryPerformance frequency
  
  Call Clk.CpuClk(c1)                                       'Get the start cpu cycle count
    Call QueryPerformanceCounter(cStart)                    'Get the start time
    cStop = cStart + cCycles                                'Calculate the stop time, start + freq = 1 second
  
    Do
      Call QueryPerformanceCounter(cNow)                    'Get the current time
    Loop Until (cNow >= cStop)                              'Loop until stop time
  Call Clk.CpuClk(c2)                                       'Get the stop cpu cycle count
  c2 = c2 - c1 - cOver2                                     'Calculate the actual number of cpu cycles
  Factor = CDbl(cNow - cStart) / CDbl(cCycles)              'Calculate a factor for the actual duration which may be slightly greater than 1 second (cNow > cStop)
  cCPU = CDbl(c2) / Factor                                  'Factor the cpu cycle count per actual second
  
'--Long loop test--------------------------------------------
  Call Clk.CpuClk(c1)                                       'Get the start cpu cycle count
    For li = 1 To 1000                                      'For loop using longs
      lj = li * LMULT
    Next li
  Call Clk.CpuClk(c2)                                       'Get the stop cpu cycle count
  cLong = c2 - c1 - cOver2                                  'Calculate the actual number of cpu cycles

'--Integer loop test-----------------------------------------
  Call Clk.CpuClk(c1)                                       'Get the start cpu cycle count
    For ii = 1 To 1000                                      'For loop using integers
      ij = ii * IMULT
    Next ii
  Call Clk.CpuClk(c2)                                       'Get the stop cpu cycle count
  cInteger = c2 - c1 - cOver2                               'Calculate the actual number of cpu cycles
  
'--Display results-------------------------------------------
  'Note: the results are displayed here rather than inline with the tests else calls into
  'the listview control could easily push the CpuClk code/data out of the cache. So all
  'results are saved into variables and displayed here when all the tests are completed.
  With lv.ListItems
    'Copy the last results to the previous results
    .Item(2).SubItems(3) = .Item(2).SubItems(1)
    .Item(2).SubItems(4) = .Item(2).SubItems(2)
    .Item(3).SubItems(3) = .Item(3).SubItems(1)
    .Item(3).SubItems(4) = .Item(3).SubItems(2)
    .Item(4).SubItems(3) = .Item(4).SubItems(1)
    .Item(4).SubItems(4) = .Item(4).SubItems(2)
    .Item(5).SubItems(3) = .Item(5).SubItems(1)
    .Item(5).SubItems(4) = .Item(5).SubItems(2)
    .Item(6).SubItems(3) = .Item(6).SubItems(1)
    .Item(6).SubItems(4) = .Item(6).SubItems(2)
    .Item(7).SubItems(3) = .Item(7).SubItems(1)
    .Item(7).SubItems(4) = .Item(7).SubItems(2)
    .Item(8).SubItems(3) = .Item(8).SubItems(1)
    .Item(8).SubItems(4) = .Item(8).SubItems(2)
    
    .Item(1).SubItems(1) = Format$(cCycles * CUR2INT, "#,###")                    'QueryPerformace frequency
    .Item(1).SubItems(2) = Format$((1# / (cCycles * CUR2INT)) * USEC, "0.0000")   'QueryPerformance resolution
    .Item(2).SubItems(1) = Format$(cCPU * CUR2INT, "#,###")                       'CPU clock frequency
    .Item(2).SubItems(2) = Format$((1# / (cCPU * CUR2INT)) * USEC, "0.0000")      'CPU cycle time
    .Item(3).SubItems(1) = Format(cOver1 * CUR2INT, "#,###")
    .Item(3).SubItems(2) = Format(cOver1 / cCPU * USEC, "0.0000")
    .Item(4).SubItems(1) = Format(cOver2 * CUR2INT, "#,###")
    .Item(4).SubItems(2) = Format(cOver2 / cCPU * USEC, "0.0000")
    .Item(5).SubItems(1) = Format(cQpc1 * CUR2INT, "#,###")
    .Item(5).SubItems(2) = Format(cQpc1 / cCPU * USEC, "0.0000")
    .Item(6).SubItems(1) = Format(cQpc2 * CUR2INT, "#,###")
    .Item(6).SubItems(2) = Format(cQpc2 / cCPU * USEC, "0.0000")
    .Item(7).SubItems(1) = Format(cLong * CUR2INT, "#,###")
    .Item(7).SubItems(2) = Format(cLong / cCPU * USEC, "0.0000")
    .Item(8).SubItems(1) = Format(cInteger * CUR2INT, "#,###")
    .Item(8).SubItems(2) = Format(cInteger / cCPU * USEC, "0.0000")
  End With
  
'--Cleanup---------------------------------------------------
  Screen.MousePointer = vbDefault
  cmdTest.Enabled = True
  Set Clk = Nothing
End Sub
