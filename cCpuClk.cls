VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCpuClk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
' cCpuClk class - allows the user to execute the Pentium class RDTSC op-code
' to retrieve the CPU clock cycle count.
'
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long

Private hMem      As Long                                     'Allocated memory handle
Private pOldCode  As Long                                     'Pointer to the original (listed) CpuClk routine
Private pNewCode  As Long                                     'Pointer to the machine code replacement code
Private pMyself   As Long                                     'Pointer to me, the class instance

'This routine will be replaced with a machine code subroutine when the class is created
Public Sub CpuClk(Cycles As Currency)
End Sub

'Replace the CpuClk routine in the vtable with a machine code sub
Private Sub Class_Initialize()
  hMem = GlobalAlloc(0, 16)                                   'Allocate memory for m/c and save handle
  pNewCode = GlobalLock(hMem)                                 'Get pointer to the locked/allocated memory
  Call CopyMemory(pMyself, ByVal ObjPtr(Me), 4)               'Get pointer to object Me
  Call CopyMemory(ByVal pNewCode + &H0, &H4C8B310F, 4)        'Copy m/c code to allocated memory
  Call CopyMemory(ByVal pNewCode + &H4, &H1890824, 4)         'Copy m/c code to allocated memory
  Call CopyMemory(ByVal pNewCode + &H8, &H33045189, 4)        'Copy m/c code to allocated memory
  Call CopyMemory(ByVal pNewCode + &HC, &H8C2C0, 4)           'Copy m/c code to allocated memory
  Call CopyMemory(pOldCode, ByVal pMyself + &H1C, 4)          'Get old CpuClk VT entry
  Call CopyMemory(ByVal pMyself + &H1C, pNewCode, 4)          'Put new CpuClk VT entry
  Call GlobalUnlock(hMem)                                     'Unlock memory
End Sub

'Restore the original CpuClk routine in the vtable
Private Sub Class_Terminate()
  Call CopyMemory(ByVal pMyself + &H1C, pOldCode, 4)          'Restore old CpuClk VT entry
  Call GlobalFree(hMem)                                       'Release memory used for m/c
End Sub
