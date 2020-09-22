VERSION 5.00
Begin VB.Form frmQueryPerformanceCounter 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Query Performance Counter - By Endra"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   5445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkExactTime 
      Caption         =   "Exact Time"
      Height          =   495
      Left            =   4200
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Calculate QueryPerformanceCounter"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3975
   End
   Begin VB.Label lblResult 
      Caption         =   "Result will appear here."
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
   End
End
Attribute VB_Name = "frmQueryPerformanceCounter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'This is another and better way of calculating the time that the computer took to perform some operations.
'Except, in this one, we're using Query Performance Counter, which is included starting from the 386 cpu.
'
'To be real, QueryPerformanceCounter makes GetTickCount look like something not very fast.
'Here are a couple reasons why:
'
'   1)It accesses the CPU's high performance counter which changes its tick value much more frequently than the
'   Windows system timer. This allows us to resolve differences on the order of microseconds (10^-6), rather than
'   milliseconds (10^-3)!
'
'   2) Another important application of QueryPerformanceCounter is, as its name implies, performance timing.
'   Try using GetTickCount to tell you how long it takes to make an API call. It can't! Most simple functions
'   are processed in mere microseconds. GetTickCount would detect no change. Starting to see its limitations?
'   If you're testing two routines to determine which is faster, you want the best resolution you can get.

'** Declarations **
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

'** Variables **
Dim curFreq As Currency
Dim curStart As Currency
Dim curEnd As Currency
Dim dblResult As Double

'** Subs and Events **

Private Sub cmdStart_Click()
    
    Dim i As Long
    Dim x As Long
    Dim lLoops As Long
    
    lLoops = 500000 'how many times to loop
    
    QueryPerformanceFrequency curFreq 'Get the timer frequency
    QueryPerformanceCounter curStart 'Get the start time

    'the code to test goes here!
    For i = 1 To lLoops
        x = x + 1
        x = x - 1
    Next i
    'stop the code to test here!
    
    QueryPerformanceCounter curEnd 'Get the end time
    dblResult = (curEnd - curStart) / curFreq 'Calculate the duration (in seconds)
    
    If chkExactTime.Value = vbChecked Then
    
        lblResult.Caption = "Looped " & lLoops & " times in: " & dblResult & " seconds." & vbNewLine
    
    Else
        
        lblResult.Caption = "Looped " & lLoops & " times in approximatly: " & Format(dblResult, "0000.0000") & " seconds."

    End If
End Sub
