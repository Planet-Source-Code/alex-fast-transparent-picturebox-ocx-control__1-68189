VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'* CODED BY: BattleStorm
'* EMAIL: battlestorm@cox.net
'* UPDATED: 08/02/2002
'* PURPOSE: Used to time how long code
'*     takes to execute with a great
'*     deal of accuracy. Will return
'*     value in milliseconds accurate
'*     to 2 decimal places.
'* COPYRIGHT: This program and source
'*     code is freeware and can be copied
'*     and/or distributed as long as you
'*     mention the original author. I am
'*     not responsible for any harm as the
'*     outcome of using any of this code.

Option Explicit

'API calls
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

'Variables
Private m_HasCounter As Boolean
Private m_Freq As Currency
Private m_Start As Currency
Private m_Stop As Currency
Private m_Elapsed As Currency

'Check for performance counter and set frequency
Private Sub Class_Initialize()
  m_HasCounter = QueryPerformanceFrequency(m_Freq)
End Sub

'Start timer
Public Sub StartTimer()
  QueryPerformanceCounter m_Start
End Sub

'Stop timer
Public Sub StopTimer()
  QueryPerformanceCounter m_Stop
End Sub

'Returns a string representing elasped time in milliseconds accurate to 2 decimal digits
Public Property Get Elasped() As String
  m_Elapsed = (((m_Stop - m_Start) / m_Freq) * 1000)
  Elasped = Format(m_Elapsed, "###,###,##0.00")
End Property
