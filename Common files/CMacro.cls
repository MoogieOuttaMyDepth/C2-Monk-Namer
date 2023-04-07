VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMacro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File:     CMacro.cls
' Date:     30/01/98
' Author:   Robin E. Charlton.
' Edited by GimmeCat 07/04/2023
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Const fMacro = 1
Const OLEServer = "SFC2.OLE"

'local variable(s) to hold property value(s)
Private mvarMacro As Long 'local copy
Private mvarCmd As String 'local copy
Private mvarResult As String 'local copy
Private mvarResultIndex As Integer 'local copy
Private mvarHost As Object 'local copy


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public Let Cmd
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let Cmd(ByVal vData As String)
    On Error GoTo ErrorHandler
    
    mvarCmd = vData
    Call mvarHost.LoadCommand(mvarMacro, mvarCmd)
   
    Exit Property
    
ErrorHandler:
    Call MsgBox("Could create macro command.", vbCritical Or vbOKOnly)
    Call Error(Err.Number)
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public Get Cmd
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get Cmd() As String
Attribute Cmd.VB_UserMemId = 0
    Cmd = mvarCmd
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public: Get Result.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get result() As String
    result = mvarResult
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private: Initialize.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    On Error GoTo ErrorHandler
    
    Set mvarHost = CreateObject(OLEServer)
    mvarMacro = mvarHost.CreateCommand(fMacro)
    If (mvarMacro = 0) Then Call Error
    
    Exit Sub
    
ErrorHandler:
    Call MsgBox("Could not construct interface to Creatures.", vbCritical Or vbOKOnly)
    Call Error(Err.Number)
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private: Terminate.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
    On Error GoTo ErrorHandler
    
    If (mvarMacro <> 0) Then
        If (Not mvarHost.DestroyCommand(mvarMacro)) Then
            Call Error
        End If
    End If
    
    Set mvarHost = Nothing
    Exit Sub
    
ErrorHandler:
    Call MsgBox("An error occurred when destroying the macro command.", vbCritical Or vbOKOnly)
    'Call Error(Err.Number)
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public: Execute.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Execute()
    On Error GoTo ErrorHandler
    
    If (Not mvarHost.RequestCommand(mvarMacro, mvarResult)) Then
        Call Error
    End If
    
    mvarResultIndex = 1
    Exit Sub
    
ErrorHandler:
    Call MsgBox("Could not execute macro.", vbCritical Or vbOKOnly)
    Error Err.Number
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public: NextInResult.
' Parse next data item from the result string.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NextInResult() As Variant

    Dim Length As Integer
    
    Length = Len(mvarResult)
    If (mvarResultIndex > Length) Then
        NextInResult = ""
        Exit Function
    End If
    
    Dim i%, c$
    For i% = mvarResultIndex To Length
    
        c$ = Mid$(mvarResult, i%, 1)
        If (c$ = "|") Then
            Exit For
        End If
    Next i%
    
    NextInResult = Mid$(mvarResult, mvarResultIndex, i% - mvarResultIndex)
    mvarResultIndex = i% + 1

End Function

Public Function hex2ascii(ByVal hextext As String) As String

Dim y As Integer
Dim Value, num As String

For y = 1 To Len(hextext)
    num = Mid(hextext, y, 2)
    Value = Value & Chr(Val("&h" & num))
    y = y + 1
Next y

hex2ascii = Value
End Function

Public Function strReverse_Character_Pairs(ByVal strValue As String) As String

  Dim lngLoop As Long
  Dim strReturn As String

  strReturn = ""

  For lngLoop = Len(strValue) - 1& To 1& Step -2&
      strReturn = strReturn & Mid$(strValue, lngLoop, 2)
  Next lngLoop

  strReverse_Character_Pairs = strReturn
End Function
