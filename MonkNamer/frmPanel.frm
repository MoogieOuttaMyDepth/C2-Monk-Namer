VERSION 5.00
Begin VB.Form frmPanel 
   Caption         =   "C2 Monk Namer"
   ClientHeight    =   930
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2385
   Icon            =   "frmPanel.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   930
   ScaleWidth      =   2385
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   2040
      TabIndex        =   3
      Top             =   110
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   455
      Left            =   370
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "(None yet!)"
      Top             =   400
      Width           =   1600
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   182
      Top             =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Zoom to:"
      Height          =   195
      Left            =   1300
      TabIndex        =   2
      Top             =   120
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Last named:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   870
   End
End
Attribute VB_Name = "frmPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Creatures As Object
Dim Interval As Integer
Dim arrSplitString() As String
Dim Macro As CMacro

Private Sub Form_Load()
    On Error GoTo ErrorHandler
    Set Macro = New CMacro
    Timer1.Enabled = True
    Check1.Value = 0
    Exit Sub
    
ErrorHandler:
    Call MsgBox("Could not establish communication with Creatures.", vbCritical Or vbOKOnly)
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set Macro = Nothing
End Sub

Private Sub Timer1_Timer()
    On Error GoTo ErrorHandler
    Interval = Interval + 1
    'Text1.Text = Interval
    If (Interval > 2000) Then
        Macro.Cmd = "doif totl 4 0 0 gt 0 enum 4 0 0 dde: putv targ dde: putv monk dde: getb cnam next endi"
        Call Macro.Execute
        
        If (Macro.result <> "") Then
            'Text2.Text = Macro.result (debug only)
            arrSplitString = Split(Macro.result, "|")
            
        Else
            Text1.Text = "No data returned from host."
        End If
        
        'vars moved here so they get cleared properly even if errors
        Dim intMonk, indx As Integer
        Dim strMonk, strString, strNorn, strHex As String
        
        For indx = LBound(arrSplitString) To UBound(arrSplitString)
            If arrSplitString(indx) = "<UnNamed>" Then
                'Get UnNamed's Moniker
                strMonk = arrSplitString(indx - 1)
                'Get UnNamed's UNID
                strNorn = arrSplitString(indx - 2)
                'zoom to newly named or not
                If Check1.Value = 1 Then
                    Macro.Cmd = "targ " & strNorn & " setv norn targ dde: panc"
                    Call Macro.Execute
                End If
                'Moniker string -> int
                intMonk = Val(strMonk)
                'Moniker numbers -> hex string
                strHex = Hex$(intMonk)
                'Swap endian of hex
                strMonk = Macro.strReverse_Character_Pairs(strHex)
                'Convert hex to ascii chars
                strString = Macro.hex2ascii(strMonk)
                'Construct and send naming macro
                Macro.Cmd = "targ " & strNorn & " dde: putb [" & strString & "] cnam"
                Text1.Text = strString
                Call Macro.Execute
                'To name 1 at a time, if multiple unnameds found at once
                Exit For
            End If
        Next
        Interval = 0
    End If
Exit Sub
    
ErrorHandler:
    Call MsgBox("Connection to the game was interrupted.", vbCritical Or vbOKOnly)
    Text1.Text = "Error!"
    Macro.Cmd = ""
    Interval = 0
End Sub
