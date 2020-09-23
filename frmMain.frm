VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DocAttack"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   6345
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraCurrent 
      Caption         =   "Current Password"
      Height          =   1410
      Left            =   4230
      TabIndex        =   12
      Top             =   630
      Width           =   1995
      Begin VB.CommandButton btnAttack 
         Caption         =   "Attack"
         Height          =   285
         Left            =   540
         TabIndex        =   13
         Top             =   630
         Width           =   960
      End
      Begin VB.Label lblCurrent 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   240
         Left            =   135
         TabIndex        =   15
         Top             =   270
         UseMnemonic     =   0   'False
         Width           =   1770
      End
      Begin VB.Label lblProgress 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   240
         Left            =   135
         TabIndex        =   14
         Top             =   1035
         UseMnemonic     =   0   'False
         Width           =   1770
      End
   End
   Begin VB.Frame fraMakeup 
      Caption         =   "Makeup"
      Enabled         =   0   'False
      Height          =   1410
      Left            =   2160
      TabIndex        =   8
      Top             =   630
      Width           =   1995
      Begin VB.CheckBox chkSym 
         Caption         =   "Symbols"
         Enabled         =   0   'False
         Height          =   285
         Left            =   135
         TabIndex        =   11
         Top             =   900
         Value           =   2  'Grayed
         Width           =   1725
      End
      Begin VB.CheckBox chkLet 
         Caption         =   "Letters"
         Enabled         =   0   'False
         Height          =   285
         Left            =   135
         TabIndex        =   10
         Top             =   585
         Value           =   2  'Grayed
         Width           =   1725
      End
      Begin VB.CheckBox chkNum 
         Caption         =   "Numbers"
         Enabled         =   0   'False
         Height          =   285
         Left            =   135
         TabIndex        =   9
         Top             =   270
         Value           =   2  'Grayed
         Width           =   1725
      End
   End
   Begin VB.Frame fraSize 
      Caption         =   "Length"
      Height          =   1410
      Left            =   90
      TabIndex        =   3
      Top             =   630
      Width           =   1995
      Begin VB.TextBox txtMax 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   1125
         TabIndex        =   7
         Text            =   "9"
         Top             =   810
         Width           =   465
      End
      Begin VB.TextBox txtMin 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   1125
         TabIndex        =   6
         Text            =   "6"
         Top             =   405
         Width           =   465
      End
      Begin VB.Label lbl 
         Caption         =   "Maximum"
         Height          =   285
         Index           =   2
         Left            =   225
         TabIndex        =   5
         Top             =   833
         Width           =   1365
      End
      Begin VB.Label lbl 
         Caption         =   "Minimum"
         Height          =   285
         Index           =   1
         Left            =   225
         TabIndex        =   4
         Top             =   428
         Width           =   1365
      End
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   285
      Left            =   5670
      TabIndex        =   2
      Top             =   180
      Width           =   285
   End
   Begin VB.TextBox txtDocument 
      Height          =   285
      Left            =   1980
      TabIndex        =   1
      Top             =   180
      Width           =   3660
   End
   Begin VB.Timer tmrAttack 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   585
      Top             =   2340
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   45
      Top             =   2295
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lbl 
      Caption         =   "Protected Document:"
      Height          =   240
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   225
      Width           =   1590
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mobjWord As Object
Public mstrPassword As String
Public mstrFile As String
Public mintMin As Integer
Public mintMax As Integer
Dim PW() As Integer

Private Sub btnAttack_Click()
    Select Case btnAttack.Caption
        Case "Attack"
            ' Set the values in the variables
            If Initialize Then
                ' Create a word session
                Set mobjWord = CreateObject("Word.Application")
                ' Change caption on attack button
                 btnAttack.Caption = "Pause"
                 ' Change the caption on the progress label
                 lblProgress.Caption = "Working..."
                ' Enable the timer
                tmrAttack.Enabled = True
            End If
        Case "Resume"
            ' Pause the attack
            tmrAttack.Enabled = True
            ' Change caption on attack button
             btnAttack.Caption = "Pause"
             ' Change the caption on the progress label
             lblProgress.Caption = "Working..."
        Case "Pause"
            ' Pause the attack
            tmrAttack.Enabled = False
            ' Change caption on attack button
             btnAttack.Caption = "Resume"
             ' Change the caption on the progress label
             lblProgress.Caption = "Waiting..."
        Case "Exit"
            Unload Me
    End Select
End Sub

Private Function NextPassword() As Boolean
    NextPassword = Increment(1)
    UpdatePassword
End Function

Private Function Increment(intChr As Integer) As Boolean
    Increment = False
    ' Determine if chr can be incremented
    If PW(intChr) < 126 Then    ' last ascii code(~)
        PW(intChr) = PW(intChr) + 1
    Else
        ' Reset this chr to the first ascii 33(!)
        PW(intChr) = 33
        ' Determine if the max
        If intChr = mintMax Then
            ' Reached the maximum
            Increment = False
            ' Disable the timer
            tmrAttack.Enabled = False
            ' Update the display
            btnAttack.Caption = "Exit"
            lblProgress.Caption = "Failure"
            lblProgress.ForeColor = vbRed
        Else
            ' Increment the next chr
            Increment = Increment(intChr + 1)
        End If
    End If
End Function

Private Sub UpdatePassword()
    Dim intX As Integer
    mstrPassword = ""
    For intX = mintMax To 1 Step -1
        mstrPassword = mstrPassword & Chr(PW(intX))
    Next intX
    lblCurrent.Caption = mstrPassword
End Sub

Private Sub cmdBrowse_Click()
    ' Configure common dialog
    With cd1
        .DefaultExt = ".doc"
        .DialogTitle = "Select protected document"
        .Filter = ".doc"
        .ShowOpen
        ' Fill in document name
        txtDocument.Text = .FileName
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If btnAttack.Caption = "Attack" Then
        ' Close the application
        mobjWord.application.quit
    End If
    ' Remove the object from memory
    Set mobjWord = Nothing
    ' Unload this form
    Unload Me
End Sub

Private Sub tmrAttack_Timer()
    Dim objDoc As Object
    
On Error GoTo ErrHandler
    ' Attempt to open the document
    Set objDoc = mobjWord.Documents.Open(mstrFile, , , , LTrim(mstrPassword))
    ' Successful attack
    tmrAttack.Enabled = False
    ' Close the document
    mobjWord.Documents.Close
    ' Remove objects from memory
    Set objDoc = Nothing
    ' Update the display
    btnAttack.Caption = "Exit"
    ' Change the caption on the progress label
    lblProgress.Caption = "Successful"
Exit Sub

ErrHandler:
    ' Determine if password error
    If Err.Number = 5408 Then
        ' Clear the error
        Err.Clear
        ' increment the password
        NextPassword
    End If
End Sub

Private Function Initialize() As Boolean
    Dim intX As Integer
    Dim blnErr As Boolean
    ' Validate document
    If txtDocument.Text = "" Then
        MsgBox "Please select a document prior to attempting an attack.", vbCritical + vbOKOnly
        blnErr = True
    Else
        ' Pull the document
        mstrFile = txtDocument.Text
    End If
    ' Validate minimum
    If txtMin.Text = "" Then
        MsgBox "Please select a minimum length prior to attempting an attack.", vbCritical + vbOKOnly
        blnErr = True
    Else
        ' Pull the minimum password length
        mintMin = CInt(txtMin.Text)
    End If
    ' Validate maximum
    If txtMax.Text = "" Then
        MsgBox "Please select a maximum length prior to attempting an attack.", vbCritical + vbOKOnly
        blnErr = True
    Else
        ' Pull the minimum password length
        mintMax = CInt(txtMax.Text)
    End If
    ' Ensure that maximum is greater than minimum
    If mintMax < mintMin Then
        MsgBox "Please select a maximum length that is greater than or equal to the minimum prior to attempting an attack.", vbCritical + vbOKOnly
        blnErr = True
    End If
    If Not blnErr Then
        ' Configure password array
        ReDim PW(1 To mintMax) As Integer
        ' Load the password array with spaces
        For intX = 1 To mintMax
            PW(intX) = 32 'space ascii code( )
        Next intX
        ' Load the password array for the minimum
        For intX = 1 To mintMin
            PW(intX) = 33 'first ascii code(!)
        Next intX
        ' Update the password
        UpdatePassword
    End If
    Initialize = Not blnErr
End Function

Private Sub txtMax_KeyPress(KeyAscii As Integer)
    ' Limit user to numeric keystroke (and the backspace)
    If ((KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8) Then
        MsgBox "Please limit to this value to only numbers", vbCritical + vbOKOnly
        KeyAscii = 0
    End If
End Sub

Private Sub txtMin_KeyPress(KeyAscii As Integer)
    ' Limit user to numeric keystroke (and the backspace)
    If ((KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8) Then
        MsgBox "Please limit to this value to only numbers", vbCritical + vbOKOnly
        KeyAscii = 0
    End If
End Sub
