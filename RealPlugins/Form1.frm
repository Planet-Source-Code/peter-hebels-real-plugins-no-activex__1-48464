VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Plugin test"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   3015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Select a plugin:"
      Height          =   1455
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2775
      Begin VB.CommandButton Command2 
         Caption         =   "About plugin"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1120
         Width           =   2535
      End
      Begin VB.ListBox List1 
         Height          =   840
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calculate"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Caption         =   "Enter some numbers:"
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   2775
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.FileListBox File1 
      Height          =   870
      Left            =   120
      TabIndex        =   5
      Top             =   2880
      Visible         =   0   'False
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This example is written by Peter Hebels, website: http://www.phsoft.nl
' The author of this code can not be held responsible for any damages may caused by the
' use of this code.

Option Explicit

Dim DllName As String
Const Number$ = "0123456789"

Private Sub Command1_Click()
    Dim StrFunction As String
    
    'Make sure there is a plugin selected
    If List1.Text = "" Then
        'If no plugin is selected a message is shown
        MsgBox "No plugin selected, please select a plugin from the list first.", vbExclamation, "Error"
        Exit Sub
    End If
    
    'Here we check if there are some numbers entered into the textboxes
    If Text1.Text = "" Then
        'Show a message when no numbers are found
        MsgBox "Please enter a number in the first textbox.", vbExclamation, "Error"
        Exit Sub
    ElseIf Text2.Text = "" Then
        MsgBox "Please enter a number in the second textbox.", vbExclamation, "Error"
        Exit Sub
    End If
    
    'Here we check what the dll does, if 1 is returned it multiplys if 2 is returned it
    'divides
    If LoadDll(DllName, "GetFunct") = 1 Then
        StrFunction = " * "
    ElseIf LoadDll(DllName, "GetFunct") = 2 Then StrFunction = " / "
        StrFunction = " / "
    End If
    
    'Show the awnser to the user.
    MsgBox "Awnser from plugin: " & Text1 & StrFunction & Text2 & " = " & LoadDll(DllName, "CalculateFunct", Text1, Text2)
End Sub

Private Sub Command2_Click()
If List1.Text = "" Then
    MsgBox "Please select a plugin first.", vbExclamation, "Error"
    Exit Sub
End If

'Show the about box.
LoadDll DllName, "ShowAbout", Me.hwnd
End Sub

Private Sub Form_Load()
Dim ThePluginName As String
Dim I As Long
On Error Resume Next
    'Look for dll files with filenames that start with 'plg' and add them to the
    'listbox.
    File1.Pattern = "*.dll"
    File1.Path = App.Path
    
    For I = 0 To File1.ListCount
        ThePluginName = Left$(File1.List(I), 3)
        
        If StrConv(ThePluginName, vbLowerCase) = "plg" Then
            List1.AddItem (File1.List(I))
        End If
    Next I
End Sub

Private Sub List1_Click()
    'Here the dll's filename is saved into the string variable.
    DllName = List1.Text
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    'We can only calculate numbers, other chars are ignored.
    If KeyAscii <> 8 Then
        If InStr(Number$, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 8 Then
        If InStr(Number$, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub
