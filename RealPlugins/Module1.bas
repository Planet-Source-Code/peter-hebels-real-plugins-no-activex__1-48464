Attribute VB_Name = "Module1"
' This example is written by Peter Hebels, website: http://www.phsoft.nl
' The author of this code cannot be held responsible for any damages may caused by the
' use of this code.

Option Explicit

'Some API funcions.
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cBytes As Long)

Private MlngParameters() As Long
Private MlngAddress As Long
Private MlngCP As Long
Private MbytCode() As Byte

'This function loads the dll and passes the function params to it
Public Function LoadDll(DllFileName As String, FunctionToCall As String, ParamArray FuncParams()) As Long
  Dim DllId As Long
  Dim I As Long
  
  ReDim MlngParameters(0)
  ReDim MbytCode(0)
  
  MlngAddress = 0
  
  'Load the library, the number returned is later used to get the ProcAdress
  DllId = LoadLibrary(ByVal DllFileName)
  If DllId = 0 Then
     'If there was a error loading the dll then show a message and exit the funcion.
     MsgBox "DLL not found", vbCritical
     Exit Function
  End If
  
  MlngAddress = GetProcAddress(DllId, ByVal FunctionToCall)
  If MlngAddress = 0 Then
     'If there was an error while calling the function within the dll then show a message,
     'unload the dll and exit the function.
     MsgBox "Function entry not found", vbCritical
     FreeLibrary DllId
     Exit Function
  End If
  
  'Create a array of params that are passed to the dll.
  ReDim MlngParameters(UBound(FuncParams) + 1)
  For I = 1 To UBound(MlngParameters)
     MlngParameters(I) = CLng(FuncParams(I - 1))
  Next I
  
  'Used to pass params to the dll, it also recieves the returned messages.
  LoadDll = CallWindowProc(PrepareCode, 0, 0, 0, 0)
  
  'Unload the dll, this is importand to do because if you don't your application will crash
  'when you close it.
  FreeLibrary DllId
End Function

'This code is used to pass params to the dll, it also recieves the returned messages from
'the dll. This is also the most difficult part of this project :)
Private Function PrepareCode() As Long
   Dim lngX As Long
   Dim CodeStart As Long
   
   ReDim MbytCode(18 + 32 + 6 * UBound(MlngParameters))
   
   CodeStart = GetAlignedCodeStart(VarPtr(MbytCode(0)))
   MlngCP = CodeStart - VarPtr(MbytCode(0))
   For lngX = 0 To MlngCP - 1
       MbytCode(lngX) = &HCC
   Next
   
   'Do some Assembly here
   vbAddByteToCode &H58 'pop eax
   vbAddByteToCode &H59 'pop ecx
   vbAddByteToCode &H59 'pop ecx
   vbAddByteToCode &H59 'pop ecx
   vbAddByteToCode &H59 'pop ecx
   vbAddByteToCode &H50 'push eax
   
   For lngX = UBound(MlngParameters) To 1 Step -1
       vbAddByteToCode &H68 'push wwxxyyzz
       vbAddLongToCode MlngParameters(lngX)
   Next
   
   vbAddCallToCode MlngAddress
   vbAddByteToCode &HC3
   vbAddByteToCode &HCC
   PrepareCode = CodeStart
End Function

Private Sub vbAddCallToCode(lngAddress As Long)
   vbAddByteToCode &HE8
   vbAddLongToCode lngAddress - VarPtr(MbytCode(MlngCP)) - 4
End Sub

Private Sub vbAddLongToCode(lng As Long)
   Dim intX As Integer
   Dim byt(3) As Byte
   CopyMemory byt(0), lng, 4
   For intX = 0 To 3
       vbAddByteToCode byt(intX)
   Next
End Sub

Private Sub vbAddByteToCode(byt As Byte)
   MbytCode(MlngCP) = byt
   MlngCP = MlngCP + 1
End Sub

Private Function GetAlignedCodeStart(lngAddress As Long) As Long
   GetAlignedCodeStart = lngAddress + (15 - (lngAddress - 1) Mod 16)
   If (15 - (lngAddress - 1) Mod 16) = 0 Then GetAlignedCodeStart = GetAlignedCodeStart + 16
End Function
