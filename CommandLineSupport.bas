Attribute VB_Name = "CommandLineSupport"
Option Explicit

Private Declare Function GetCommandLineW Lib "kernel32.dll" () As Long
Private Declare Function CommandLineToArgv Lib "shell32" Alias "CommandLineToArgvW" (ByVal lpCmdLine As Long, pNumArgs As Long) As Long

Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long

Public Function GetFirstCommandIfAny() As String

Dim strReturnArray() As String
    strReturnArray = GetArguments(GetCommandLineW)
    
    If UBound(strReturnArray) > 0 Then
        GetFirstCommandIfAny = strReturnArray(1)
    End If

End Function

Private Function GetArguments(lpCmdLine As Long) As String()

Dim m_Arguments() As String
Dim lpData As Long
Dim nArgs As Long
Dim lpArgument As Long
Dim i As Integer

   ' which is an array of Unicode strings.
   lpData = CommandLineToArgv(lpCmdLine, nArgs)
   If lpData Then
      ReDim m_Arguments(0 To nArgs - 1) As String
      ' Extract individual arguments from array, starting
      ' with element 1, because 0 contains the potentially
      ' unqualified appname.
      For i = 1 To nArgs - 1
         lpArgument = PointerToDWord(lpData + (i * 4))
         m_Arguments(i) = PointerToStringW(lpArgument)
      Next i
   End If
   Call GlobalFree(lpData)
   
   GetArguments = m_Arguments()
    
End Function

Private Function GetCommandLine() As Byte()
    
Dim ptrCommand As Long
Dim ptrLength As Long

Dim bytReturn() As Byte

    ptrCommand = GetCommandLineW
    ptrLength = lstrlenW(ptrCommand) * 2
    
    ReDim bytReturn(ptrLength) As Byte
    CopyMemory bytReturn(0), ByVal ptrCommand, ptrLength
    
    GetCommandLine = bytReturn
    
End Function

Private Function PointerToDWord(ByVal lpDWord As Long) As Long
   Dim nRet As Long
   If lpDWord Then
      CopyMemory nRet, ByVal lpDWord, 4
      PointerToDWord = nRet
   End If
End Function

Private Function PointerToStringW(lpStringW As Long) As String
   Dim Buffer() As Byte
   Dim nLen As Long
   
   If lpStringW Then
      nLen = lstrlenW(lpStringW) * 2
      If nLen Then
         ReDim Buffer(0 To (nLen - 1)) As Byte
         CopyMemory Buffer(0), ByVal lpStringW, nLen
         PointerToStringW = Buffer
      End If
   End If
End Function
