Attribute VB_Name = "mTxtBin"
'--------------------------------------------------
' E-Mail Checker
' By H G Laughland
'
' Text/Binary Conversion Module
' Used to protect login details.
'
' Code obtained from Planet Source Code
' This is not my code!
'--------------------------------------------------

Public Function ChrAscii(Char As String) As Long
 Dim GetAscii&
  For GetAscii& = 0 To 255
    If Mid(Char$, 1, 1) = Chr(GetAscii) Then
      ChrAscii = GetAscii
      Exit Function
    End If
  Next GetAscii&
End Function
'Double Click on the user control, and i
'     n the General Declarations
'Put this... these are the subs the you
'     will use


Public Function TextToBinary(StringT As String) As String
 Dim Ascii, FinalBinary$, GetNum&
 FinalBinary$ = ""

 For GetNum& = 1 To Len(StringT$)
  Ascii = ChrAscii(Mid(StringT$, GetNum, 1))
  ' 128
  If Ascii >= 128 Then
    FinalBinary$ = FinalBinary$ & "1"
    Ascii = Ascii - 128
  Else
   FinalBinary$ = FinalBinary$ & "0"
  End If
        
  ' 64
  If Ascii >= 64 Then
   FinalBinary$ = FinalBinary$ & "1"
   Ascii = Ascii - 64
  Else
   FinalBinary$ = FinalBinary$ & "0"
  End If
        
  ' 32
  If Ascii >= 32 Then
   FinalBinary$ = FinalBinary$ & "1"
   Ascii = Ascii - 32
  Else
   FinalBinary$ = FinalBinary$ & "0"
  End If
        
  ' 16
  If Ascii >= 16 Then
   FinalBinary$ = FinalBinary$ & "1"
   Ascii = Ascii - 16
  Else
   FinalBinary$ = FinalBinary$ & "0"
  End If
        
  ' 8
  If Ascii >= 8 Then
   FinalBinary$ = FinalBinary$ & "1"
   Ascii = Ascii - 8
  Else
   FinalBinary$ = FinalBinary$ & "0"
  End If
        
  ' 4
  If Ascii >= 4 Then
   FinalBinary$ = FinalBinary$ & "1"
   Ascii = Ascii - 4
  Else
   FinalBinary$ = FinalBinary$ & "0"
  End If
        
  ' 2
  If Ascii >= 2 Then
   FinalBinary$ = FinalBinary$ & "1"
   Ascii = Ascii - 2
  Else
   FinalBinary$ = FinalBinary$ & "0"
  End If
        
  ' 1
  If Ascii >= 1 Then
   FinalBinary$ = FinalBinary$ & "1"
   Ascii = Ascii - 1
  Else
   FinalBinary$ = FinalBinary$ & "0"
  End If

  If Mid(StringT$, GetNum + 1, 1) = Chr(32) Then
   FinalBinary$ = FinalBinary$ '& " "
  Else
   FinalBinary$ = FinalBinary$ '& Chr(32)
  End If

 Next GetNum&
 TextToBinary$ = FinalBinary$
End Function


Public Function BinaryToText(BinaryString As String) As String
Dim GetBinary&, Num$, Binary&, FinalString$, NewString$

NextChr:

For GetBinary& = 1 To 8
    Num$ = Mid(BinaryString$, GetBinary&, 1)
    Select Case Num$
     Case "1"
        If GetBinary = 1 Then
            Binary = Binary + 128
        ElseIf GetBinary = 2 Then
            Binary = Binary + 64
        ElseIf GetBinary = 3 Then
            Binary = Binary + 32
        ElseIf GetBinary = 4 Then
            Binary = Binary + 16
        ElseIf GetBinary = 5 Then
            Binary = Binary + 8
        ElseIf GetBinary = 6 Then
            Binary = Binary + 4
        ElseIf GetBinary = 7 Then
            Binary = Binary + 2
        ElseIf GetBinary = 8 Then
            Binary = Binary + 1
        End If
    End Select

Next GetBinary&
FinalString$ = FinalString$ & Chr(Binary)
NewString$ = Mid(BinaryString$, 9)

If NewString$ = "" Then
 BinaryToText$ = FinalString$
Else
 BinaryString$ = NewString$
 Binary = 0
 GoTo NextChr
End If
End Function
