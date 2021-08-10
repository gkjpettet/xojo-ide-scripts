/// Sets the description of a method selected in the navigator
/// to the description within the doc comments at the top of the
/// method body.
///
/// Makes a number of assumptions:
/// 1. That a method is selected in the navigator.
/// 2. That the method body begins with a doc comment. Doc
/// comments begin with `///`. The description is taken to be
/// all of the contiguous lines of text at the top of method
/// body.

// Copy the currently selected item in the navigator's
// declaration to the clipboard
DoCommand("Copy")

// Split the declaration into trimmed lines.
Var decLines() As String = Clipboard.Split(EndOfLine)
For i As Integer = 0 To decLines.LastIndex
decLines(i) = decLines(i).Trim
Next i

// Check the signature to ensure it's a method.
// Exit early with a warning if it's not.
Var signature As String = decLines(0)
If signature.IndexOf("Sub ") = -1 And signature.IndexOf("Function ") = -1 Then
Print("This script only works when a method is selected in the navigator.")
Return
End If

// Get the method body contents as an array of lines.
Var lines() As String = Text.Split(EndOfLine)

// Get the index of the first non-empty line.
Var firstLineIndex As Integer = -1
Var empty As Boolean = True
For firstLineIndex = 0 To lines.LastIndex
If lines(firstLineIndex).Trim.Length > 0 Then
empty = False
Exit
End If
Next firstLineIndex

If empty Then
// The method body is empty. Clear the item's description and exit.
ItemDescription = ""
Return
End If

// Make sure that the first non-empty line is a doc comment.
If Not lines(firstLineIndex).BeginsWith("///") Then
// No item description to set.
ItemDescription = ""
Return
End If

// Get all contiguous non-empty doc comments.
Var description As String = lines(firstLineIndex).ReplaceAll("///", "").Trim
For i As Integer = firstLineIndex + 1 To lines.LastIndex
Var line As String = lines(i).Trim
If line.BeginsWith("///") And line.Length > 3 Then
line = line.Replace("///", "").Trim
description = description + " " + line
Else
Exit
End If
Next i

ItemDescription = description
