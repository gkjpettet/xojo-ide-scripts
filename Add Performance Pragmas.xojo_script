/// Inserts performance increasing pragma at the top of the current method, after any header text.

// Get the current method contents as lines.
Var current() As String = Text.Split(EndOfLine)

// Construct the pragmas to insert.
Var pragmas() As String
pragmas.Add("#If Not TargetWeb")
pragmas.Add("#Pragma DisableBackgroundTasks")
pragmas.Add("#EndIf")
pragmas.Add("#Pragma NilObjectChecking False")
pragmas.Add("#Pragma StackOverflowChecking False")
pragmas.Add("#Pragma DisableBoundsChecking")

// We want to insert the pragmas at the top of the method, after any comment header that may be
// present. My comment headers begin with `///`.

// Find the index of the last comment header line.
Var lastHeaderLine As Integer = 0
For i As Integer = 0 To current.LastIndex
lastHeaderLine = i
If current(i).Length < 3 Or current(i).Left(3) <> "///" Then
Exit
End If
Next i

// Insert the pragmas after the comment header.
Var pragmaEnd As Integer
If lastHeaderLine = current.LastIndex Then
pragmas.AddAt(0, "")
For Each prag As String In pragmas
current.Add(prag)
Next prag
pragmaEnd = current.LastIndex
Else
// If the line after the comment header is blank, insert the pragmas straight afterwards.
If current(lastHeaderLine).Trim = "" Then
lastHeaderLine = lastHeaderLine + 1
Else
// Insert a blank line before the pragmas.
pragmas.AddAt(0, "")
End If
For Each prag As String In pragmas
current.AddAt(lastHeaderLine, prag)
lastHeaderLine = lastHeaderLine + 1
Next prag
pragmaEnd = lastHeaderLine - 1
End If

// Ensure there is a blank line after the inserted pragmas.
If pragmaEnd = current.LastIndex Then
current.Add("")
Else
If current(pragmaEnd + 1).Trim <> "" Then
current.AddAt(pragmaEnd + 1, "")
End If
End If

Text = String.FromArray(current, EndOfLine)


