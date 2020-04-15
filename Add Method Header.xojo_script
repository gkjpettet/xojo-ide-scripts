Class Parameter
// Property declarations.
Dim Name As String
Dim type As String
Dim HasDefaultValue As Boolean = False
Dim DefaultValue As String = ""

Sub Constructor(name As String, type As String, hasDefaultValue As Boolean, defaultValue As String)
Self.Name = name
Self.Type = type
Self.HasDefaultValue = hasDefaultValue
Self.DefaultValue = defaultValue
End Sub

Function ToString() As String
Return "Name: " + Name + ", Type: " + Type + _
", HasDefault: " + If(HasDefaultValue = True, "True", "False") + _
", DefaultValue: " + DefaultValue 
End Function
End Class

// Copy the currently selected item in the navigator.
DoCommand("Copy")

// Get the signature.
// E.g: Public Function Create(file As FolderItem) As Boolean
Dim lines() As String

// Set the newline character.
Dim EOL As String = Chr(13)

// Split into lines and get the first one (the signature).
lines = split(clipboard, EOL)
Dim signature As String = lines(0)

// Functions return values, Subs do not.
Dim subStartPos As Integer = signature.InStr("Sub")
Dim functionStartPos As Integer = signature.InStr("Function")
Dim hasReturnValue As Boolean = If(functionStartPos > subStartPos, True, False)
If functionStartPos = 0 And subStartPos = 0 Then
Print("This script only works when a method is selected in the navigator")
Return
End If
If subStartPos > 0 Then
// A sub.
// Remove everything upto and including "Sub" from the start of the signature.
signature = signature.Right(signature.Len - subStartPos - 2).Trim
Else
// A function.
// Remove everything upto and including "Function" from the start of the signature.
signature = signature.Right(signature.Len - functionStartPos - 7).Trim
End If

// Get the method name.
Dim leftParenPos As Integer = signature.InStr("(")
Dim methodName As String = signature.Left(leftParenPos - 1)

// Remove the method name and left parenthesis from the signature.
signature = signature.Replace(methodName + "(", "")

// We need to find the closing parenthesis for the parameters. We can't just 
// search for the first `)` as the parameters may include arrays which contain 
// parentheses. What we'll do is search from the end of the string but there is 
// an edge case which is if the function returns an array...
// Does the function return an array?
Dim returnsArrayValue As Boolean = False
If hasReturnValue Then
If signature.Right(2) = "()" Then
returnsArrayValue = True
End If
End If

Dim paramClosingParenPos As Integer
Dim sigChars() As String = signature.Split("")
For i As Integer = If(returnsArrayValue, sigChars.LastRowIndex - 1, sigChars.LastRowIndex) DownTo 0
If sigChars(i) = ")" Then
paramClosingParenPos = i
Exit
End If
Next i

// Get the parameters (all characters up to the closing parenthesis).
Dim paramString As String = signature.Left(paramClosingParenPos)

// Convert the parameters into an array of Parameter objects.
Dim paramStrings() As String = Split(paramString, ",")
Dim params() As Parameter
Dim name, type, defaultValue As String
Dim hasDefault As Boolean
For Each p As String In paramStrings
name = p.Left(p.InStr(" As") - 1).Trim
p = p.Replace(name + " As ", "")

If p.InStr("=") = 0 Then
hasDefault = False
type = p
defaultValue = ""
Else
hasDefault = True
type = p.Left(p.InStr(" =") - 1).Trim
p = p.Replace(type + " = ", "")
defaultValue = p.Trim
End If

params.Append(New Parameter(name, type, hasDefault, defaultValue))
Next p

// Remove the parameters and closing parenthesis from the signature.
signature = signature.Replace(paramString + ")", "")

Dim returnValue As String = ""
If hasReturnValue Then
returnValue = signature.Replace("As", "").Trim
End If

// Build the comment header.
// =========================
// Description
Dim header As String
header = header + "///" + Chr(13)
header = header + "' DESCRIPTION" + Chr(13)

// Parameters.
If params.Ubound >= 0 Then
header = header + "'" + Chr(13)
End If
For Each p As Parameter In params// ' - Parameter PARAM_NAME: PARAM_DESCRIPTION
header = header + "' - Parameter " + p.Name + ": " + "PARAM_DESCRIPTION" + Chr(13)
Next p

// Return value
If hasReturnValue Then
header = header + "'" + Chr(13) + "' - Returns: RETURN_DESCRIPTION" + Chr(13)
End If
header = header + "///" + Chr(13)

// Prepend the header comment to the body of the method.
Text = header + Text

// Select the DESCRIPTION field in the comment header.
Dim descriptionStart As Integer = Text.InStr("' DESCRIPTION")
If descriptionStart > 0 Then
SelStart = descriptionStart + 1
SelLength = SelStart + 5
End If
