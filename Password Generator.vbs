passwordLength = InputBox("Enter desired password length (up to 50):", "Password Generator")
If passwordLength = "" Then
    WScript.Quit
End If
possibleChars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890!@#$%^&*()"
password = ""
For i = 1 To passwordLength
    randomIndex = Int(Len(possibleChars) * Rnd())
    password = password & Mid(possibleChars, randomIndex + 1, 1)
Next
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.CreateTextFile("New Password.txt")
objFile.Write password
objFile.Close
MsgBox "New password saved to New Password.txt", vbInformation, "Password Generator"