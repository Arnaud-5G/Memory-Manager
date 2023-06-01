' Commands : *All, *AllData, *Linex x=[Int]
Function Read(Path, Var)

    Const nullValue = "NULL" ' represents the NULL return value [String]
    Const newLn = Chr(13) & Chr(10) ' same as vbCrLf [Chr]
    output = ""

    ' create the fso object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    ' access the file
    Set File = FSO.OpenTextFile(Path, 1) ' 1 = reading, 2 = write, 8 = write at the end

    ' get the data of the file then close it
    allData = File.readall
    File.Close

    If Mid(Var, 1) = "*" Then
        ' removes the * to keep only the command
        command = Split(Var, "*", 2)(1)
        ' verify if the command is Line if yes then keep only Line
        If Mid(command, 1, 4) = "Line" Then command = "Line"

        Select Case command
            ' if command all
            Case "All"
                ' get all text in the file
                For Each ln In Split(allData, newLn)
                    output = output & ln & newLn
                Next
            Case "AllData"
                ' get all text except vars in the file
                For Each ln In Split(allData, newLn)
                    output = output & Split(ln, " ", 2)(1) & newLn
                Next
            Case "Line"
                ' get all text in one line
                lnNum = Split(Var, "*Line", 2)(1)
                i = 0
                For Each ln In Split(allData, newLn)
                    i++
                    If i = lnNum Then output = ln
                Next
        End Select
    Else
        ' if no commands
        For Each ln In Split(allData, newLn)
            VarName = Split(ln, " ", 2)(0)
            If VarName = Var Then
                output = Split(ln, " ", 2)(1)
                Exit For
            End If
        Next
    End If

    ' memory management (vbs does not have one for objects only vars have one)
    Set File = Nothing
    Set FSO = Nothing

    ' return
    If output = "" Then Read = nullValue
    Else Then Read = output
End Function

Function Write(Path, Var, Text)
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set File = FSO.OpenTextFile(Path, 1) ' 1 = reading, 2 = write, 8 = write at the end
    PathOver = Split(Path, "\")
    i = -1
    For each x in PathOver
        i = i+1
    Next
        
    PathOverwrite = ""
    y = -1

    For each x in PathOver
        y = y+1
        If Not y = i Then
            PathOverwrite = PathOverwrite & x & "\"
        Else
            PathOverwrite = PathOverwrite & "temp.dat"
        End If
    Next

    Set FileOver = FSO.CreateTextFile(PathOverwrite, 2) ' 1 = reading, 2 = write, 8 = write at the end
        
    Do While Not File.AtEndofStream
        ln = File.ReadLine
        VarName = Split(ln, " ", 2)(0)
        If VarName = Var Then
            FileOver.WriteLine Var & " " & Text
        Else
            FileOver.WriteLine ln
        End If
    Loop

    File.Close
    FileOver.Close

    Set FileOver = FSO.OpenTextFile(PathOverwrite, 1)
    Set File = FSO.OpenTextFile(Path, 2)

    File.Write ""
    text = FileOver.ReadAll
    File.Write text
    File.Close
    FileOver.Close

    FSO.DeleteFile(PathOverwrite)

End Function