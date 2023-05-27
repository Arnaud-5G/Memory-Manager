Function Read(Path, Var)
    ' If Mid(Var, 1) = "*" Then
    '     command = Split(Var, "*", 2)(1)
    '     If Mid(command, 1, 4) = "Line" Then command = "Line"

    '     Select Case command
    '         Case "All"

    '         Case "AllData"
                
    '         Case "Line"
            
    '     End Select
    ' End If


    If Var = "*All" Then
        Set FSO = CreateObject("Scripting.FileSystemObject")
        Set File = FSO.OpenTextFile(Path, 1) ' 1 = reading, 2 = write, 8 = write at the end

        allData = File.readall
        File.Close
        For Each ln In Split(allData, vbCrLf)
            output = output & ln & vbCrLf
        Next

        ' memory management (vbs does not have one for objects only vars have one)
        Set File = Nothing
        Set FSO = Nothing

        ' return
        Read = output
    ElseIf Var = "All" Then
        Set FSO = CreateObject("Scripting.FileSystemObject")
        Set File = FSO.OpenTextFile(Path, 1) ' 1 = reading, 2 = write, 8 = write at the end

        allData = File.readall
        File.Close
        For Each ln In Split(allData, vbCrLf)
            output = output & Split(ln, " ", 2)(1) & vbCrLf
        Next

        ' memory management (vbs does not have one for objects only vars have one)
        Set File = Nothing
        Set FSO = Nothing

        ' return
        Read = output
    Else
        Set FSO = CreateObject("Scripting.FileSystemObject")
        Set File = FSO.OpenTextFile(Path, 1) ' 1 = reading, 2 = write, 8 = write at the end

        allData = File.readall
        File.Close
        For Each ln In Split(allData, vbCrLf)
            VarName = Split(ln, " ", 2)(0)
            If VarName = Var Then
                output = Split(ln, " ", 2)(1)
                Exit For
            End If
        Next

        ' memory management (vbs does not have one for objects only vars have one)
        Set File = Nothing
        Set FSO = Nothing

        ' return
        Read = output
    End If
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

Write "C:\Users\20200791\Desktop\test.dat", "Test", "Hello"