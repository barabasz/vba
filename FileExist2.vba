Function FileExist2(ByVal sFile As String) As Boolean
    If (Len(Dir(sFile)) > 0) And (Len(sFile) > 0) Then
        FileExist2 = True
    Else
        FileExist2 = False
    End If
End Function
