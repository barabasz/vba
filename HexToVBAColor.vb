Public Function HexToVBAColor(hexColor As String) As Long
    Dim red As String
    Dim green As String
    Dim blue As String
    If Left(hexColor, 1) = "#" Then hexColor = Right(hexColor, 6)
    red = Left(hexColor, 2)
    green = Mid(hexColor, 3, 2)
    blue = Right(hexColor, 2)
    HexToVBAColor = Val("&H" & blue & green & red)
End Function
