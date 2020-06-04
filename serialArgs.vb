Imports System.IO
Public Enum eSerialType
    unspecified = 0
    xml = 1
    json = 2

End Enum

Public Class serialArgs : Inherits EventArgs

    Public File As FileInfo
    Public serialType As Type

    Sub New(F As FileInfo, st As Type)
        File = F
        serialType = st

    End Sub

End Class
