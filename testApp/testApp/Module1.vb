Imports MedatechUK.CLI

Module Module1

    Public args As New clArg

    Sub Main()

        With args
            .Colourise(ConsoleColor.Blue, "Hello {0}!", "world")
            .line("Waiting for {0}", "stuff")
            .StartProgress(100)
            For i As Integer = 0 To 100
                .Progress = i
                Threading.Thread.Sleep(10)

            Next
            .wait()

        End With

    End Sub

End Module
