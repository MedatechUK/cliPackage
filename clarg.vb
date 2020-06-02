Imports System.IO

Public Class clArg
    Inherits Dictionary(Of String, String)

    Private Enum eMode
        Switch
        Param
    End Enum

    Sub New()

        Dim Args = Environment.GetCommandLineArgs()
        Console.WriteLine("")

        Dim i As Integer = 1
        Dim m As eMode = eMode.Switch
        Dim thisSwitch As String = ""

        If Args.Length = i Then Exit Sub

        Do
            Select Case Args(i).Substring(0, 1)
                Case "-", "/"
                    Add(Args(i).Substring(1).ToLower, "")
                    thisSwitch = Args(i).Substring(1).ToLower
                    m = eMode.Param
                Case Else
                    Select Case m
                        Case eMode.Param
                            Me(thisSwitch.ToLower) = Args(i)
                            thisSwitch = ""
                            m = eMode.Switch

                        Case eMode.Switch
                            Add(Args(i).ToLower, "")

                    End Select

            End Select
            i += 1
        Loop Until i = Args.Count

    End Sub

    Public Sub syntax()
        Dim Syntaxfile As New FileInfo(Path.Combine(Path.GetDirectoryName(Reflection.Assembly.GetExecutingAssembly.Location), "syntax.txt"))
        If Syntaxfile.Exists Then
            Using sr As New StreamReader(Syntaxfile.FullName)
                Console.Write(sr.ReadToEnd)
            End Using

        End If
        Console.WriteLine("")

    End Sub

    Sub wait()
        If Keys.Contains("w") Then
            Console.WriteLine("")
            Console.Write("Finished. Press any key.")
            Console.CursorVisible = False
            Console.ReadKey()
            Console.WriteLine("")
            Console.CursorVisible = True

        End If

    End Sub

    Sub Colourise(colour As ConsoleColor, str As String, ParamArray param() As String)

        Dim last As ConsoleColor = Console.ForegroundColor
        Console.ForegroundColor = colour
        Console.Write(
                String.Format(
                str,
                param
            )
        )
        Console.ForegroundColor = last

    End Sub

    Sub line(str As String, ParamArray param() As String)
        Console.Write(String.Format(str, param))

        Console.Write(
            String.Format(
                " {0} ",
                New String("."c, Console.WindowWidth - (10 + String.Format(str, param).Length))
            )
        )
    End Sub

#Region "Progress bar"

    Private p As Integer = -1
    Private x As Integer
    Private y As Integer
    Private _max As Integer
    Private _current As Integer

    Public Property Progress As Integer
        Get
            Return _current
        End Get
        Set(value As Integer)
            _current = value
            If _max > 0 Then
                If CInt(_current / _max * 100) > p Then
                    p = CInt(_current / _max * 100)
                    Console.CursorLeft = x
                    Console.CursorTop = y
                    Colourise(ConsoleColor.Yellow, "{0}%", p.ToString)

                End If
                If _current >= _max Then
                    Console.CursorLeft = x
                    Console.CursorTop = y
                    Colourise(ConsoleColor.Green, "Done.")
                    Console.CursorVisible = True

                End If
            End If
        End Set
    End Property

    Public Sub StartProgress(max As Integer)
        _max = max
        Console.CursorVisible = False
        x = Console.CursorLeft
        y = Console.CursorTop
        _current = 0

    End Sub

#End Region

#Region "Console yes no"
    Public Function YesNo(format As String, ParamArray args() As String) As Boolean
        Console.Write(format, args)
        Return _YesNo()
    End Function

    Public Function YesNo(format As String) As Boolean
        Console.Write(format)
        Return _YesNo()
    End Function

    Private Function _YesNo() As Boolean
        Console.Write(" > ")
        Dim k As ConsoleKeyInfo
        Do Until Console.KeyAvailable
            k = Console.ReadKey
            Exit Do
        Loop
        Console.WriteLine("")
        Return k.Key = ConsoleKey.Y

    End Function

#End Region

#Region "Logging"

    Private Function LogFolder() As DirectoryInfo
        Return New DirectoryInfo(
            Path.Combine(
               Directory.GetCurrentDirectory,
                String.Format(
                    "log\{0}",
                    Now.ToString("yyyy-MM")
                )
            )
        )

    End Function

    Private Function currentlog() As FileInfo
        With LogFolder()
            If Not .Exists Then .Create()
            Return New FileInfo(
                Path.Combine(
                    .FullName,
                    String.Format(
                        "{0}.txt",
                        Now.ToString("yyMMdd")
                    )
                )
            )

        End With

    End Function

    Public Sub Log(ByVal str, ByVal ParamArray args())
        Console.WriteLine(str, args)
        Using log As New StreamWriter(currentlog.FullName, True)
            log.WriteLine("{0}> {1}", Format(Now, "HH:mm:ss"), String.Format(str, args))
        End Using

    End Sub

#End Region

End Class
