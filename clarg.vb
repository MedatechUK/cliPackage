Imports System.IO
Imports System.Net.Mail
Imports System.Reflection
Imports System.Xml
Imports System.Xml.Serialization
Imports Newtonsoft.Json

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

        Dim f As Boolean = False
        With Assembly.GetEntryAssembly
            For Each s As String In .GetManifestResourceNames
                If s.Contains("syntax.txt") Then
                    f = True
                    Using sr As New StreamReader(.GetManifestResourceStream(s))
                        Console.Write(sr.ReadToEnd)

                    End Using
                End If
            Next

            If Not f Then
                Console.Write("Help file not found.")

            End If

        End With

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

    Sub Colourise(colour As ConsoleColor, str As String)

        Dim last As ConsoleColor = Console.ForegroundColor
        Console.ForegroundColor = colour
        Console.Write(str)
        Console.ForegroundColor = last

    End Sub

    Sub line(str As String, ParamArray param() As String)

        Console.Write(String.Format(str, param))
        Try
            Console.Write(
                String.Format(
                    " {0} ",
                    New String("."c, Console.WindowWidth - (10 + String.Format(str, param).Length))
                )
            )

        Catch ex As Exception
            Console.WriteLine()
            Try
                Console.Write(
                String.Format(
                    " {0} ",
                    New String("."c, Console.WindowWidth - (10))
                )
            )

            Catch ex2 As Exception

            End Try


        End Try

    End Sub

    Sub line(str As String)

        Console.Write(str)
        Try
            Console.Write(
                String.Format(
                    " {0} ",
                    New String("."c, Console.WindowWidth - (10 + str.Length))
                )
            )

        Catch ex As Exception
            Console.WriteLine()
            Try
                Console.Write(
                String.Format(
                    " {0} ",
                    New String("."c, Console.WindowWidth - (10))
                )
            )

            Catch ex2 As Exception

            End Try

        End Try

    End Sub

    Public Function Attempt(Sender As EventHandler, e As EventArgs, str As String, ParamArray param() As String) As Boolean
        Try
            line(str, param)
            Sender.Invoke(Me, e)

            Colourise(ConsoleColor.Green, "OK")
            Console.WriteLine()
            Return True

        Catch ex As Exception
            Log(ex.Message)
            Colourise(ConsoleColor.Red, "FAIL")
            Console.WriteLine()
            Colourise(ConsoleColor.Red, ex.Message)
            Console.WriteLine()
            Return False

        End Try

    End Function

#Region "Deserialisation"

    Private ret As Object
    Public Function Deserial(Inputfile As FileInfo, DeserialObject As Type, Optional SerialType As eSerialType = eSerialType.unspecified) As Object

        ret = Nothing
        Dim sucsess As Boolean = False

        Select Case SerialType
            Case eSerialType.unspecified
                Select Case Inputfile.Extension.ToLower.Substring(1)
                    Case "xml", "config"
                        sucsess = Attempt(AddressOf xd, New serialArgs(Inputfile, DeserialObject), "Deserialising {0}", Inputfile.FullName)

                    Case "json", "jsn", "jso"
                        sucsess = Attempt(AddressOf jd, New serialArgs(Inputfile, DeserialObject), "Deserialising {0}", Inputfile.FullName)

                    Case Else
                        Throw New Exception(String.Format("Unknown deserial type {0}.", Inputfile.Extension.ToUpper))

                End Select

            Case eSerialType.json
                sucsess = Attempt(AddressOf jd, New serialArgs(Inputfile, DeserialObject), "Deserialising {0}", Inputfile.FullName)

            Case eSerialType.xml
                sucsess = Attempt(AddressOf xd, New serialArgs(Inputfile, DeserialObject), "Deserialising {0}", Inputfile.FullName)

        End Select

        If sucsess Then
            Return ret

        Else
            Throw New Exception("Deserialisation failed.")

        End If

    End Function

    Private Sub xd(sender As Object, e As serialArgs)

        If Not e.File.Exists Then _
            Throw New Exception(String.Format("File {0} not found.", e.File.Name))

        Dim s As New XmlSerializer(e.serialType)
        Using sr As New StreamReader(e.File.FullName)
            ret = s.Deserialize(sr)

        End Using

    End Sub

    Private Sub jd(sender As Object, e As serialArgs)

        If Not e.File.Exists Then _
            Throw New Exception(String.Format("File {0} not found.", e.File.Name))

        Using sr As New StreamReader(e.File.FullName)
            ret = JsonConvert.DeserializeObject(sr.ReadToEnd, e.serialType)

        End Using

    End Sub

#End Region

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

    Public Function DatedFolder(root As DirectoryInfo) As DirectoryInfo
        Dim ret As DirectoryInfo
        With root
            If Not .Exists Then .Create()
            ret = New DirectoryInfo(Path.Combine(.FullName, Now.ToString("yyyy-MM")))
            With ret
                If Not .Exists Then .Create()
                Return ret
            End With
        End With

    End Function

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
        Using log As New StreamWriter(currentlog.FullName, True)
            log.WriteLine("{0}> {1}", Format(Now, "HH:mm:ss"), String.Format(str, args))
        End Using

    End Sub

    Public Sub Log(ByVal str)
        Using log As New StreamWriter(currentlog.FullName, True)
            log.WriteLine("{0}> {1}", Format(Now, "HH:mm:ss"), str)
        End Using

    End Sub
#End Region

#Region "SMTP Notify Errors"

    Public Sub errNotify(Subject As String, fn As FileInfo, ex As Exception)
        Attempt(AddressOf Notify, New notifyArgs(Subject, fn, ex), "Sending error report")

    End Sub

    Private Sub Notify(sender As Object, e As notifyArgs)

        If Not e.fn.Exists Then Throw New Exception(String.Format("Config file {0} not found.", e.fn.FullName))

        Dim doc As New XmlDocument
        Try
            doc.Load(e.fn.FullName)

        Catch ex As Exception
            Throw New Exception(String.Format("Load {0} failed. {1}", e.fn.FullName, ex.Message))

        End Try

        Dim node As XmlNode = doc.SelectSingleNode("//notifyerror")
        If node Is Nothing Then Throw New Exception(String.Format("No notifyerror node in {0}.", e.fn.FullName))

        Dim add As New List(Of String)
        For Each email As XmlNode In node.SelectNodes("notify")
            add.Add(email.Attributes("address").Value)
        Next

        If add.Count = 0 Then Throw New Exception(String.Format("No email address in {0}.", e.fn.FullName))

        Dim erMail = New MailMessage(node.Attributes("from").Value, add(0))
        With erMail
            With .CC
                For i As Integer = 1 To add.Count - 1
                    .Add(add(i))
                Next
            End With

            .Subject = e.Subject
            .Body = Now.ToString & vbCrLf & e.ex.Message & vbCrLf & vbCrLf & e.ex.StackTrace

            If node.Attributes("smtp") Is Nothing Then Throw New Exception(String.Format("No SMTP set in {0}.", e.fn.FullName))
            Using c As New SmtpClient(node.Attributes("smtp").Value)
                c.Send(erMail)

            End Using

        End With

    End Sub

#End Region

End Class

Public Class notifyArgs : Inherits EventArgs
    Public fn As FileInfo
    Public ex As Exception
    Public Subject As String

    Sub New(SubjectStr As String, file As FileInfo, exep As Exception)
        Subject = SubjectStr
        fn = file
        ex = exep

    End Sub

End Class