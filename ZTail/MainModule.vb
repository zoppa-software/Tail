Option Strict On
Option Explicit On
Imports System.IO
Imports System.Text

''' <summary>アプリケーションのエントリーポイントを提供します。</summary>
Public Module MainModule

    Public Const NUMBER_SW As String = "n"

    Public Const FILE_SW As String = "f"

    Public Const ENCODE_SW As String = "encode"

    Public Sub Main()
        Dim analysis = SwitchAnalyzer.Create().
            SetDescription("シンプルな Tailコマンド。").
            SetRemark("ファイルの末尾を表示します。").
            SetAuthor("zoppa software").
            SetMailAddress("zoppa@ab.auone-net.jp").
            SetValueName("file path").
            SetSwitch(NUMBER_SW, "n"c, valueCount:=1, valueName:="line", description:="出力する行数を指定する").
            SetSwitch(FILE_SW, "f"c, description:="ファイルの追記を監視する").
            SetSwitch(ENCODE_SW, "e"c, "encode", valueCount:=1, valueName:="encode", description:="文字コードを指定する(shift-jis,UTF-8など)").
            Parse()

        ' 出力行数を読み込む
        Dim lineCount As Integer = 10
        If analysis.HasSwitch(NUMBER_SW) Then
            lineCount = analysis.GetSwitchParameter(Of Integer)(NUMBER_SW, 0)
        End If

        ' エンコード指定を読み込む
        Dim encode As Text.Encoding = Text.Encoding.Default
        If analysis.HasSwitch(ENCODE_SW) Then
            encode = GetEncoder(analysis.GetSwitchParameter(ENCODE_SW)(0))
        End If

        ' ファイルの追加を監視する
        Dim watch As Boolean = analysis.HasSwitch(FILE_SW)
        If watch AndAlso analysis.GetParameters().Length > 1 Then
            Throw New ArgumentException("ファイルの追加を監視する場合、ファイルは1つしか指定できません。")
        End If

        ' ファイル数ループ
        For Each fPath In analysis.GetParameters()
            If Not IO.File.Exists(fPath) Then
                WriteErrorLine($"'{fPath}'は存在しないファイルです")
                System.Environment.Exit(1)
            End If

            Dim file As New FileInfo(fPath)
            If file.Exists Then
                If analysis.GetParameters().Length > 1 Then
                    WriteLine("===== {file.FullName} =====")
                End If

                ' 最初の行出力を行う
                Dim lastPoint = FirstTail(file.FullName, encode, lineCount)

                ' 追加を監視する
                If watch Then
                    Do While True
                        file.Refresh()
                        If file.Exists Then
                            Dim line = NextTail(file.FullName, encode, lastPoint)
                            If line <> "" Then
                                Write(line)
                            End If
                        End If
                        Threading.Thread.Sleep(300)
                    Loop
                End If
            End If
        Next
    End Sub

    ''' <summary>最初の行出力を行う。</summary>
    ''' <param name="path">対象ファイル。</param>
    Private Function FirstTail(ByVal path As String, enc As Text.Encoding, lineCount As Integer) As Long
        Dim lastPoint As Long = 0
        Try
            Dim lines As New Queue(Of String)()
            Using fs = New FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                ' 最初のファイルサイズを保持する
                lastPoint = fs.Length

                ' シークした位置以降の文字列を取得
                Using sr As New StreamReader(fs, enc)
                    Do While sr.Peek() <> -1
                        Dim ln = sr.ReadLine()

                        lines.Enqueue(ln)
                        If lines.Count > lineCount Then
                            lines.Dequeue()
                        End If
                    Loop
                End Using

                ' 読み込んだ文字列を出力
                For Each ln In lines
                    WriteLine(ln)
                Next
            End Using
        Catch ex As Exception
            ' 空実装
        End Try
        Return lastPoint
    End Function

    Private Function NextTail(ByVal path As String, enc As Text.Encoding, ByRef lastPoint As Long) As String
        Dim res As String = ""
        Try
            Using fs = New FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                ' ファイルサイズを確認
                If fs.Length < lastPoint Then
                    lastPoint = 0
                End If

                ' 前回のファイルサイズにシーク
                fs.Seek(lastPoint, SeekOrigin.Begin)
                lastPoint = fs.Length

                ' シークした位置以降の文字列を取得
                Using sr As New StreamReader(fs, enc)
                    res = sr.ReadToEnd()
                End Using
            End Using
        Catch ex As Exception
            res = ""
        End Try
        Return res
    End Function

End Module
