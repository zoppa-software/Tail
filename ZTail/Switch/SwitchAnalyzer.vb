Option Strict On
Option Explicit On

Imports System.Reflection
Imports System.Runtime.CompilerServices
Imports Tail.IOHelper

Namespace ApplicationSwitch

    Public NotInheritable Class SwitchAnalyzer

        Private Const NEST_STR = "    "

        Private Const HELP_SW As String = "help"

        Private Const VERSION_SW As String = "version"

        Private mAsmTitle As String = ""

        Private mDescription As String = ""

        Private mVersion As New Version(0, 1, 0, 1)

        Private mCopyright As String = ""

        Private mAuthor As String = ""

        Private mMailAddress As String = ""

        Private mRequired As Boolean = False

        Private mValueName As String = "Value"

        Private mRemark As String = ""

        Private mSwitchList As New List(Of CommandLineSwitch)

        Private mSwitchDic As New SortedList(Of String, CommandLineSwitch)

        Public Shared Function Create() As SwitchAnalyzer
            Dim res As New SwitchAnalyzer()
            With [Assembly].GetExecutingAssembly()
                res.mAsmTitle = .GetCustomAttribute(Of AssemblyTitleAttribute)().Title
                res.mDescription = If(.GetCustomAttribute(Of AssemblyDescriptionAttribute)()?.Description, "")
                res.mVersion = .GetName().Version
                res.mCopyright = If(.GetCustomAttribute(Of AssemblyCopyrightAttribute)()?.Copyright, "")
                res.mAuthor = ""
                res.mMailAddress = ""
            End With

            res.SetSwitch(HELP_SW, "h"c, "help", description:="ヘルプを表示します。")
            res.SetSwitch(VERSION_SW, "v"c, "version", description:="バージョンを表示します。")

            Return res
        End Function

        Public Overrides Function ToString() As String
            Dim res As New Text.StringBuilder()
            res.AppendLine("Name")
            res.AppendFormat("{0}{1} {2}", NEST_STR, Me.mAsmTitle, If(Me.mDescription.Length > 0, "- " & Me.mDescription, "")).AppendLine()
            res.AppendFormat("{0}{1} {2}", NEST_STR, Me.mVersion, Me.mCopyright).AppendLine()
            If Me.mAuthor.Length > 0 Then
                res.AppendLine("Author")
                res.AppendFormat("{0}{1} {2}", NEST_STR, Me.mAuthor, Me.mMailAddress).AppendLine()
            End If
            res.AppendLine("Usage")
            res.AppendFormat("{0}{1} [option] {2}",
                         NEST_STR,
                         Me.mAsmTitle,
                         If(Me.mRequired, $"[{Me.mValueName}]", "")).AppendLine()
            If Me.mSwitchList.Count > 0 Then
                res.AppendLine("Option")

                Dim msg As New List(Of String)()
                For Each sw In Me.mSwitchList
                    msg.Add(String.Format("{0}{1}{2}{3}",
                    If(sw.SwitchChar.HasValue, $"-{sw.SwitchChar} ", ""),
                    If(sw.SwitchName.Length > 0, $"--{sw.SwitchName} ", ""),
                    If(sw.ValueCount > 0, $"[{sw.ValueName}{If(sw.ValueCount > 1, " ...", "")}] ", ""),
                    If(sw.Required, "(!required) ", "")
                ))
                Next

                Dim nlen = msg.Max(Function(m) m.Length)

                For i As Integer = 0 To Me.mSwitchList.Count - 1
                    Dim sw = Me.mSwitchList(i)
                    res.AppendFormat("{0}{1,-" & nlen + 2 & "} {2}",
                                 NEST_STR,
                                 msg(i),
                                 sw.Description).AppendLine()
                Next
            End If
            If Me.mRemark.Length > 0 Then
                res.AppendLine("Remark")
                res.AppendLine(" " & Me.mRemark)
            End If
            Return res.ToString()
        End Function

        Public Function SetTitle(asmTitle As String) As SwitchAnalyzer
            Me.mAsmTitle = asmTitle
            Return Me
        End Function

        Public Function SetDescription(description As String) As SwitchAnalyzer
            Me.mDescription = description
            Return Me
        End Function

        Public Function SetVersion(version As Version) As SwitchAnalyzer
            Me.mVersion = version
            Return Me
        End Function

        Public Function SetVersion(major As Integer,
                               Optional minor As Integer = 1,
                               Optional build As Integer = 0,
                               Optional revision As Integer = 0) As SwitchAnalyzer
            Me.mVersion = New Version(major, minor, build, revision)
            Return Me
        End Function

        Public Function SetCopyright(copyright As String) As SwitchAnalyzer
            Me.mCopyright = copyright
            Return Me
        End Function

        Public Function SetAuthor(author As String) As SwitchAnalyzer
            Me.mAuthor = author
            Return Me
        End Function

        Public Function SetMailAddress(mailAddress As String) As SwitchAnalyzer
            Me.mMailAddress = mailAddress
            Return Me
        End Function

        Public Function SetequiredParameter(required As Boolean) As SwitchAnalyzer
            Me.mRequired = required
            Return Me
        End Function

        Public Function SetValueName(valueName As String) As SwitchAnalyzer
            Me.mRequired = True
            Me.mValueName = valueName
            Return Me
        End Function

        Public Function SetRemark(remark As String) As SwitchAnalyzer
            Me.mRemark = remark
            Return Me
        End Function

        Public Function SetSwitches(ParamArray switches() As CommandLineSwitch) As SwitchAnalyzer
            For Each sw In switches
                If Me.mSwitchDic.ContainsKey(sw.Name) Then
                    Me.mSwitchList.Remove(Me.mSwitchDic(sw.Name))
                    Me.mSwitchList.Insert(Me.mSwitchList.Count - 2, sw)
                    Me.mSwitchDic(sw.Name) = sw
                Else
                    If Me.mSwitchList.Count >= 2 Then
                        Me.mSwitchList.Insert(Me.mSwitchList.Count - 2, sw)
                    Else
                        Me.mSwitchList.Add(sw)
                    End If
                    Me.mSwitchDic.Add(sw.Name, sw)
                End If
            Next
            Return Me
        End Function

        Public Function SetSwitch(name As String,
                              Optional switchChar As Char? = Nothing,
                              Optional switchName As String = "",
                              Optional valueName As String = "",
                              Optional valueCount As Integer = 0,
                              Optional required As Boolean = False,
                              Optional description As String = "") As SwitchAnalyzer
            Return Me.SetSwitches(
            New CommandLineSwitch(name) With {
                .SwitchChar = switchChar,
                .SwitchName = switchName,
                .ValueName = valueName,
                .ValueCount = valueCount,
                .Required = required,
                .Description = description
            }
        )
        End Function

        Public Function Parse(Optional isThrowableException As Boolean = False) As Result
            Return Me.Parse(System.Environment.GetCommandLineArgs(), isThrowableException)
        End Function

        Public Function Parse(args() As String, Optional isThrowableException As Boolean = False) As Result
            Dim hasHOrV As Boolean = False
            Try
                Dim sws As New Dictionary(Of CommandLineSwitch, List(Of String))()
                Dim prms As New List(Of String)()

                Dim i As Integer = 1
                Do While i < args.Length
                    ' スイッチを解析する
                    Dim stp As Integer = 0
                    For Each sw In Me.mSwitchDic.Values
                        Dim cnt = Me.Analysis(sw, args, i, sws)
                        stp = If(stp > cnt, stp, cnt)
                    Next

                    ' 有効なスイッチなら次へ、そうで無ければパラメータとして保持する
                    If stp > 0 Then
                        i += stp
                    Else
                        For j = i To args.Length - 1
                            prms.Add(args(j).Trim())
                        Next
                        Exit Do
                    End If
                Loop


                Dim res As New Result(sws, prms)

                If res.HasSwitch(HELP_SW) Then
                    WriteLine(Me.ToString())
                    hasHOrV = True
                End If

                If res.HasSwitch(VERSION_SW) Then
                    WriteLine(Me.mVersion.ToString())
                    hasHOrV = True
                End If

                If Me.mRequired AndAlso Not res.HasParameter() Then
                    Throw New ArgumentException($"パラメータが必要です。")
                End If

                Return res

            Catch ex As Exception
                WriteErrorLine(ex.Message)
                If isThrowableException Then
                    Throw
                Else
                    If Not hasHOrV Then
                        WriteLine(Me.ToString())
                        WriteLine()
                    End If
                    Return New Result(True)
                End If
            End Try
        End Function

        Private Function Analysis(sw As CommandLineSwitch, args() As String, index As Integer, sws As Dictionary(Of CommandLineSwitch, List(Of String))) As Integer
            Dim p = args(index).Trim()
            If p.StartsWith("--") Then
                If sw.SwitchName.Length > 0 AndAlso p = $"--{sw.SwitchName}" Then
                    Return Me.CountParameter(sw, args, index, sws)
                End If
            ElseIf p.StartsWith("-") Then
                If sw.SwitchChar.HasValue AndAlso p.IndexOf(sw.SwitchChar.Value) > 0 Then
                    Return Me.CountParameter(sw, args, index, sws)
                End If
            End If
            Return 0
        End Function

        Private Function CountParameter(sw As CommandLineSwitch, args() As String, index As Integer, sws As Dictionary(Of CommandLineSwitch, List(Of String))) As Integer
            If sws.ContainsKey(sw) Then
                Throw New ArgumentException($"Switch '{sw.Name}' is already defined.")
            End If

            Dim needCnt As Integer = 0
            For i As Integer = index + 1 To Math.Min(args.Length - 1, index + sw.ValueCount)
                If args(i).StartsWith("-") Then
                    Exit For
                Else
                    needCnt += 1
                End If
            Next

            If needCnt < sw.ValueCount Then
                Throw New ArgumentException($"'{sw.Name}'スイッチは{sw.ValueCount}個のパラメータが必要です。")
            End If

            Dim params As New List(Of String)()
            For i As Integer = index + 1 To index + needCnt
                params.Add(args(i).Trim())
            Next

            sws.Add(sw, params)

            Return (needCnt + 1)
        End Function

        Public NotInheritable Class Result

            Private ReadOnly mSwitches As Dictionary(Of String, String())

            Private ReadOnly mParameters As String()

            Public ReadOnly Property IsHelp() As Boolean
                Get
                    Return Me.HasSwitch(HELP_SW)
                End Get
            End Property

            Public ReadOnly Property IsVersion() As Boolean
                Get
                    Return Me.HasSwitch(VERSION_SW)
                End Get
            End Property

            Public ReadOnly Property IsEmpty() As Boolean
                Get
                    Return (Me.mSwitches.Count + Me.mParameters.Length <= 0)
                End Get
            End Property

            Public ReadOnly Property IsError As Boolean

            Public Sub New(isErr As Boolean)
                Me.mSwitches = New Dictionary(Of String, String())()
                Me.mParameters = New String() {}
                Me.IsError = isErr
            End Sub

            Public Sub New(switches As Dictionary(Of CommandLineSwitch, List(Of String)), parameters As List(Of String))
                Me.mSwitches = New Dictionary(Of String, String())()
                For Each kv In switches
                    Me.mSwitches.Add(kv.Key.Name, kv.Value.ToArray())
                Next
                Me.mParameters = parameters.ToArray()
                Me.IsError = False
            End Sub

            Public Function HasSwitch(sw As String) As Boolean
                Return Me.mSwitches.ContainsKey(sw)
            End Function

            Public Function GetSwitchParameter(sw As String) As String()
                Dim res As String() = Nothing
                If Me.mSwitches.TryGetValue(sw, res) Then
                    Return res
                Else
                    Return New String() {}
                End If
            End Function

            Public Function GetSwitchParameter(Of T)(sw As String, index As Integer) As T
                Dim v = Me.GetSwitchParameter(sw)(index)
                Select Case GetType(T)
                    Case GetType(Integer)
                        Return CType(CObj(Convert.ToInt32(v)), T)
                    Case GetType(Double)
                        Return CType(CObj(Convert.ToDouble(v)), T)
                    Case GetType(String)
                        Return CType(CObj(v), T)
                    Case Else
                        Return CType(CObj(v), T)
                End Select
            End Function

            Public Function GetParameters() As String()
                Return Me.mParameters
            End Function

            Public Function HasParameter() As Boolean
                Return Me.mParameters.Length > 0
            End Function

        End Class

    End Class

    ''' <summary>拡張メソッド群。</summary>
    Module SwitchAnalyzerHelper

        ''' <summary>指定したアセンブリから指定の属性を取得します。</summary>
        ''' <typeparam name="T">属性の型。</typeparam>
        ''' <param name="asm">アセンブリ。</param>
        ''' <returns>属性。</returns>
        <Extension()>
        Public Function GetCustomAttribute(Of T As Attribute)(asm As Assembly) As T
            Return CType(Attribute.GetCustomAttribute(asm, GetType(T)), T)
        End Function

        ''' <summary>指定文字列のエンコーダーを取得する。</summary>
        ''' <param name="eparam">コードページ、またはエンコード名。</param>
        ''' <returns>エンコーダーオブジェクト。</returns>
        Public Function GetEncoder(eparam As String) As Text.Encoding
            Dim eno As Integer = 0
            If Integer.TryParse(eparam, eno) Then
                Return Text.Encoding.GetEncoding(eno)
            Else
                Return Text.Encoding.GetEncoding(eparam)
            End If
        End Function

    End Module

End Namespace
