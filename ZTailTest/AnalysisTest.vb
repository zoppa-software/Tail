Imports System
Imports Tail
Imports Tail.ApplicationSwitch
Imports Xunit

Namespace ZTailTest

    Public Class AnalysisTest

        <Fact>
        Sub TestCase01()
            Dim args = New String() {"exe", "-n", "10", "-f", "-e", "shift-jis", "test1.txt", "test2.txt"}
            Dim analysis = SwitchAnalyzer.Create().
                SetSwitch("number", "n"c, valueCount:=1, valueName:="line", description:="出力する行数を指定する").
                SetSwitch("file", "f"c, description:="ファイルの追記を監視する").
                SetSwitch("encode", "e"c, "encode", valueCount:=1, valueName:="encode", description:="文字コードを指定する(shift-jis,UTF-8など)").
                Parse(args)
            Assert.True(analysis.HasSwitch("number"))
            Assert.Equal("10", analysis.GetSwitchParameter("number")(0))
            Assert.True(analysis.HasSwitch("file"))
            Assert.True(analysis.HasSwitch("encode"))
            Assert.Equal("shift-jis", analysis.GetSwitchParameter("encode")(0))

            Assert.True(analysis.HasParameter)
            Assert.Equal(2, analysis.GetParameters().Length)
            Assert.Equal("test1.txt", analysis.GetParameters()(0))
            Assert.Equal("test2.txt", analysis.GetParameters()(1))


            Dim args2 = New String() {"exe"}
            Dim analysis2 = SwitchAnalyzer.Create().
                Parse(args2)
            Assert.True(analysis2.IsEmpty)
        End Sub

        <Fact>
        Sub TestErrorCase01()
            Dim args = New String() {"exe", "-n", "-f", "-e", "shift-jis", "test1.txt", "test2.txt"}
            Assert.Throws(Of ArgumentException)(
                Sub()
                    SwitchAnalyzer.Create().
                        SetSwitch("number", "n"c, valueCount:=1, valueName:="line", description:="出力する行数を指定する").
                        SetSwitch("file", "f"c, description:="ファイルの追記を監視する").
                        SetSwitch("encode", "e"c, "encode", valueCount:=1, valueName:="encode", description:="文字コードを指定する(shift-jis,UTF-8など)").
                        Parse(args, True)
                End Sub
            )

            Dim args2 = New String() {"exe", "-f", "-e", "shift-jis"}
            Assert.Throws(Of ArgumentException)(
                Sub()
                    SwitchAnalyzer.Create().
                        SetequiredParameter(True).
                        SetSwitch("file", "f"c, description:="ファイルの追記を監視する").
                        SetSwitch("encode", "e"c, "encode", valueCount:=1, valueName:="encode", description:="文字コードを指定する(shift-jis,UTF-8など)").
                        Parse(args2, True)
                End Sub
            )
        End Sub

    End Class

End Namespace

