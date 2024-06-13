Imports System
Imports Tail
Imports Xunit

Namespace ZTailTest
    Public Class AnalysisTest

        <Fact>
        Sub TestCase01()
            Dim args = New String() {"exe", "-n", "10", "-f", "-e", "shift-jis", "test1.txt", "test2.txt"}
            Dim analysis = SwitchAnalyzer.Create().
                SetSwitch("number", "n"c, valueCount:=1, valueName:="line", description:="�o�͂���s�����w�肷��").
                SetSwitch("file", "f"c, description:="�t�@�C���̒ǋL���Ď�����").
                SetSwitch("encode", "e"c, "encode", valueCount:=1, valueName:="encode", description:="�����R�[�h���w�肷��(shift-jis,UTF-8�Ȃ�)").
                Parse(args)
            Assert.True(analysis.HasSwitch("number"))
            Assert.Equal("10", analysis.GetSwitchParameter("number")(0))
            Assert.True(analysis.HasSwitch("file"))
            Assert.True(analysis.HasSwitch("encode"))
            Assert.Equal("shift-jis", analysis.GetSwitchParameter("encode")(0))
        End Sub

        <Fact>
        Sub TestErrorCase01()
            Dim args = New String() {"exe", "-n", "-f", "-e", "shift-jis", "test1.txt", "test2.txt"}
            Assert.Throws(Of ArgumentException)(
                Sub()
                    SwitchAnalyzer.Create().
                        SetSwitch("number", "n"c, valueCount:=1, valueName:="line", description:="�o�͂���s�����w�肷��").
                        SetSwitch("file", "f"c, description:="�t�@�C���̒ǋL���Ď�����").
                        SetSwitch("encode", "e"c, "encode", valueCount:=1, valueName:="encode", description:="�����R�[�h���w�肷��(shift-jis,UTF-8�Ȃ�)").
                        Parse(args)
                End Sub
            )

            Dim args2 = New String() {"exe", "-f", "-e", "shift-jis"}
            Assert.Throws(Of ArgumentException)(
                Sub()
                    SwitchAnalyzer.Create().
                        SetequiredParameter(True).
                        SetSwitch("file", "f"c, description:="�t�@�C���̒ǋL���Ď�����").
                        SetSwitch("encode", "e"c, "encode", valueCount:=1, valueName:="encode", description:="�����R�[�h���w�肷��(shift-jis,UTF-8�Ȃ�)").
                        Parse(args2)
                End Sub
            )
        End Sub

    End Class

End Namespace

