Option Strict On
Option Explicit On
Imports Tail.SwitchAnalyzer

''' <summary>コマンドラインスイッチを表します。</summary>
Public NotInheritable Class CommandLineSwitch

    Public ReadOnly Property Name As String

    Public Property SwitchChar As Char?

    Public Property SwitchName As String

    Public Property ValueCount As Integer

    Public Property ValueName As String

    Public Property Required As Boolean

    Public Property Description As String

    Public Sub New(name As String)
        Me.Name = name
    End Sub

End Class