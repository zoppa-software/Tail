Option Strict On
Option Explicit On

Namespace ApplicationSwitch

    ''' <summary>コマンドラインスイッチを表します。</summary>
    Public NotInheritable Class CommandLineSwitch

        Private Shared mUniqueCounter As Integer = 0

        Private ReadOnly mUnique As Integer = System.Threading.Interlocked.Increment(mUniqueCounter)

        ''' <summary>スイッチの名前です。</summary>
        Public ReadOnly Property Name As String

        ''' <summary>短いスイッチです。</summary>
        Public Property SwitchChar As Char?

        ''' <summary>長いスイッチです。</summary>
        Public Property SwitchName As String

        ''' <summary>引数の数です（マイナス値なら引数の数は無制限です）</summary>
        Public Property ValueCount As Integer

        ''' <summary>引数の名前です（Helpで表示）</summary>
        Public Property ValueName As String

        ''' <summary>引数が必須ならば真です。</summary>
        Public Property Required As Boolean

        ''' <summary>スイッチの説明です。</summary>
        Public Property Description As String

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="name">スイッチ名。</param>
        Public Sub New(name As String)
            Me.mUnique = System.Threading.Interlocked.Increment(mUniqueCounter)
            Me.Name = name
        End Sub


        Public Overrides Function GetHashCode() As Integer
            Return Me.Name.GetHashCode() Xor Me.mUnique.GetHashCode()
        End Function

    End Class

End Namespace