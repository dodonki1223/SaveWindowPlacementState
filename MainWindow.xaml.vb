Imports System.ComponentModel
Imports System.Windows.Interop
Imports System.Runtime.InteropServices

#Region "構造体"

'------------------------------------------------★構造体メモ★------------------------------------------------'
'   StructLayoutについて                                                                                       '
'     名前空間: System.Runtime.InteropServices                                                                 '
'     アセンブリ: mscorlib                                                                                     '
'     継承・インターフェイス: Attribute, _Attribute                                                            '
'     StructLayout属性は、メモリ上でのフィールド(メンバ変数)の配置方法を指定するための属性です                 '
'   LayoutKind.Sequentialについて                                                                              '
'     ランタイムによる自動的な並べ替えを行わず、コード上で記述されている順序のままフィールドを配置する         '
'   ※マネージ環境で用いられる構造体はアンマネージ環境のものと互換性がありません。                             '
'     相互運用では構造体の定義にStructLayoutアトリビュートを付加することでアンマネージ互換の構造体を定義します '
'--------------------------------------------------------------------------------------------------------------'

''' <summary>POINT構造体</summary>
''' <remarks>
'''     2 次元空間における、x 座標と y 座標の組を表します
''' </remarks>
<Serializable(), StructLayout(LayoutKind.Sequential)>
Public Structure POINT

    ''' <summary>X座標</summary>
    ''' <remarks></remarks>
    Public X As Integer

    ''' <summary>Y座標</summary>
    ''' <remarks></remarks>
    Public Y As Integer

    ''' <summary>コンストラクタ</summary>
    ''' <param name="pX">X座標</param>
    ''' <param name="pY">Y座標</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal pX As Integer, ByVal pY As Integer)

        Me.X = X
        Me.Y = Y

    End Sub

End Structure

''' <summary>RECT構造体</summary>
''' <remarks>
'''     四角形の左上隅および右下隅の座標を定義します
''' </remarks>
<Serializable(), StructLayout(LayoutKind.Sequential)>
Public Structure RECT

    ''' <summary>Left</summary>
    ''' <remarks>四角形の左上隅のＸ座標を指定します</remarks>
    Public Left As Integer

    ''' <summary>Top</summary>
    ''' <remarks>四角形の左上隅のＹ座標を指定します</remarks>
    Public Top As Integer

    ''' <summary>Right</summary>
    ''' <remarks>四角形の右下隅のＸ座標を指定します</remarks>
    Public Right As Integer

    ''' <summary>Bottom</summary>
    ''' <remarks>四角形の右下隅のＹ座標を指定します</remarks>
    Public Bottom As Integer

    ''' <summary>コンストラクタ</summary>
    ''' <param name="pLeft">Left</param>
    ''' <param name="pTop">Top</param>
    ''' <param name="pRight">Right</param>
    ''' <param name="pBottom">Bottom</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal pLeft As Integer, ByVal pTop As Integer, ByVal pRight As Integer, ByVal pBottom As Integer)

        Me.Left = pLeft
        Me.Top = pTop
        Me.Right = pRight
        Me.Bottom = pBottom

    End Sub

End Structure

''' <summary>WINDOWPLACEMENT構造体</summary>
''' <remarks>
'''     画面上のウィンドウの配置についての情報が含まれています
''' </remarks>
<Serializable(), StructLayout(LayoutKind.Sequential)>
Public Structure WINDOWPLACEMENT

    ''' <summary>length</summary>
    ''' <remarks>
    '''   構造体サイズをバイト単位で指定します。
    ''' </remarks>
    Public length As Integer

    ''' <summary>flags</summary>
    ''' <remarks>
    '''   最小化されたウィンドウとウィンドウを復元する方法の位置を制御するフラグを指定します。
    '''   このメンバーは、次のフラグの一方または両方を指定できます。
    '''     ・WPF_SETMINPOSITION最小化されたウィンドウの x 位置と y 位置を指定できることを示すします。
    '''       このフラグである必要があります指定された座標で設定されているかどうか、 ptMinPositionメンバーです。
    '''     ・WPF_RESTORETOMAXIMIZED復元されたウィンドウが最大化される、最小化する前に、
    '''       最大化されていたかどうかに関係なくを指定します。 この設定は、ウィンドウを閉じたときにのみ有効です。 
    '''       既定の復元動作は変わりません。 このフラグは有効な場合にのみ、このメンバーは値が指定されて、 showCmdメンバーです。
    ''' </remarks>
    Public flags As Integer

    ''' <summary>showCmd</summary>
    ''' <remarks>
    '''   ウィンドウの現在の表示状態を指定します。 このメンバーは、次の値のいずれかになります。
    '''    ・SW_HIDEウィンドウを非表示にし、別のウィンドウをアクティブ化を渡します。
    '''    ・SW_MINIMIZE指定されたウィンドウを最小化し、システムの一覧の最上位ウィンドウを表示します。
    '''    ・SW_RESTOREにアクティブとウィンドウが表示されます。 
    '''      ウィンドウが最小化または最大化されている場合は、Windows が復元され、元のサイズと位置 (同じSW_SHOWNORMAL)。
    '''    ・SW_SHOWウィンドウをアクティブにし、現在のサイズと位置に表示されます。
    '''    ・SW_SHOWMAXIMIZEDウィンドウをアクティブにし、最大化されたウィンドウとして表示されます。
    '''      このメンバーはウィンドウをアクティブにし、アイコンとして表示します。
    '''    ・SW_SHOWMINNOACTIVEウィンドウをアイコンとして表示します。 
    '''      現在アクティブなウィンドウは、アクティブなままです。
    '''    ・SW_SHOWNA現在の状態で、ウィンドウを表示します。 
    '''      現在アクティブなウィンドウは、アクティブなままです。
    '''    ・SW_SHOWNOACTIVATE最新のサイズと位置で、ウィンドウを表示します。 
    '''      現在アクティブなウィンドウは、アクティブなままです。
    '''    ・SW_SHOWNORMALにアクティブとウィンドウが表示されます。 
    '''      ウィンドウが最小化または最大化されている場合は、Windows が復元され、元のサイズと位置 (同じSW_RESTORE)。
    ''' </remarks>
    Public showCmd As Integer

    ''' <summary>minPosition</summary>
    ''' <remarks>
    '''   ウィンドウが最小化されているときは、ウィンドウの左上隅の位置を指定します。
    ''' </remarks>
    Public minPosition As POINT

    ''' <summary></summary>
    ''' <remarks>
    '''   ウィンドウを最大化するときは、ウィンドウの左上隅の位置を指定します。
    ''' </remarks>
    Public maxPosition As POINT

    ''' <summary></summary>
    ''' <remarks>
    '''   ウィンドウが標準の (復元) の位置にある場合は、ウィンドウの座標を指定します。
    ''' </remarks>
    Public normalPosition As RECT

End Structure

#End Region

''' <summary>ウィンドウの状態を保持するサンプルコード</summary>
''' <remarks>
'''   サンプルコードの取得先：https://msdn.microsoft.com/ja-jp/library/aa972163(v=vs.90).aspx
''' </remarks>
Partial Public Class MainWindow
    Inherits System.Windows.Window

#Region "DLL関数"

    ''' <summary>指定されたウィンドウの表示状態、および通常表示のとき、最小化されたとき、最大化されたときの位置を返します。</summary>
    ''' <param name="hWnd">ウィンドウのハンドル</param>
    ''' <param name="lpwndpl">位置データ</param>
    ''' <returns>
    '''   関数が成功すると、0 以外の値が返ります。
    '''   関数が失敗すると、0 が返ります。拡張エラー情報を取得するには、 関数を使います。
    ''' </returns>
    ''' <remarks>
    '''   この関数が取得する WINDOWPLACEMENT 構造体の flags メンバは、常に 0 です。
    '''   hWnd パラメータで指定したウィンドウが最大化されている場合、
    '''   showCmd メンバが SW_SHOWMAXIMIZED に設定されます。
    '''   ウィンドウが最小化されている場合は、showCmd メンバが SW_SHOWMINIMIZED に設定されます。
    '''   それ以外の場合は、SW_SHOWNORMAL に設定されます。
    '''   WINDOWPLACEMENT 構造体の length メンバは、sizeof(WINDOWPLACEMENT) に設定されていなければなりません。
    '''   このメンバが正しく設定されていないと、0（FALSE）が返ります。
    '''   ウィンドウの位置座標の正しい扱い方の詳細については、 構造体の説明を参照してください。
    ''' </remarks>
    <DllImport("user32.dll")>
    Private Shared Function GetWindowPlacement(ByVal hWnd As IntPtr, ByRef lpwndpl As WINDOWPLACEMENT) As Boolean
    End Function

    ''' <summary>指定されたウィンドウの表示状態を設定し、そのウィンドウの通常表示のとき、最小化されたとき、および最大化されたときの位置を設定します</summary>
    ''' <param name="hWnd">ウィンドウのハンドル</param>
    ''' <param name="lpwndpl">位置データ</param>
    ''' <returns>
    '''   関数が成功すると、0 以外の値が返ります。
    '''   関数が失敗すると、0 が返ります。拡張エラー情報を取得するには、 関数を使います。
    ''' </returns>
    ''' <remarks>
    '''   WINDOWPLACEMENT 構造体で指定された情報を適用するとウィンドウが完全に画面の外に出てしまう場合は、
    '''   ウィンドウが画面に現れるように座標が自動調整されます。
    '''   この調整では、画面の解像度の変更や複数モニタの構成も考慮されます。
    '''   WINDOWPLACEMENT 構造体の length メンバは、sizeof(WINDOWPLACEMENT) に設定されていなければなりません。
    '''   このメンバが正しく設定されていないと、0（FALSE）が返ります。
    '''   ウィンドウの位置座標の正しい扱い方の詳細については、 構造体の説明を参照してください。
    ''' </remarks>
    <DllImport("user32.dll")>
    Private Shared Function SetWindowPlacement(ByVal hWnd As IntPtr, ByRef lpwndpl As WINDOWPLACEMENT) As Boolean
    End Function

#End Region

    ''' <summary>SW_SHOWMINIMIZED</summary>
    ''' <remarks>
    '''   ウィンドウをアクティブにして最小化
    ''' </remarks>
    Private Const SW_SHOWMINIMIZED As Integer = 2

    ''' <summary>SW_SHOWNORMAL</summary>
    ''' <remarks>
    '''   ウィンドウをアクティブにして表示かつ、ウィンドウが最小化または最大化されているときは位置とサイズを元に戻す
    ''' </remarks>
    Private Const SW_SHOWNORMAL As Integer = 1

#Region "コンストラクタ"

    ''' <summary>コンストラクタ</summary>
    ''' <remarks></remarks>
    Public Sub New()

        InitializeComponent()

    End Sub

#End Region

#Region "イベント"

    ''' <summary>ウィンドウの初期化イベント</summary>
    ''' <param name="e">イベントデータ</param>
    ''' <remarks></remarks>
    Protected Overrides Sub OnSourceInitialized(ByVal e As EventArgs)

        '基底クラスのSourceInitializedイベントを発生させる
        MyBase.OnSourceInitialized(e)

        Try

            'ウィンドウの位置／サイズ／状態をアプリケーション設定のWindowPlacementプロパティから復元
            Dim wp As WINDOWPLACEMENT = My.Settings.WindowPlacement
            wp.length = Marshal.SizeOf(GetType(WINDOWPLACEMENT))
            wp.flags = 0

            'ウィンドウの状態が最小化の場合、位置とサイズを元に戻す
            wp.showCmd = IIf((wp.showCmd = SW_SHOWMINIMIZED), SW_SHOWNORMAL, wp.showCmd)

            'ウィンドウの位置／サイズ／状態を設定
            MainWindow.SetWindowPlacement(New WindowInteropHelper(Me).Handle, wp)

        Catch ex As Exception

        End Try

    End Sub

    ''' <summary>MainWindowのClosingイベント</summary>
    ''' <param name="e">キャンセルできるイベントのデータ</param>
    ''' <remarks></remarks>
    Protected Overrides Sub OnClosing(ByVal e As CancelEventArgs)

        '基底クラスのClosedイベントを発生させる
        MyBase.OnClosing(e)

        'ウィンドウの位置／サイズ／状態を取得
        Dim wp As WINDOWPLACEMENT = New WINDOWPLACEMENT
        MainWindow.GetWindowPlacement(New WindowInteropHelper(Me).Handle, wp)

        'ウィンドウの位置／サイズ／状態をアプリケーション設定のWindowPlacementプロパティへ代入
        My.Settings.WindowPlacement = wp

        'アプリケーション設定を保存
        My.Settings.Save()

    End Sub

#End Region

End Class
