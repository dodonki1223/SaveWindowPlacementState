Imports System.ComponentModel
Imports System.Windows.Interop
Imports System.Runtime.InteropServices

#Region "�\����"

'------------------------------------------------���\���̃�����------------------------------------------------'
'   StructLayout�ɂ���                                                                                       '
'     ���O���: System.Runtime.InteropServices                                                                 '
'     �A�Z���u��: mscorlib                                                                                     '
'     �p���E�C���^�[�t�F�C�X: Attribute, _Attribute                                                            '
'     StructLayout�����́A��������ł̃t�B�[���h(�����o�ϐ�)�̔z�u���@���w�肷�邽�߂̑����ł�                 '
'   LayoutKind.Sequential�ɂ���                                                                              '
'     �����^�C���ɂ�鎩���I�ȕ��בւ����s�킸�A�R�[�h��ŋL�q����Ă��鏇���̂܂܃t�B�[���h��z�u����         '
'   ���}�l�[�W���ŗp������\���̂̓A���}�l�[�W���̂��̂ƌ݊���������܂���B                             '
'     ���݉^�p�ł͍\���̂̒�`��StructLayout�A�g���r���[�g��t�����邱�ƂŃA���}�l�[�W�݊��̍\���̂��`���܂� '
'--------------------------------------------------------------------------------------------------------------'

''' <summary>POINT�\����</summary>
''' <remarks>
'''     2 ������Ԃɂ�����Ax ���W�� y ���W�̑g��\���܂�
''' </remarks>
<Serializable(), StructLayout(LayoutKind.Sequential)>
Public Structure POINT

    ''' <summary>X���W</summary>
    ''' <remarks></remarks>
    Public X As Integer

    ''' <summary>Y���W</summary>
    ''' <remarks></remarks>
    Public Y As Integer

    ''' <summary>�R���X�g���N�^</summary>
    ''' <param name="pX">X���W</param>
    ''' <param name="pY">Y���W</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal pX As Integer, ByVal pY As Integer)

        Me.X = X
        Me.Y = Y

    End Sub

End Structure

''' <summary>RECT�\����</summary>
''' <remarks>
'''     �l�p�`�̍��������щE�����̍��W���`���܂�
''' </remarks>
<Serializable(), StructLayout(LayoutKind.Sequential)>
Public Structure RECT

    ''' <summary>Left</summary>
    ''' <remarks>�l�p�`�̍�����̂w���W���w�肵�܂�</remarks>
    Public Left As Integer

    ''' <summary>Top</summary>
    ''' <remarks>�l�p�`�̍�����̂x���W���w�肵�܂�</remarks>
    Public Top As Integer

    ''' <summary>Right</summary>
    ''' <remarks>�l�p�`�̉E�����̂w���W���w�肵�܂�</remarks>
    Public Right As Integer

    ''' <summary>Bottom</summary>
    ''' <remarks>�l�p�`�̉E�����̂x���W���w�肵�܂�</remarks>
    Public Bottom As Integer

    ''' <summary>�R���X�g���N�^</summary>
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

''' <summary>WINDOWPLACEMENT�\����</summary>
''' <remarks>
'''     ��ʏ�̃E�B���h�E�̔z�u�ɂ��Ă̏�񂪊܂܂�Ă��܂�
''' </remarks>
<Serializable(), StructLayout(LayoutKind.Sequential)>
Public Structure WINDOWPLACEMENT

    ''' <summary>length</summary>
    ''' <remarks>
    '''   �\���̃T�C�Y���o�C�g�P�ʂŎw�肵�܂��B
    ''' </remarks>
    Public length As Integer

    ''' <summary>flags</summary>
    ''' <remarks>
    '''   �ŏ������ꂽ�E�B���h�E�ƃE�B���h�E�𕜌�������@�̈ʒu�𐧌䂷��t���O���w�肵�܂��B
    '''   ���̃����o�[�́A���̃t���O�̈���܂��͗������w��ł��܂��B
    '''     �EWPF_SETMINPOSITION�ŏ������ꂽ�E�B���h�E�� x �ʒu�� y �ʒu���w��ł��邱�Ƃ��������܂��B
    '''       ���̃t���O�ł���K�v������܂��w�肳�ꂽ���W�Őݒ肳��Ă��邩�ǂ����A ptMinPosition�����o�[�ł��B
    '''     �EWPF_RESTORETOMAXIMIZED�������ꂽ�E�B���h�E���ő剻�����A�ŏ�������O�ɁA
    '''       �ő剻����Ă������ǂ����Ɋ֌W�Ȃ����w�肵�܂��B ���̐ݒ�́A�E�B���h�E������Ƃ��ɂ̂ݗL���ł��B 
    '''       ����̕�������͕ς��܂���B ���̃t���O�͗L���ȏꍇ�ɂ̂݁A���̃����o�[�͒l���w�肳��āA showCmd�����o�[�ł��B
    ''' </remarks>
    Public flags As Integer

    ''' <summary>showCmd</summary>
    ''' <remarks>
    '''   �E�B���h�E�̌��݂̕\����Ԃ��w�肵�܂��B ���̃����o�[�́A���̒l�̂����ꂩ�ɂȂ�܂��B
    '''    �ESW_HIDE�E�B���h�E���\���ɂ��A�ʂ̃E�B���h�E���A�N�e�B�u����n���܂��B
    '''    �ESW_MINIMIZE�w�肳�ꂽ�E�B���h�E���ŏ������A�V�X�e���̈ꗗ�̍ŏ�ʃE�B���h�E��\�����܂��B
    '''    �ESW_RESTORE�ɃA�N�e�B�u�ƃE�B���h�E���\������܂��B 
    '''      �E�B���h�E���ŏ����܂��͍ő剻����Ă���ꍇ�́AWindows ����������A���̃T�C�Y�ƈʒu (����SW_SHOWNORMAL)�B
    '''    �ESW_SHOW�E�B���h�E���A�N�e�B�u�ɂ��A���݂̃T�C�Y�ƈʒu�ɕ\������܂��B
    '''    �ESW_SHOWMAXIMIZED�E�B���h�E���A�N�e�B�u�ɂ��A�ő剻���ꂽ�E�B���h�E�Ƃ��ĕ\������܂��B
    '''      ���̃����o�[�̓E�B���h�E���A�N�e�B�u�ɂ��A�A�C�R���Ƃ��ĕ\�����܂��B
    '''    �ESW_SHOWMINNOACTIVE�E�B���h�E���A�C�R���Ƃ��ĕ\�����܂��B 
    '''      ���݃A�N�e�B�u�ȃE�B���h�E�́A�A�N�e�B�u�Ȃ܂܂ł��B
    '''    �ESW_SHOWNA���݂̏�ԂŁA�E�B���h�E��\�����܂��B 
    '''      ���݃A�N�e�B�u�ȃE�B���h�E�́A�A�N�e�B�u�Ȃ܂܂ł��B
    '''    �ESW_SHOWNOACTIVATE�ŐV�̃T�C�Y�ƈʒu�ŁA�E�B���h�E��\�����܂��B 
    '''      ���݃A�N�e�B�u�ȃE�B���h�E�́A�A�N�e�B�u�Ȃ܂܂ł��B
    '''    �ESW_SHOWNORMAL�ɃA�N�e�B�u�ƃE�B���h�E���\������܂��B 
    '''      �E�B���h�E���ŏ����܂��͍ő剻����Ă���ꍇ�́AWindows ����������A���̃T�C�Y�ƈʒu (����SW_RESTORE)�B
    ''' </remarks>
    Public showCmd As Integer

    ''' <summary>minPosition</summary>
    ''' <remarks>
    '''   �E�B���h�E���ŏ�������Ă���Ƃ��́A�E�B���h�E�̍�����̈ʒu���w�肵�܂��B
    ''' </remarks>
    Public minPosition As POINT

    ''' <summary></summary>
    ''' <remarks>
    '''   �E�B���h�E���ő剻����Ƃ��́A�E�B���h�E�̍�����̈ʒu���w�肵�܂��B
    ''' </remarks>
    Public maxPosition As POINT

    ''' <summary></summary>
    ''' <remarks>
    '''   �E�B���h�E���W���� (����) �̈ʒu�ɂ���ꍇ�́A�E�B���h�E�̍��W���w�肵�܂��B
    ''' </remarks>
    Public normalPosition As RECT

End Structure

#End Region

''' <summary>�E�B���h�E�̏�Ԃ�ێ�����T���v���R�[�h</summary>
''' <remarks>
'''   �T���v���R�[�h�̎擾��Fhttps://msdn.microsoft.com/ja-jp/library/aa972163(v=vs.90).aspx
''' </remarks>
Partial Public Class MainWindow
    Inherits System.Windows.Window

#Region "DLL�֐�"

    ''' <summary>�w�肳�ꂽ�E�B���h�E�̕\����ԁA����ђʏ�\���̂Ƃ��A�ŏ������ꂽ�Ƃ��A�ő剻���ꂽ�Ƃ��̈ʒu��Ԃ��܂��B</summary>
    ''' <param name="hWnd">�E�B���h�E�̃n���h��</param>
    ''' <param name="lpwndpl">�ʒu�f�[�^</param>
    ''' <returns>
    '''   �֐�����������ƁA0 �ȊO�̒l���Ԃ�܂��B
    '''   �֐������s����ƁA0 ���Ԃ�܂��B�g���G���[�����擾����ɂ́A �֐����g���܂��B
    ''' </returns>
    ''' <remarks>
    '''   ���̊֐����擾���� WINDOWPLACEMENT �\���̂� flags �����o�́A��� 0 �ł��B
    '''   hWnd �p�����[�^�Ŏw�肵���E�B���h�E���ő剻����Ă���ꍇ�A
    '''   showCmd �����o�� SW_SHOWMAXIMIZED �ɐݒ肳��܂��B
    '''   �E�B���h�E���ŏ�������Ă���ꍇ�́AshowCmd �����o�� SW_SHOWMINIMIZED �ɐݒ肳��܂��B
    '''   ����ȊO�̏ꍇ�́ASW_SHOWNORMAL �ɐݒ肳��܂��B
    '''   WINDOWPLACEMENT �\���̂� length �����o�́Asizeof(WINDOWPLACEMENT) �ɐݒ肳��Ă��Ȃ���΂Ȃ�܂���B
    '''   ���̃����o���������ݒ肳��Ă��Ȃ��ƁA0�iFALSE�j���Ԃ�܂��B
    '''   �E�B���h�E�̈ʒu���W�̐������������̏ڍׂɂ��ẮA �\���̂̐������Q�Ƃ��Ă��������B
    ''' </remarks>
    <DllImport("user32.dll")>
    Private Shared Function GetWindowPlacement(ByVal hWnd As IntPtr, ByRef lpwndpl As WINDOWPLACEMENT) As Boolean
    End Function

    ''' <summary>�w�肳�ꂽ�E�B���h�E�̕\����Ԃ�ݒ肵�A���̃E�B���h�E�̒ʏ�\���̂Ƃ��A�ŏ������ꂽ�Ƃ��A����эő剻���ꂽ�Ƃ��̈ʒu��ݒ肵�܂�</summary>
    ''' <param name="hWnd">�E�B���h�E�̃n���h��</param>
    ''' <param name="lpwndpl">�ʒu�f�[�^</param>
    ''' <returns>
    '''   �֐�����������ƁA0 �ȊO�̒l���Ԃ�܂��B
    '''   �֐������s����ƁA0 ���Ԃ�܂��B�g���G���[�����擾����ɂ́A �֐����g���܂��B
    ''' </returns>
    ''' <remarks>
    '''   WINDOWPLACEMENT �\���̂Ŏw�肳�ꂽ����K�p����ƃE�B���h�E�����S�ɉ�ʂ̊O�ɏo�Ă��܂��ꍇ�́A
    '''   �E�B���h�E����ʂɌ����悤�ɍ��W��������������܂��B
    '''   ���̒����ł́A��ʂ̉𑜓x�̕ύX�╡�����j�^�̍\�����l������܂��B
    '''   WINDOWPLACEMENT �\���̂� length �����o�́Asizeof(WINDOWPLACEMENT) �ɐݒ肳��Ă��Ȃ���΂Ȃ�܂���B
    '''   ���̃����o���������ݒ肳��Ă��Ȃ��ƁA0�iFALSE�j���Ԃ�܂��B
    '''   �E�B���h�E�̈ʒu���W�̐������������̏ڍׂɂ��ẮA �\���̂̐������Q�Ƃ��Ă��������B
    ''' </remarks>
    <DllImport("user32.dll")>
    Private Shared Function SetWindowPlacement(ByVal hWnd As IntPtr, ByRef lpwndpl As WINDOWPLACEMENT) As Boolean
    End Function

#End Region

    ''' <summary>SW_SHOWMINIMIZED</summary>
    ''' <remarks>
    '''   �E�B���h�E���A�N�e�B�u�ɂ��čŏ���
    ''' </remarks>
    Private Const SW_SHOWMINIMIZED As Integer = 2

    ''' <summary>SW_SHOWNORMAL</summary>
    ''' <remarks>
    '''   �E�B���h�E���A�N�e�B�u�ɂ��ĕ\�����A�E�B���h�E���ŏ����܂��͍ő剻����Ă���Ƃ��͈ʒu�ƃT�C�Y�����ɖ߂�
    ''' </remarks>
    Private Const SW_SHOWNORMAL As Integer = 1

#Region "�R���X�g���N�^"

    ''' <summary>�R���X�g���N�^</summary>
    ''' <remarks></remarks>
    Public Sub New()

        InitializeComponent()

    End Sub

#End Region

#Region "�C�x���g"

    ''' <summary>�E�B���h�E�̏������C�x���g</summary>
    ''' <param name="e">�C�x���g�f�[�^</param>
    ''' <remarks></remarks>
    Protected Overrides Sub OnSourceInitialized(ByVal e As EventArgs)

        '���N���X��SourceInitialized�C�x���g�𔭐�������
        MyBase.OnSourceInitialized(e)

        Try

            '�E�B���h�E�̈ʒu�^�T�C�Y�^��Ԃ��A�v���P�[�V�����ݒ��WindowPlacement�v���p�e�B���畜��
            Dim wp As WINDOWPLACEMENT = My.Settings.WindowPlacement
            wp.length = Marshal.SizeOf(GetType(WINDOWPLACEMENT))
            wp.flags = 0

            '�E�B���h�E�̏�Ԃ��ŏ����̏ꍇ�A�ʒu�ƃT�C�Y�����ɖ߂�
            wp.showCmd = IIf((wp.showCmd = SW_SHOWMINIMIZED), SW_SHOWNORMAL, wp.showCmd)

            '�E�B���h�E�̈ʒu�^�T�C�Y�^��Ԃ�ݒ�
            MainWindow.SetWindowPlacement(New WindowInteropHelper(Me).Handle, wp)

        Catch ex As Exception

        End Try

    End Sub

    ''' <summary>MainWindow��Closing�C�x���g</summary>
    ''' <param name="e">�L�����Z���ł���C�x���g�̃f�[�^</param>
    ''' <remarks></remarks>
    Protected Overrides Sub OnClosing(ByVal e As CancelEventArgs)

        '���N���X��Closed�C�x���g�𔭐�������
        MyBase.OnClosing(e)

        '�E�B���h�E�̈ʒu�^�T�C�Y�^��Ԃ��擾
        Dim wp As WINDOWPLACEMENT = New WINDOWPLACEMENT
        MainWindow.GetWindowPlacement(New WindowInteropHelper(Me).Handle, wp)

        '�E�B���h�E�̈ʒu�^�T�C�Y�^��Ԃ��A�v���P�[�V�����ݒ��WindowPlacement�v���p�e�B�֑��
        My.Settings.WindowPlacement = wp

        '�A�v���P�[�V�����ݒ��ۑ�
        My.Settings.Save()

    End Sub

#End Region

End Class
