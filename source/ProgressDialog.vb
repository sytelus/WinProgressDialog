'
'The .Net wrapper for the IProgressDialog which shows standard Windows progress dialogs
'Author: Shital Shah
'Date: Dec, 2003
'TLB source: http://www.msjogren.net/dotnet/eng/samples/vb6_progdlg.asp
'

Imports System.Runtime.InteropServices

Public Class ProgressDialog
    Implements IDisposable

    'Following is used to load standard video animations from shell
    Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Integer
    Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Integer) As Integer

    ' File operation animations resource IDs in shell32.dll
    Public Enum StandardProgressAnimations
        FileMove = 160
        FileCopy = 161
        FlyingPapers = 165
        SearchGlobe = 166
        PermanentDelete = 164
        FromRecycleBinDelete = 163
        ToRecycleBinDelete = 162
        SearchComputer = 152
        SearchDocument = 151
        SearchFlashlight = 150
        Custom = 0
    End Enum

    Private m_DescriptionLine1 As String = ""
    Public Property DescriptionLine1() As String
        Get
            Return m_DescriptionLine1
        End Get
        Set(ByVal Value As String)
            m_WinProgress.SetLine(1, Value, 1, IntPtr.Zero)
            m_DescriptionLine1 = Value
        End Set
    End Property

    Private m_DescriptionLine2 As String = ""
    Public Property DescriptionLine2() As String
        Get
            Return m_DescriptionLine2
        End Get
        Set(ByVal Value As String)
            m_WinProgress.SetLine(2, Value, 1, IntPtr.Zero)
            m_DescriptionLine2 = Value
        End Set
    End Property

    Private m_DescriptionLine3 As String = ""
    Public Property DescriptionLine3() As String
        Get
            Return m_DescriptionLine3
        End Get
        Set(ByVal Value As String)
            m_WinProgress.SetLine(3, Value, 1, IntPtr.Zero)
            m_DescriptionLine3 = Value
        End Set
    End Property

    Private m_DialogTitle As String
    Public Property DialogTitle() As String
        Get
            Return m_DialogTitle
        End Get
        Set(ByVal Value As String)
            m_WinProgress.SetTitle(Value)
            m_DialogTitle = Value
        End Set
    End Property

    Private m_CancelButtonCaption As String
    Public Property CancelButtonCaption() As String
        Get
            Return m_CancelButtonCaption
        End Get
        Set(ByVal Value As String)
            m_WinProgress.SetCancelMsg(Value, IntPtr.Zero)
            m_CancelButtonCaption = Value
        End Set
    End Property

    Private m_ProgressAnimation As StandardProgressAnimations = StandardProgressAnimations.FlyingPapers
    Public Property ProgressAnimation() As StandardProgressAnimations
        Get
            Return m_ProgressAnimation
        End Get
        Set(ByVal Value As StandardProgressAnimations)
            If (m_Shell32Handle = 0) And (Value <> StandardProgressAnimations.Custom) Then
                m_Shell32Handle = LoadLibrary("shell32.dll")
            End If
            If Value <> StandardProgressAnimations.Custom Then
                m_WinProgress.SetAnimation(m_Shell32Handle, CType(Value, Short))
            End If
            m_ProgressAnimation = Value
        End Set
    End Property

    Private m_ProgressBarVisible As Boolean = True
    Public Property ProgressBarVisible() As Boolean
        Get
            Return m_ProgressBarVisible
        End Get
        Set(ByVal Value As Boolean)
            m_ProgressBarVisible = Value
        End Set
    End Property


    Private m_WinProgress As VBProgressDialog.IProgressDialog = Nothing
    Private m_Shell32Handle As Integer = 0

    Public Sub New()
        m_WinProgress = New VBProgressDialog.ProgressDialogClass
        Me.DescriptionLine1 = "Performing task"
        Me.DescriptionLine2 = "Please wait..."
        Me.DialogTitle = "Progress Status"
    End Sub

    Public Sub Show(ByVal parentHandle As Integer)
        Me.Show(parentHandle, Me.DialogTitle, Me.DescriptionLine1)
    End Sub
    Public Sub Show(ByVal parentHandle As Integer, ByVal title As String, ByVal description As String, ByVal maxValue As Decimal)
        If (m_Shell32Handle = 0) And (Me.ProgressAnimation = StandardProgressAnimations.FlyingPapers) Then
            Me.ProgressAnimation = StandardProgressAnimations.FlyingPapers
        End If
        Dim flags As VBProgressDialog.PROGDLG_FLAGS = VBProgressDialog.PROGDLG_FLAGS.PROGDLG_MODAL Or VBProgressDialog.PROGDLG_FLAGS.PROGDLG_AUTOTIME Or VBProgressDialog.PROGDLG_FLAGS.PROGDLG_NOMINIMIZE
        If m_ProgressBarVisible = False Then
            flags = flags Or VBProgressDialog.PROGDLG_FLAGS.PROGDLG_NOPROGRESSBAR
        End If
        Me.MaxValue = maxValue
        Me.SetupDialog(title, description)
        m_WinProgress.StartProgressDialog(parentHandle, Nothing, flags, IntPtr.Zero)
    End Sub
    Public Sub Show(ByVal parentHandle As Integer, ByVal title As String, ByVal description As String)
        Me.Show(parentHandle, title, description, Me.MaxValue)
    End Sub

    Public Sub SetupDialog(ByVal title As String, ByVal description As String)
        Me.SetupDialog(title, description, Me.ProgressAnimation)
    End Sub
    Public Sub SetupDialog(ByVal title As String, ByVal description As String, ByVal animationStyle As StandardProgressAnimations)
        Me.DialogTitle = title
        Me.DescriptionLine1 = description
        Me.ProgressAnimation = animationStyle
    End Sub

    Public Function HasUserCancelled() As Boolean
        Return (m_WinProgress.HasUserCancelled <> 0)
    End Function

    Public Sub Hide()
        m_WinProgress.StopProgressDialog()
    End Sub

    Private m_MaxValue As Decimal = Decimal.MaxValue
    Public Property MaxValue() As Decimal
        Get
            Return m_MaxValue
        End Get
        Set(ByVal Value As Decimal)
            m_MaxValue = Value
        End Set
    End Property

    Public Function UpdateProgress(ByVal current As Decimal, ByVal progressDetail As String) As Boolean
        Return Me.UpdateProgress(current, Me.MaxValue, progressDetail)
    End Function
    Public Function UpdateProgress(ByVal current As Decimal, ByVal total As Decimal, ByVal progressDetail As String) As Boolean
        MaxValue = total
        m_WinProgress.SetProgress64(current, total)
        If Not progressDetail Is Nothing Then
            Me.DescriptionLine2 = progressDetail
        End If
        Return Me.HasUserCancelled
    End Function
    Public Function UpdateProgress(ByVal current As Decimal, ByVal total As Decimal) As Boolean
        Return Me.UpdateProgress(current, total, Nothing)
    End Function
    Public Function UpdateProgress(ByVal current As Decimal) As Boolean
        Return Me.UpdateProgress(current, Me.MaxValue, Nothing)
    End Function

    Public Sub ResetEstimationTimer()
        m_WinProgress.Timer(VBProgressDialog.PDTIMER_FLAGS.PDTIMER_RESET, IntPtr.Zero)
    End Sub

    Public Sub Dispose() Implements System.IDisposable.Dispose
        If m_Shell32Handle <> 0 Then
            FreeLibrary(m_Shell32Handle)
        End If
        If Not m_WinProgress Is Nothing Then
            Try
                m_WinProgress.StopProgressDialog()
            Finally
                Marshal.ReleaseComObject(m_WinProgress)
            End Try
        End If
    End Sub
End Class
