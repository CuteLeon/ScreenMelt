Module mScreenMelt
    Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Integer) As Integer
    Private Declare Function GetDesktopWindow Lib "user32" () As Integer
    Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer) As Integer
    Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Integer) As Integer
    Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Integer, ByVal hObject As Integer) As Integer
    Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Integer, ByVal x As Integer, ByVal y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hSrcDC As Integer, ByVal xSrc As Integer, ByVal ySrc As Integer, ByVal dwRop As Integer) As Integer
    Private Declare Sub InvalidateRect Lib "user32" (ByVal hwnd As Integer, lpRect As IntPtr, ByVal bErase As Integer)
    Private Const FragmentWidth = 30
    Private Const FragmentHeight = 30
    Private Const Excursion = 1
    Private Const SRCCOPY = &HCC0020
    Dim X As Integer, Y As Integer
    Dim FragmentHDC， FragmentBitmap As Integer '碎片
    Dim DesktopHDC As Integer '桌面
    Dim ScreenWidth As Integer = My.Computer.Screen.Bounds.Width
    Dim ScreenHeight As Integer = My.Computer.Screen.Bounds.Height

    Public Function Main() As Form
        '从桌面的句柄获取桌面HDC
        DesktopHDC = GetWindowDC(GetDesktopWindow())
        '从桌面HDC创建兼容的HDC
        FragmentHDC = CreateCompatibleDC(DesktopHDC)
        '从桌面HDC创建兼容的位图
        FragmentBitmap = CreateCompatibleBitmap(DesktopHDC, FragmentWidth, FragmentHeight)

        SelectObject(FragmentHDC, FragmentBitmap)
        Dim Index As Integer
        Do While True
            For Index = 0 To 100
                '随机产生碎片的位置
                X = (My.Computer.Screen.Bounds.Width) * Rnd()
                Y = (My.Computer.Screen.Bounds.Height) * Rnd()
                '把碎片位图从桌面位图复制到碎片HDC
                BitBlt(FragmentHDC, 0, 0, FragmentWidth, FragmentHeight, DesktopHDC, X, Y, SRCCOPY)
                '碎片位图位置随机偏移
                X = X + (Excursion - (Excursion * 2) * Rnd())
                Y = Y + (Excursion - (Excursion * 2) * Rnd())
                '重新复制到桌面HDC
                BitBlt(DesktopHDC, X, Y, FragmentWidth, FragmentHeight, FragmentHDC, 0, 0, SRCCOPY)
            Next
            '线程停顿
            Threading.Thread.Sleep(1)
        Loop

        Return New Form
    End Function
End Module
