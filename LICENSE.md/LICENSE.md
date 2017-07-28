'枚举窗口列表中的所有父窗口 (顶级和被所有窗口)
'返回值 Long，非零表示成功，零表示失败
'lpEnumFunc Long，指向为每个子窗口都调用的一个函数的指针。用AddressOf运算符获得函数在标准模式下的地址
'lParam Long，在枚举期间，传递给dwcbkd32.ocx定制控件之EnumWindows事件的值。这个值的含义是由程序员规定的
Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long)
'为指定的父窗口枚举子窗口
'返回值 Long，非零表示成功，零表示失败
'hWndParent Long，欲枚举子窗口的父窗口的句柄
'lpEnumFunc Long，为每个子窗口调用的函数的指针。用AddressOf运算符获得函数在一个标准模块中的地址
'lParam Long，在枚举期间，传递给dwcbkd32.ocx定制控件之EnumWindows事件的值。这个值的含义是由程序员规定的。（原文：Value that is passed to the EnumWindows event of the dwcbkd32.ocx custom control during enumeration. The meaning of this value is defined by the programmer.）
Private Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function EnumWindowStations Lib "user32" Alias "EnumWindowStationsA" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long


Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long


Dim ChildHwnd As String

Sub test_Main()

    Dim hwnd As Long
    
    hwnd = FindWindow(vbNullString, "RIMSII Toolbar - Real RIMS2 (AD Hong Kong)")
    
    ChildHwnd = ""
    Call EnumChildWindows(hwnd, AddressOf EnumChildProc, ByVal 0&)
    ChildHwnd = VBA.Mid(ChildHwnd, 2)
    AllHwnd = Split(ChildHwnd, ",")
    


End Sub



'这是一个回调函数, 必须放在模块中. 用来遍历指定窗口的子窗口(控件). 这里参数中的 hWnd 即为子窗口(控件)句柄
Public Function EnumChildProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
ChildHwnd = ChildHwnd & "," & hwnd
EnumChildProc = 1
End Function


' 函数: FGetClassName
' 功能: 返回指定窗口中的类型
' 参数: hWnd 指定窗口的句柄
' 返回: 指定窗口的类型

Public Function FGetClassName(hwnd As Long) As String
Dim ClassName As String
Dim Ret As Long
'填充缓冲(如果填充太小返回会不完整).
 ClassName = VBA.Space(256)

 '调用 GetClassName 函数, 返回值为类型名的实际长度.
 Ret = GetClassName(hwnd, ClassName, 256)

 '函数返回类型. Ret 为上一步所得到的类型名的实际长度.
 FGetClassName = VBA.Left(ClassName, Ret)
End Function

' 函数: GetText
' 功能: 返回指定窗口(如文本框)中的文字
' 参数: WindowHandle 指定窗口的句柄
' 返回: 指定窗口的文字
Public Function GetText(WindowHandle As Long) As String
Dim strBuffer As String '字符串缓冲
Dim Char As String '储存密码掩码以待恢复

    '填充缓冲(如果填充太小返回会不完整).
    strBuffer = VBA.Space(255)

    '发送消息 EM_GETPASSWORDCHAR(返回密码掩码) 给指定窗口. 这里返回掩码给Char(比如可能 Char=*).
    Char = SendMessage(WindowHandle, &HD2, 0, 0)

    '发送消息 EM_SETPASSWORDCHAR(设置密码掩码) 给指定窗口. 这里设置了0(Null), 即除去除密码掩码.
    PostMessage WindowHandle, &HCC, 0, 0

    '如果是Edit控件则等待消息发送成功, 即等待掩码被去除.
    If InStr("Edit", FGetClassName(WindowHandle)) And Char <> "0" Then Sleep (10)

    '发送消息 WM_GETTEXT(返回所含文字) 给指定窗口. 这里得到Edit控件的文字, 即密码. 注意"ByVal", 如果少这个则VB崩溃.
    SendMessage WindowHandle, &HD, 255, ByVal strBuffer

    '发送消息 EM_SETPASSWORDCHAR(设置密码掩码) 给指定窗口. 这里设置为Char, 即恢复原先掩码.
    PostMessage WindowHandle, &HCC, ByVal Char, 0

    '函数返回所得文字(密码), 之所以要用Trim去空格是因为第一步中用空格填充了255个字符.
    GetText = VBA.Trim(strBuffer)
End Function


''为指定的父窗口枚举子窗口
''返回值 Long，非零表示成功，零表示失败
''hWndParent Long，欲枚举子窗口的父窗口的句柄
''lpEnumFunc Long，为每个子窗口调用的函数的指针。用AddressOf运算符获得函数在一个标准模块中的地址
''lParam Long，在枚举期间，传递给dwcbkd32.ocx定制控件之EnumWindows事件的值。这个值的含义是由程序员规定的。（原文：Value that is passed to the EnumWindows event of the dwcbkd32.ocx custom control during enumeration. The meaning of this value is defined by the programmer.）
'Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal LPARAM As Long) As Long

'由于上面这个函数每次调用都会得到下一个子窗体（控件）的句柄，并赋值给hWnd,实际使用中，我把所有子句柄存放在ChildHwnd字符串中，遍历完毕，再

'Dim AllHwnd() As String

'去除多余的无效字符
'ChildHwnd =vba. Mid(ChildHwnd, 2)
'转换成数组
'AllHwnd =vba. Split(ChildHwnd, ",")

Public Function EnumWindowsProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
    Dim lpBuffer As String * 1024
    Dim dwWindowCaption As String
    Dim lpLength As Long

    lpLength = GetWindowText(hwnd, lpBuffer, 1024)
    dwWindowCaption = Left(lpBuffer, lpLength)
'    MsgBox dwWindowCaption
    Debug.Print dwWindowCaption

    If InStr(dwWindowCaption, "Word") > 0 Then
        '停止查找函数返回0
        EnumWindowsProc = 0
    Else
        '继续查找函数返回1
        EnumWindowsProc = 1
    End If

End Function
