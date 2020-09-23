VERSION 5.00
Begin VB.Form frmSBChanger 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Look at the startbutton!"
   ClientHeight    =   75
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2370
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   75
   ScaleWidth      =   2370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrPaint 
      Interval        =   1
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "frmSBChanger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ***************************************************
' * SBChanger, by Cakkie - slisse@planetinternet.be *
' ***************************************************
' * Purpose: Change the picture on the start button *
' * Usage: Just place a bitmap in the applications  *
' *        folder and name it sb.bmp                *
' *        Bitmap width/height: 55px/22px           *
' * Other: Have fun                                 *
' ***************************************************

Dim hwndTB As Long ' handle to taskbar window
Dim hWndSB As Long ' handle to startbutton window
Dim hDcSB As Long ' handle to startbutton device context
Dim mRect As RECT ' will hold coordinates for start button

Dim hDcTmp As Long ' handle to device context to hold new picture
Dim hBmpTmp As Long ' temporary bitmap
Dim hBmpTmp2 As Long ' temporary bitmap used for cleanup
Dim nWidth As Long ' width of startbutton
Dim nHeight As Long ' height of startbutton

Dim sPath As String ' path to picture
    
Private Sub Form_Load()

    'Get handle to taskbar and startbutton
    hwndTB = FindWindow("Shell_TrayWnd", "")
    hWndSB = FindWindowEx(hwndTB, 0, "button", vbNullString)

    ' get dc of startbutton
    hDcSB = GetWindowDC(hWndSB)
    
    ' get coordinates of startbutton
    Call GetWindowRect(hWndSB, mRect)
    
    ' compute width and height
    nWidth = mRect.Right - mRect.Left
    nHeight = mRect.Bottom - mRect.Top
    
    ' initialize dc and bitmap
    hDcTmp = CreateCompatibleDC(hDcSB)
    hBmpTmp = CreateCompatibleBitmap(hDcTmp, nWidth, nHeight)
    
    ' set path and load picture
    sPath = App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "sb.bmp"
    hBmpTmp2 = SelectObject(hDcTmp, LoadPicture(sPath))

End Sub

Private Sub tmrPaint_Timer()
    
    ' paint to startbutton
    Call BitBlt(hDcSB, 0, 0, nWidth, nHeight, hDcTmp, 0, 0, SRCCOPY)

End Sub

Private Sub Form_Unload(Cancel As Integer)

    ' clean up the mess, keep the planet clean
    hBmpTmp = SelectObject(hDcTmp, hBmpTmp2)
    DeleteObject hBmpTmp
    DeleteDC hDcTmp

End Sub
