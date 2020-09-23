VERSION 5.00
Begin VB.UserControl ucThumbs 
   Alignable       =   -1  'True
   BackColor       =   &H80000001&
   ClientHeight    =   7305
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3090
   KeyPreview      =   -1  'True
   MousePointer    =   1  'Arrow
   PropertyPages   =   "ucThumbs.ctx":0000
   ScaleHeight     =   487
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   206
   ToolboxBitmap   =   "ucThumbs.ctx":0011
   Begin VB.PictureBox p2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   1
      Left            =   0
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   189
      TabIndex        =   4
      Top             =   0
      Width           =   2835
      Begin VB.TextBox srhTxt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   30
         MousePointer    =   3  'I-Beam
         TabIndex        =   5
         Top             =   30
         Width           =   2775
      End
   End
   Begin VB.PictureBox p2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Index           =   0
      Left            =   0
      ScaleHeight     =   3
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   189
      TabIndex        =   3
      Top             =   7230
      Width           =   2835
   End
   Begin VB.VScrollBar VS 
      Height          =   7215
      Left            =   2850
      MousePointer    =   1  'Arrow
      TabIndex        =   2
      Top             =   30
      Width           =   195
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   7005
      Left            =   30
      MousePointer    =   1  'Arrow
      ScaleHeight     =   467
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   185
      TabIndex        =   1
      Top             =   270
      Width           =   2775
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         Height          =   1125
         Index           =   0
         Left            =   150
         Top             =   150
         Visible         =   0   'False
         Width           =   1125
      End
   End
   Begin VB.FileListBox File1 
      Height          =   675
      Left            =   1500
      TabIndex        =   0
      Top             =   2820
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Menu Popup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu Rename 
         Caption         =   "Rename"
      End
      Begin VB.Menu Delete 
         Caption         =   "Delete"
      End
      Begin VB.Menu Move 
         Caption         =   "Move"
      End
      Begin VB.Menu Copy 
         Caption         =   "Copy"
      End
      Begin VB.Menu sp 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu Properties 
         Caption         =   "Properties"
      End
      Begin VB.Menu Browse 
         Caption         =   "Browse"
      End
      Begin VB.Menu View 
         Caption         =   "View"
         Begin VB.Menu sz 
            Caption         =   "Small"
            Index           =   0
         End
         Begin VB.Menu sz 
            Caption         =   "Medium"
            Index           =   1
         End
         Begin VB.Menu sz 
            Caption         =   "Large"
            Index           =   2
         End
         Begin VB.Menu sz 
            Caption         =   "Custom"
            Index           =   3
         End
         Begin VB.Menu sz 
            Caption         =   "Auto Size"
            Index           =   4
         End
      End
   End
End
Attribute VB_Name = "ucThumbs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
'This is an Image File list box with many of the properties of a file list. Created 02-08-08 by: Russell Sanders
'
'Prperties,Events to know:
'   MouseDown, Move, Up all fire with one extra param. "imgIndex"
'
'
'dependent code
'   MODULES:
'       GDIPlusAPI from Avery his code named; "Use GDI+ ( aka GDIPlus ) with VB6 and Win98!", found here:
'       http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=37541&lngWId=1
'       There have been no changes made to his original code. I did add two veriables to keep me from
'       having to add a new module; but, they are in no way conected to Averys' code.
'   Class
'       cDlg
'       This file is from http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=9260&lngWId=1
'       I have added "browse for folder", "Copy To", "Move To", "Delete to recycle bin", "Rename", And
'       "File Properties". These additions were made by me at different times and have different sources
'       for the most part "MSDN" and "AllAPI.net"
'       mDlg
'       This file was added to allow the user to select the starting directory
'       for cDlg.Thanks to Ed Wilk for pointing this out and also for the link
'       to the code used to fix the problem: http://vbnet.mvps.org/code/callback/browsecallback.htm
'       NOTE:
'           the "cDlg" class is declared in the "mDlg" module as "CD" so once these two files are added to the
'           project you need do nothing but call it's functions and subs like "FileName = CD.ShowOpen".
'           or "Folder = CD.BrowseForFolder"
'   Property Page
'       propThumbs
'       I made this just to be doing something.
'
'ToDo's
'   add the abillity to use as a list or a filelist. As a list the user would assign each file
'   to include regardless of it's path. And like it is now as a file list all images in a path
'   with a matching pattern will be included.
'
'   work on the resizing. So much is going on during the resize It would look better to draw a
'   focus rectangle on the mouse move and only redraw the entire image when you mouse up
'
'Changes:
'
'update 02-12-08
'   Added the ability to size the thumbs.
'   and headed off a few erors.
'
'update 02-22-08
'   added resizability to the usercontrol. The code used can be used with any usercontrol, I think!
'   Those code blocks are seperated from the other code and commented start and end
'   I also worked out a few kinks in the code. and a general clean.
'
'  Major Fix: I was returning the Image index as "Index" in the mouse events
'               "Index" is used by VB to indicate Which Control in a control array
'               thereby creating a bad conflict. so I changed "Index" to "imgIndex" as seen below.
'
'     Fixed a error with the "path" Property Should be "Path" thanks to Roger Gilchrist
'
'     Fixed a few others errors the default settings wern't being set
'     and the property page needed updating to accomidate the different size thumbs
'
'   Fixed a problem when selecting the area under an image where the name is drawn. Your selection
'   would be made to the image bellow if there was one. I have now excluded the space between the images
'   and the space above the first row from the selection you must now select the image or just under
'   the image to selct that image.
'
'   Fixed a problem if you selected to the right of any image in the UC and the first image in the next row
'   would be selected
'
'   Added the ability to move the UC if it's not docked. drag it by a blank area on the control
'
'
'update 02-23-08
'   Added an option for the menu so you can use it or not. Also added the ability to hide the edit
'   items "Copy,Move,Delete and Rename" controled by the property AllowChanges"
'
'   Made the move and sizing optional through the "AllowSizing" property
'
'   Added a few shortcut keys
'   Ctrl + L Or T,R,B Will dock the control Left, Top, Right or Bottom
'   Ctrl + N will undock it if the parent isn't an MDI form
'
'   Shift And Arrow Keys will move or size the control depending on it's alignment
'   "AllowSizing" must be true
'
'   If the control has focus and a selection has been made you can scan through the
'   selections using the up and down arrow keys
'
'   The enter key fires the mouse down on the selected image
'
'   Fixed a few problems with the property page and finished adding property types
'
'   added keyboard events
'
'update 02-27-08 'requseted by Roger Gilchrist
'   added a maintain ratio property it will center the drawings in the thumbs
'   the text bellow will start at the center of the thumb.
'
'   added auto size to the tmbSize Property it will always size the thumb to the smallest
'   of the width or height of the control minus the border.
'
'
'update 03-06-08
'   added a "HotTracking" property select the image the mouse is over.
'
'
'update 03-16-08
'   Fixed a problem with the browse dialog opening the last folder you
'   selected or one of your choice. Thanks to Ed Wilk for pointing this out
'   and also for the link to the code used to fix the problem
'   http://vbnet.mvps.org/code/callback/browsecallback.htm
'
'   The usercontrol will now support unicode text in the image names but
'   the search text box doesn't.
'
'
'update 05-03-08
'   Fixed the control to use an array to store the images after creation.
'   this was to speed up drawing during resizing.
'
'
'update 05-29-08
'   removed most referances to the filelist and replaced them with properties of
'   there own. For example "selected" and "listIndex".
'
Private token As Long               ' Needed to close GDI+
Private a As Long                   'used as a counter
Private PartsWide As Long           'the number of thumbs wide our drawing surface is
Private ChangePath As Boolean       'have we chosen a new path since the last draw
Private UpdateImages As Boolean     'redraw all images
Private curSel As Integer           'the currently selected file
Private lstIndex As Integer         '0 based index of the selected file -1 indicates no selection
Private mFileType As String         'a string indicating the file extensions and patterns for searching
Private mtmbSize As Long            'one of five thumb size options
Private mMenu As Boolean            'used to block the menu display
Private mAllowChanges As Boolean    'used to block changing the files name or location
Private mAllowSizing As Boolean     'used to block the sizing of the user control
Private mcurDrawSize As Long        'holds the current size of the thumbs can be changed through the "tmbSize" property
Private mMaintainRatio As Boolean   'tries to keep the images at there origional ratio
Private mHotTracking As Boolean     'follow the mouse as it moves over the thumbs
Const m_def_HotTracking As Boolean = False
Const m_def_MaintainRatio As Boolean = False
Const m_def_curDrawSize As Long = 75
Const m_def_AllowSizing As Boolean = True
Const m_def_Menu As Boolean = True
Const m_def_tmbSize As Long = 1     'note: the larger the thumb the slower the drawing
Const m_def_AllowChanges As Boolean = True
Const m_def_FileType As String = "*.bmp;*.jpg;*.gif;*.jpeg;*.pcx;*.png;*.tga;*.tiff;*.ico;*.wbmp"

Event MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single, imgIndex As Integer)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single, imgIndex As Integer)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single, imgIndex As Integer)
Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event SelChanged()                  'when user selects a file in the list

Public Enum AppearanceConst
    Flat = 0
    Three_D = 1
End Enum

Public Enum BorderStyleConst
    None = 0
    Fixed = 1
End Enum

Public Enum thumbSZ                 'simple thumb size enum
    Sml = 0                         '50 * 50
    Med = 1                         '75 * 75
    Lrg = 2                         '125 * 125
    Cust = 3                        'set by user in property curDrawSize
    Auto = 4                        'ask for by user. when selected will resize the drawing to fit the smaller of the width or height of the control
End Enum

'used to determin the size and position of the text-----------------------------------------
Private Type POINTAPI
    x As Long
    Y As Long
End Type

Private TextSize As POINTAPI        'height and width of one char. based on the font used in the control
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As POINTAPI) As Long
'---------------------------------------------------------------------------
'Resizng start
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long 'used for moving the control
Private Declare Sub Sleep Lib "KERNEL32" (ByVal dwMilliseconds As Long) 'prevent over running the processor
Private Declare Function ReleaseCapture Lib "user32" () As Long 'used moving the control
Private Type sizer
    szRight As Boolean              'all sz var refer to mouse positions by way of true or false
    szLeft As Boolean               'say for example the mouse is at the top of the control
    szTop As Boolean                'szTop would be set to true. If it were also left
    szBot As Boolean                'szTopLeft would be true exc... exc... exc...
    szTopRight As Boolean
    szTopLeft As Boolean
    szBotRight As Boolean
    szBotLeft As Boolean
End Type
Private SZr As sizer                'used to access the above type
Private SplitCoordX As Single       'used to hold data about moveing and sizeing the control
Private SplitCoordY As Single
Private BMove As Boolean            'used to signal if we are sizing
Private cMove As Boolean            'used to signal if we are moving
Private intPnt As POINTAPI          'when moving this is the point where mouse down occured
'end resizing

'testing other methods to write text on the picture that allow for unicode
Private Type rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function DrawText Lib "user32" Alias "DrawTextW" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As rect, ByVal wFormat As Long) As Long
Private Const DT_LEFT = &H0
Private Const DT_SINGLELINE = &H20
Private Const DT_TOP = &H0
Private Const DT_WORDBREAK = &H10
Private Const DT_MODIFYSTRING = &H10000
Private Const DT_END_ELLIPSIS = &H8000&
Private Const DT_RTLREADING = &H20000
Private Const DT_CENTER = &H1
Private Const DT_NOPREFIX = &H800
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal length As Long)

Private Type imgData                'A type structure to hold the image and it's properties
    imgPic As Long                  'memory address to the image thumb
    imgWidth As Single              'width of the image
    imgHeight As Single             'height of the image
    imgPath As String               'Path to the image
    imgName As String               'image name without extension
    imgType As String               'image extension bmp, gif. exc...
End Type

Private graphics As Long            'acts sort of like an hdc in a pictureBox
Private images() As imgData         'used to store the thumbs when created this prevents reloading the images if the path isn't changed

Private Sub pDrawText(hdc As Long, ByVal sText As String, x1 As Long, y1 As Long, x2 As Long, y2 As Long)
Dim R As rect
    R.Left = x1: R.Top = y1: R.Right = x2: R.Bottom = y2
    'DrawAntiAliasText sText, CSng(x1), CSng(y1)
    DrawText hdc, StrConv(sText, vbUnicode), Len(sText), R, DT_CENTER Or DT_TOP Or DT_SINGLELINE Or DT_MODIFYSTRING Or DT_END_ELLIPSIS Or DT_RTLREADING
End Sub

Public Property Get MaintainRatio() As Boolean 'tries to maintain tha aspect ratio of the image if true
    MaintainRatio = mMaintainRatio
End Property
Public Property Let MaintainRatio(NewValue As Boolean)
    mMaintainRatio = NewValue
    PropertyChanged MaintainRatio
    UpdateImages = True
        If Loading = False Then Refresh
End Property

Public Property Get AllowSizing() As Boolean 'let the user allow or disallow moving and sizing
    AllowSizing = mAllowSizing
End Property
Public Property Let AllowSizing(NewValue As Boolean)
    mAllowSizing = NewValue
    PropertyChanged AllowSizing
End Property

Public Property Get AllowChanges() As Boolean 'let user allow or disallow menu items listed bellow in disable edit
Attribute AllowChanges.VB_ProcData.VB_Invoke_Property = "propThumbs"
    AllowChanges = mAllowChanges
End Property
Public Property Let AllowChanges(NewValue As Boolean)
    mAllowChanges = NewValue
    EnableEdit NewValue 'call a sub to enable or disable the edit menu items
    PropertyChanged AllowChanges
End Property

Private Sub EnableEdit(editOn As Boolean)
    Copy.Visible = editOn
    Move.Visible = editOn
    Rename.Visible = editOn
    Delete.Visible = editOn
    sp(0).Visible = editOn
End Sub

Public Property Get Menu() As Boolean 'let the user allow or disallow the whole menu
    Menu = mMenu
End Property
Public Property Let Menu(NewValue As Boolean)
    mMenu = NewValue
    PropertyChanged Menu
End Property

Public Property Get tmbSize() As thumbSZ
    tmbSize = mtmbSize
End Property

Public Property Let tmbSize(NewValue As thumbSZ)
    mtmbSize = NewValue
        Select Case NewValue
        'you can change these sizes here to control the small,medium, and large draw sizes
            Case 0: curDrawSize = 50            'small
            Case 1: curDrawSize = 75            'medium
            Case 2: curDrawSize = 125           'large
            Case 3: curDrawSize = curDrawSize   'custom
            Case 4: curDrawSize = IIf(Picture1.ScaleWidth <= Picture1.ScaleHeight, Picture1.ScaleWidth - 20, UserControl.ScaleHeight - (2 * TextSize.Y) - 18)
        End Select
End Property

Public Property Get curDrawSize() As Long
    curDrawSize = mcurDrawSize
End Property

Public Property Let curDrawSize(NewValue As Long)
        If mcurDrawSize = NewValue Then Exit Property 'if the size hasn't changed
    mcurDrawSize = NewValue
    PropertyChanged curDrawSize
    ChangePath = True
    UserControl_Resize 'resize the usercontrol
End Property

'the p2 pictureboxes are just for looks one contains the search textbox and the other is at the bottom and
'create the colored line at the bottom what we're doing here is just passing the mouse events through
'for our resizing code.
Private Sub p2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Select Case Index
        Case 0: UserControl_MouseDown Button, Shift, x, (UserControl.ScaleHeight - 3) + Y
        Case 1: UserControl_MouseDown Button, Shift, x, Y
    End Select
End Sub

Private Sub p2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Select Case Index
        Case 0: UserControl_MouseMove Button, Shift, x, (UserControl.ScaleHeight - 3) + Y
        Case 1: UserControl_MouseMove Button, Shift, x, Y
    End Select
End Sub

Private Sub sz_Click(Index As Integer) ' menu option to change the size of the thumb
Dim ret As String
    Select Case Index
        Case 0: tmbSize = 0
        Case 1: tmbSize = 1
        Case 2: tmbSize = 2
        Case 3: tmbSize = 3
            ret = InputBox("Enter a size for your thumbs.", "User Defined Thumb Size", curDrawSize)
                If Val(ret) > 0 Then
                    curDrawSize = Val(ret)
                End If
        Case 4: tmbSize = 4
    End Select
End Sub

Private Sub Browse_Click()
    BrowseFolder
End Sub

Public Sub BrowseFolder() ''call a sub in cDlg to browse for a folder
Dim NewFolder As String
    NewFolder = CD.BrowseForFolder("Select an Image Folder", , IIf(lstPath <> "", lstPath, App.path))
        If Not NewFolder = "" Then path = NewFolder: lstPath = NewFolder
End Sub

Private Sub Copy_Click() 'call a sub in cDlg to copy a file to another location(might should have called it CopyTo)
Dim DestPath As String
    If ListIndex > -1 Then
        DestPath = CD.BrowseForFolder("Select a folder to copy to.", , App.path)
            If DestPath <> "" Then
                DestPath = DestPath & "\" & images(ListIndex).imgName & "." & images(ListIndex).imgType
                CD.Copy path & "\" & images(ListIndex).imgName & "." & images(ListIndex).imgType, DestPath
            End If
    End If
End Sub

Private Sub Delete_Click() 'call a sub in cDlg to delete a file to the recycling bin
Dim a As Long
    If ListIndex > -1 Then
        CD.Delete path & "\" & images(ListIndex).imgName & "." & images(ListIndex).imgType, True
        Call GdipDisposeImage(images(ListIndex).imgPic)
            For a = ListIndex + 1 To UBound(images)
                images(a - 1) = images(a)
            Next a
        ReDim Preserve images(UBound(images) - 1)
        Selected(ListIndex) = False
        File1.Refresh
        lstIndex = -1
        UpdateImages = True
        Refresh
    End If
End Sub

Private Sub Move_Click() 'call a sub in cDlg to move a file
Dim DestPath As String
Dim a As Long
    If ListIndex > -1 Then
        DestPath = CD.BrowseForFolder("Select a folder to move to.", , App.path)
            If DestPath <> "" Then
                DestPath = DestPath & "\" & images(ListIndex).imgName & "." & images(ListIndex).imgType
                CD.Move path & "\" & images(ListIndex).imgName & "." & images(ListIndex).imgType, DestPath
                Call GdipDisposeImage(images(ListIndex).imgPic)
                    For a = ListIndex + 1 To UBound(images)
                        images(a - 1) = images(a)
                    Next a
                ReDim Preserve images(UBound(images) - 1)
                Selected(ListIndex) = False
                File1.Refresh
                lstIndex = -1
                UpdateImages = True
                Refresh
            End If
    End If
End Sub

Private Sub Properties_Click() 'call a sub in cDlg to show file properties
    If ListIndex > -1 Then
        CD.FileProperties path & "\" & images(ListIndex).imgName & "." & images(ListIndex).imgType
    End If
End Sub

Private Sub Rename_Click() 'call a sub in cDlg to rename a file
Dim DestPath As String
Dim DefExt As String
Dim NewName As String, ret As String
    If ListIndex > -1 Then
        DefExt = images(ListIndex).imgType
        NewName = InputBox("Enter a new name.", , images(ListIndex).imgName)
        DestPath = NewName
            If DestPath <> "" Then
                    If Right(DestPath, Len(DefExt)) <> DefExt Then DestPath = DestPath & "." & DefExt
                DestPath = images(ListIndex).imgPath & "\" & DestPath
                    If FileExists(DestPath, vbNormal Or vbArchive Or vbHidden Or vbReadOnly) = True Then
                        ret = MsgBox("This file exist. Overwrite it?", vbYesNoCancel)
                            If ret = vbNo Then
                                Rename_Click
                                Exit Sub
                            ElseIf ret = vbCancel Then
                                Exit Sub
                            End If
                    End If
                CD.Rename images(ListIndex).imgPath & "\" & images(ListIndex).imgName & "." & images(ListIndex).imgType, DestPath
                images(ListIndex).imgName = NewName
                Selected(ListIndex) = False
                File1.Refresh
                UpdateImages = True
                lstIndex = -1
                Refresh
            End If
    End If
End Sub

Private Sub p2_Resize(Index As Integer)
On Error Resume Next
    If Index = 1 Then srhTxt.Move 1, 2, p2(1).Width - 2, 14
End Sub

Private Sub srhTxt_Change() 'use text box to search file list
    ShowUpdated
End Sub

Private Sub srhTxt_DblClick() 'clear text box
    srhTxt.Text = ""
End Sub

Private Sub UserControl_Initialize()
Dim GpInput As GdiplusStartupInput
    Loading = True 'indicate not to update the drawing until all changes are made
    GpInput.GdiplusVersion = 1
        If GdiplusStartup(token, GpInput) <> Ok Then
           MsgBox "Error loading GDI+!" & vbCrLf & "You will need to download it from Microsoft", vbCritical
        End If
    Set CD = New cDlg
    CD.hOwner = UserControl.hWnd
End Sub

Private Sub UserControl_PaintImages()
Static lstPartCount As Long
On Error Resume Next
Dim lngHeight As Single, lngWidth As Single
Dim img As Long
Dim parts() As String
Dim yDif As Single, xDif As Single 'added to maintain the offset of width and height used to center the thumbs when maintain ratio is true
        If PartsWide = 0 Then Exit Sub 'this shouldn't happen
    'set the height of the Thumbs container
        If File1.ListCount Mod PartsWide = 0 Then
            Picture1.Height = (File1.ListCount \ PartsWide) * (mcurDrawSize + TextSize.Y) + TextSize.Y + 10
        Else
            Picture1.Height = (File1.ListCount \ PartsWide) * (mcurDrawSize + TextSize.Y) + (mcurDrawSize + TextSize.Y + 15)
        End If
        If Picture1.Height < UserControl.ScaleHeight - 20 Then Picture1.Height = UserControl.ScaleHeight - 20 'ensure it is at least as tall as the UC itself
            If lstPartCount <> PartsWide Or ChangePath = True Or UpdateImages = True Then 'if we havn't changed the width enough for another line or we haven't selected a new directory don't redraw.
                Cls
                'moved here to prevent recreating the graphics object with each file
                Call GdipCreateFromHDC(hdc, graphics)
                    If ChangePath = True Then
                        unloadImages
                        ReDim images(File1.ListCount - 1)
                            For a = 0 To File1.ListCount - 1
                                Call GdipLoadImageFromFile(StrConv(File1.path & "\" & File1.list(a), vbUnicode), img)   ' Load the image
                                Call GdipGetImageDimension(img, lngWidth, lngHeight) 'get width and height of the image
                                    If lngWidth > lngHeight Then 'determin if we are wider or taller
                                        xDif = 1                  'and build the ratios based on that
                                        yDif = lngWidth / lngHeight
                                    ElseIf lngHeight > lngWidth Then
                                        yDif = 1
                                        xDif = lngHeight / lngWidth
                                    Else
                                        yDif = 1
                                        xDif = 1
                                    End If
                                Call GdipGetImageThumbnail(img, curDrawSize / xDif, curDrawSize / yDif, images(a).imgPic) 'create a thumb of the image
                                images(a).imgHeight = curDrawSize / yDif
                                images(a).imgWidth = curDrawSize / xDif
                                parts = Split(File1.list(a), ".")
                                images(a).imgName = parts(0)
                                images(a).imgType = parts(1)
                                Erase parts
                                images(a).imgPath = File1.path
                                DrawThumbnail a + 1  'call function to draw the image
                                Call GdipDisposeImage(img) 'remove the image from memory
                            Next a
                    Else
                        For a = 0 To File1.ListCount - 1
                            DrawThumbnail a + 1
                        Next a
                    End If
                lstPartCount = PartsWide
                Call GdipDeleteGraphics(graphics) 'remove the graphics object from memory
                ChangePath = False
                UpdateImages = False
            End If
    VS.Max = Picture1.Height - (UserControl.ScaleHeight - 20)
    VS.LargeChange = IIf(Not UserControl.ScaleHeight \ 2 > VS.Max, UserControl.ScaleHeight \ 2, IIf(Not VS.Max <= 0, VS.Max, 1))
    VS.SmallChange = IIf(VS.LargeChange >= mcurDrawSize + TextSize.Y, mcurDrawSize + TextSize.Y, VS.LargeChange)  '\ 2
    Loading = False
End Sub

Private Sub UserControl_InitProperties()
    tmbSize = 1
    FileType = m_def_FileType
    path = App.path
    Set Font = Picture1.Font
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
'This is here to test the sizing when it's docked thought I would let it stay.
If Shift = 2 Then
    Select Case KeyCode
        Case 78: If Not TypeOf UserControl.Parent Is MDIForm Then UserControl.Extender.align = 0
        Case 76: UserControl.Extender.align = 3 'L
        Case 82: UserControl.Extender.align = 4 'R
        Case 84: UserControl.Extender.align = 1 'T
        Case 66: UserControl.Extender.align = 2 'B
    End Select
ElseIf Shift = 1 Then
    If AllowSizing = True Then
        Select Case KeyCode
            Case 37 'left
                If UserControl.Extender.align = 3 Then 'l
                    If UserControl.Extender.Width > curDrawSize Then
                        UserControl.Extender.Width = UserControl.Extender.Width - curDrawSize
                    End If
                ElseIf UserControl.Extender.align = 4 Then 'R
                    UserControl.Extender.Width = UserControl.Extender.Width + curDrawSize
                ElseIf UserControl.Extender.align = 0 Then
                    UserControl.Extender.Left = UserControl.Extender.Left - curDrawSize
                End If
            Case 38 'up
                If UserControl.Extender.align = 1 Then 't
                    If UserControl.Extender.Height > curDrawSize Then
                        UserControl.Extender.Height = UserControl.Extender.Height - curDrawSize
                    End If
                ElseIf UserControl.Extender.align = 2 Then 'b
                    UserControl.Extender.Height = UserControl.Extender.Height + curDrawSize
                ElseIf UserControl.Extender.align = 0 Then
                    UserControl.Extender.Top = UserControl.Extender.Top - curDrawSize
                End If
            Case 39 'Right
                If UserControl.Extender.align = 3 Then 'l
                    UserControl.Extender.Width = UserControl.Extender.Width + curDrawSize
                ElseIf UserControl.Extender.align = 4 Then 'R
                    If UserControl.Extender.Width > curDrawSize Then
                        UserControl.Extender.Width = UserControl.Extender.Width - curDrawSize
                    End If
                ElseIf UserControl.Extender.align = 0 Then
                    UserControl.Extender.Left = UserControl.Extender.Left + curDrawSize
                End If
            Case 40 'Down
                If UserControl.Extender.align = 1 Then 't
                    UserControl.Extender.Height = UserControl.Extender.Height + curDrawSize
                ElseIf UserControl.Extender.align = 2 Then 'b
                    If UserControl.Extender.Height > curDrawSize Then
                        UserControl.Extender.Height = UserControl.Extender.Height - curDrawSize
                    End If
                ElseIf UserControl.Extender.align = 0 Then
                    UserControl.Extender.Top = UserControl.Extender.Top + curDrawSize
                End If
        End Select
    End If
ElseIf Shift = 0 Then
Dim x As Single, Y As Single
    Select Case KeyCode
        Case 38 'up
            If ListIndex > 0 Then
                Selected(ListIndex - 1) = True
            End If
        Case 40 'down
            If ListIndex < ListCount - 1 Then
                Selected(ListIndex + 1) = True
            End If
        Case 13 'Return(Enter) key
            'fire the mouse down event for the selected thumb
            x = IIf(ListIndex Mod PartsWide = 0, 11, curDrawSize * (ListIndex Mod PartsWide))
            Y = curDrawSize * (ListIndex Mod PartsWide)
            RaiseEvent MouseDown(1, 0, x, Y, ListIndex)
    End Select
End If
RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Loading = True
    mHotTracking = PropBag.ReadProperty("HotTracking", m_def_HotTracking)
    Shape1(0).BorderColor = PropBag.ReadProperty("BorderColor", 255)
    mMaintainRatio = PropBag.ReadProperty("MaintainRatio", m_def_MaintainRatio)
    mAllowSizing = PropBag.ReadProperty("AllowSizing", m_def_AllowSizing)
    AllowChanges = PropBag.ReadProperty("AllowChanges", m_def_AllowChanges)
    mMenu = PropBag.ReadProperty("Menu", m_def_Menu)
    Shape1(0).BorderWidth = PropBag.ReadProperty("BorderWidth", 2)
    GetTextExtentPoint32 Picture1.hdc, "A", Len("A"), TextSize
    UserControl.Appearance = PropBag.ReadProperty("Appearance", 1)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    Picture1.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H80000001)
    p2(0).BackColor = PropBag.ReadProperty("BackColor", &H80000001)
    p2(1).BackColor = PropBag.ReadProperty("BackColor", &H80000001)
    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    tmbSize = PropBag.ReadProperty("tmbSize", 1)
    UserControl.AutoRedraw = PropBag.ReadProperty("AutoRedraw", True)
    File1.Hidden = PropBag.ReadProperty("Hidden", False)
    File1.ListIndex = PropBag.ReadProperty("ListIndex", 0)
    File1.Normal = PropBag.ReadProperty("Normal", True)
    File1.ReadOnly = PropBag.ReadProperty("ReadOnly", True)
    File1.System = PropBag.ReadProperty("System", False)
    File1.FileName = PropBag.ReadProperty("FileName", "")
    FileType = PropBag.ReadProperty("FileType", m_def_FileType)
    path = PropBag.ReadProperty("Path", App.path)
    lstPath = PropBag.ReadProperty("LastPath", App.path)
    Refresh
End Sub

Public Property Get FileType() As String 'use like a pattern property of the file list
Attribute FileType.VB_ProcData.VB_Invoke_Property = "propThumbs"
    FileType = mFileType
End Property

Public Property Let FileType(ByVal New_FileType As String)
    mFileType = New_FileType
    PropertyChanged "FileType"
    ShowUpdated
End Property

Private Sub ShowUpdated()
'this sub splits the FileType adds any text in the search box to each part
'and rejoins the parts and assigns that to the file list pattern property.
'when pattern is updated a sub is called to redraw images
Dim pt() As String
    If Len(srhTxt.Text) > 0 Then
        pt = Split(FileType, ";") 'moved this line inside the if block.
            For a = 0 To UBound(pt)
                pt(a) = srhTxt.Text & pt(a)
            Next a
        Pattern = Join(pt, ";")
        Erase pt
    Else
        Pattern = FileType
    End If
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
    curSel = ListIndex 'record any selections made
    PartsWide = ((UserControl.ScaleWidth - 19) \ (curDrawSize + 10)) 'determin how many thumbs wide
        If PartsWide < 1 Then PartsWide = 1 'always show one line of thumbs
    Picture1.Width = UserControl.ScaleWidth - 20 'set the drawing serface for the thumbs
    p2(0).Move 2, UserControl.ScaleHeight - 3, Picture1.Width, 3 'just for looks it adds a colored area at the bottom of the uc
    p2(1).Move 2, 0, Picture1.Width, 18 'just for looks it adds a colored area at the top of the uc
    VS.Move UserControl.ScaleWidth - 16, 2, 13, UserControl.ScaleHeight - 4 'position the scroll bar
        If tmbSize = 4 Then curDrawSize = IIf(Picture1.ScaleWidth <= UserControl.ScaleHeight - 18, Picture1.ScaleWidth - 20, UserControl.ScaleHeight - (2 * TextSize.Y) - 18) 'maintain the curDrawSize if we are set to auto size
        If curSel > -1 Then Selected(ListIndex) = False 'if a file is selected unselect it
        If Loading = False Then Refresh 'redraw the uc thumbs
        If curSel > -1 Then Selected(curSel) = True 'reselect any file we unselected
End Sub

Private Sub UserControl_Terminate()
    Set CD = Nothing
    unloadImages
    Call GdiplusShutdown(token)
End Sub
Public Sub unloadImages()
On Error GoTo oust 'I've been thinking this isn't a good idea. If an error does happen all images may not be unloaded
    For a = 0 To UBound(images)
        Call GdipDisposeImage(images(a).imgPic)
    Next a
oust: Err.Clear
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("HotTracking", mHotTracking, m_def_HotTracking)
    Call PropBag.WriteProperty("MaintainRatio", mMaintainRatio, m_def_MaintainRatio)
    Call PropBag.WriteProperty("AllowSizing", mAllowSizing, m_def_AllowSizing)
    Call PropBag.WriteProperty("AllowChanges", mAllowChanges, m_def_AllowChanges)
    Call PropBag.WriteProperty("Menu", mMenu, m_def_Menu)
    Call PropBag.WriteProperty("curDrawSize", mcurDrawSize, m_def_curDrawSize)
    Call PropBag.WriteProperty("tmbSize", mtmbSize, m_def_tmbSize)
    Call PropBag.WriteProperty("AutoRedraw", UserControl.AutoRedraw, False)
    
    Call PropBag.WriteProperty("Path", File1.path, "")
    Call PropBag.WriteProperty("FileType", mFileType, m_def_FileType)
    Call PropBag.WriteProperty("FileName", File1.FileName, "")
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H80000001)
    Call PropBag.WriteProperty("BorderColor", Shape1(0).BorderColor, 255)
    Call PropBag.WriteProperty("BorderWidth", Shape1(0).BorderWidth, 2)
    Call PropBag.WriteProperty("Font", Picture1.Font, Ambient.Font)
    Call PropBag.WriteProperty("Hidden", File1.Hidden, False)
    Call PropBag.WriteProperty("ListIndex", File1.ListIndex, 0)
    Call PropBag.WriteProperty("Normal", File1.Normal, True)
    Call PropBag.WriteProperty("ReadOnly", File1.ReadOnly, True)
    Call PropBag.WriteProperty("System", File1.System, False)
    Call PropBag.WriteProperty("Appearance", UserControl.Appearance, 1)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("ForeColor", Picture1.ForeColor, &H80000008)
    Call PropBag.WriteProperty("LastPath", lstPath, "")
End Sub

Private Sub DrawThumbnail(cntr As Long)
On Error Resume Next
Dim img As Long, imgThumb As Long
Dim lngHeight As Single, lngWidth As Single
Dim nme As String
Static l As Long 'left of current thumb These are set to static to allow addition of thumbs after the initial load
Static t As Long 'top of current thumb  basically to use as a list if I ever do that.
Dim yDif As Single, xDif As Single 'added to maintain the offset of width and height used to center the thumbs when maintain ratio is true
        If PartsWide > 1 Then
            If cntr = 1 Then 'determin the left and top of the next thumb
                t = TextSize.Y: l = 10 ' the first thumb in the list
            ElseIf cntr <= PartsWide Then 'within the first row
                l = (cntr - 1) * (curDrawSize + 10) + 10
            ElseIf cntr Mod PartsWide = 1 Then 'starting a new row
                t = (((cntr - 1) \ PartsWide) * (curDrawSize + TextSize.Y)) + TextSize.Y
                l = 10
            Else 'add the next item in the current row
                l = ((cntr - 1) Mod PartsWide) * (curDrawSize + 10) + 10
            End If
        Else 'just one line of thumbs only the top will change
            l = 10
            t = (cntr - 1) * (curDrawSize + TextSize.Y) + TextSize.Y
        End If
        If MaintainRatio = True Then 'using this slows us down some but not bad
            Call GdipDrawImageRectI(graphics, images(cntr - 1).imgPic, l + ((curDrawSize - (images(cntr - 1).imgWidth)) \ 2), t + ((curDrawSize - (images(cntr - 1).imgHeight)) \ 2), images(cntr - 1).imgWidth, images(cntr - 1).imgHeight)
        Else
            Call GdipDrawImageRectI(graphics, images(cntr - 1).imgPic, l, t, curDrawSize, curDrawSize)
        End If
    nme = images(cntr - 1).imgName
    pDrawText Picture1.hdc, nme, l, t + curDrawSize, l + curDrawSize, t + curDrawSize + TextSize.Y
    Picture1.Refresh
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim Index As Integer
    Index = GetIndex(x, Y) 'calls a sub to calculate our file list index
        If Index = -1 And UserControl.Extender.align = 0 Then 'you must hold your mouse between thumbs or on a blank area to move the UC
            cMove = True
            SetCapture Picture1.hWnd
            intPnt.x = x
            intPnt.Y = Y
        End If
    Selected(Index) = True 'calls a sub to select the file with this index the sub checks for an invalid index
        If Button = 2 Then 'right click
            If Not Index = -1 And Menu = True Then
                    If AllowChanges = True Then
                        EnableEdit True
                    End If
                Properties.Enabled = True
                PopupMenu Popup
            ElseIf Menu = True Then
                EnableEdit False
                Properties.Enabled = False
                PopupMenu Popup
            End If
        Else
            If Not Index = -1 Then RaiseEvent MouseDown(Button, Shift, x, Y, Index)
        End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim Index As Integer
    If cMove = True Then
        UserControl_MouseMove Button, Shift, x, Y
    Else
        Index = GetIndex(x, Y)
            If Not Index = -1 Then
                    If HotTracking = True And lstIndex <> Index Then
                        UserControl.SetFocus
                        Selected(Index) = True
                    End If
                RaiseEvent MouseMove(Button, Shift, x, Y, Index)
            End If
    End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim Index As Integer
        If cMove = True Then ReleaseCapture: UserControl_MouseUp Button, Shift, x, Y
    Index = GetIndex(x, Y)
    If Not Index = -1 Then RaiseEvent MouseUp(Button, Shift, x, Y, Index)
End Sub

Public Property Get HotTracking() As Boolean
    HotTracking = mHotTracking
End Property
Public Property Let HotTracking(NewValue As Boolean)
    mHotTracking = NewValue
    PropertyChanged HotTracking
End Property

Private Function GetIndex(x As Single, Y As Single) As Integer
'we can dtermin the index based on the position in the picture
Dim lin As Long 'the y
Dim row As Long 'the x
        If Y < TextSize.Y Or x < 10 Then GetIndex = -1: Exit Function 'we are above or to the left of any image
    lin = IIf(Y Mod (curDrawSize + TextSize.Y) = 0, Y \ (curDrawSize + TextSize.Y), IIf(Y Mod (curDrawSize + TextSize.Y) <= TextSize.Y, Y \ (curDrawSize + TextSize.Y), Y \ (curDrawSize + TextSize.Y) + 1))
    row = IIf(x Mod (curDrawSize + 10) = 0, x \ (curDrawSize + 10), IIf(x Mod (curDrawSize + 10) < 10, -1, x \ (curDrawSize + 10) + 1))
        If row > PartsWide Then GetIndex = -1: Exit Function 'to far right
    GetIndex = IIf(row <> -1, (((lin - 1) * PartsWide) + row) - 1, -1)
        If GetIndex > File1.ListCount - 1 Then GetIndex = -1 'if the index overruns the file listcount we are in a blank area
End Function

Private Sub VS_Change()
    Picture1.Top = -VS.value + 18
End Sub

Private Sub VS_GotFocus()
    Picture1.SetFocus 'keep the bar from blinking
End Sub

Private Sub VS_Scroll()
    Picture1.Top = -VS.value + 18
End Sub

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Attribute Enabled.VB_ProcData.VB_Invoke_Property = "propThumbs"
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Property Get hdc() As Long
Attribute hdc.VB_Description = "Returns a handle (from Microsoft Windows) to the object's device context."
    hdc = Picture1.hdc
End Property

Public Property Get AutoRedraw() As Boolean
Attribute AutoRedraw.VB_Description = "Returns/sets the output from a graphics method to a persistent bitmap."
Attribute AutoRedraw.VB_MemberFlags = "40"
    AutoRedraw = UserControl.AutoRedraw
End Property

Public Property Let AutoRedraw(ByVal New_AutoRedraw As Boolean)
    UserControl.AutoRedraw() = New_AutoRedraw
    PropertyChanged "AutoRedraw"
End Property

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    UserControl.Refresh
    UserControl_PaintImages
End Sub

Public Sub Cls()
Attribute Cls.VB_Description = "Clears graphics and text generated at run time from a Form, Image, or PictureBox."
    Picture1.Cls
End Sub

Public Property Get path() As String
Attribute path.VB_Description = "Returns/sets the current path."
Attribute path.VB_ProcData.VB_Invoke_Property = "propThumbs"
Attribute path.VB_MemberFlags = "200"
    path = File1.path
End Property

Public Property Let path(ByVal New_Path As String)
On Error Resume Next: Err.Clear
        If Not ListIndex = -1 Then Me.Selected(ListIndex) = False 'unselect any selected files
    File1.path() = New_Path
        If Err.Number = 76 Then
            path = App.path
            Err.Clear
            MsgBox "Path doesn't Exist"
        Else
            Err.Clear
        End If
    PropertyChanged "Path"
    File1.Refresh 'refreshing the file list unselects any previous selections
    lstIndex = -1 'setting our list index to match the list index of the file list
    ChangePath = True
        If Loading = False Then Refresh
End Property

Public Property Get Pattern() As String
Attribute Pattern.VB_Description = "Returns/sets a value indicating the filenames displayed in a control at run time."
Attribute Pattern.VB_MemberFlags = "40"
    Pattern = File1.Pattern
End Property

Public Property Let Pattern(ByVal New_Pattern As String)
    File1.Pattern() = New_Pattern
    PropertyChanged "Pattern"
    File1.Refresh 'refreshing the file list unselects any previous selections
    lstIndex = -1
    ChangePath = True
        If Loading = False Then Refresh
End Property

Public Property Get FileName() As String 'returns the name of the selected file
Attribute FileName.VB_Description = "Returns/sets the path and filename of a selected file."
    FileName = File1.FileName
End Property

Public Property Let FileName(ByVal New_FileName As String)
    File1.FileName() = New_FileName
    PropertyChanged "FileName"
End Property

Public Property Get ListCount() As Integer
Attribute ListCount.VB_Description = "Returns the number of items in the list portion of a control."
    ListCount = File1.ListCount
End Property

Public Property Get list(ByVal Index As Integer) As String 'returns the name of an item by index
Attribute list.VB_Description = "Returns/sets the items contained in a control's list portion."
    If Index > -1 Then
        list = images(Index).imgName & "." & images(Index).imgType
    Else
        list = ""
    End If
End Property

Public Property Get BorderColor() As OLE_COLOR
Attribute BorderColor.VB_Description = "Returns/sets the color of an object's border."
    BorderColor = Shape1(0).BorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
    Shape1(0).BorderColor() = New_BorderColor
    PropertyChanged "BorderColor"
End Property

Public Property Get BorderWidth() As Integer
Attribute BorderWidth.VB_Description = "Returns or sets the width of a control's border."
Attribute BorderWidth.VB_ProcData.VB_Invoke_Property = "propThumbs"
    BorderWidth = Shape1(0).BorderWidth
End Property

Public Property Let BorderWidth(ByVal New_BorderWidth As Integer)
    Shape1(0).BorderWidth() = New_BorderWidth
    PropertyChanged "BorderWidth"
End Property

Private Sub Picture1_DblClick()
    RaiseEvent DblClick
        If TypeOf UserControl.Parent Is MDIForm Then Exit Sub
        If UserControl.Extender.align < 4 Then
            UserControl.Extender.align = UserControl.Extender.align + 1
        Else
            UserControl.Extender.align = 0
        End If
End Sub

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = Picture1.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set Picture1.Font = New_Font
    Set srhTxt.Font = New_Font
    GetTextExtentPoint32 Picture1.hdc, "B", 1, TextSize 'this defines the text size between the thumbs
    PropertyChanged "Font"
    UpdateImages = True
    If Loading = False Then Refresh
End Property

Public Property Get ListIndex() As Integer
    ListIndex = lstIndex
End Property

Public Property Let ListIndex(ByVal New_ListIndex As Integer)
    lstIndex = New_ListIndex
    PropertyChanged "ListIndex"
End Property

Public Property Get MultiSelect() As Integer
    MultiSelect = File1.MultiSelect
End Property

Public Property Get Hidden() As Boolean
Attribute Hidden.VB_Description = "Determines whether a FileListBox control displays files with Hidden attributes."
Attribute Hidden.VB_ProcData.VB_Invoke_Property = "propThumbs"
    Hidden = File1.Hidden
End Property
Public Property Let Hidden(ByVal New_Hidden As Boolean)
    File1.Hidden() = New_Hidden
    PropertyChanged "Hidden"
End Property

Public Property Get Normal() As Boolean
Attribute Normal.VB_Description = "Determines whether a FileListBox control displays files with Normal attributes."
Attribute Normal.VB_ProcData.VB_Invoke_Property = "propThumbs"
    Normal = File1.Normal
End Property
Public Property Let Normal(ByVal New_Normal As Boolean)
    File1.Normal() = New_Normal
    PropertyChanged "Normal"
End Property

Public Property Get ReadOnly() As Boolean
Attribute ReadOnly.VB_Description = "Returns/sets a value that determines whether files with read-only attributes are displayed in the file list or not."
Attribute ReadOnly.VB_ProcData.VB_Invoke_Property = "propThumbs"
    ReadOnly = File1.ReadOnly
End Property
Public Property Let ReadOnly(ByVal New_ReadOnly As Boolean)
    File1.ReadOnly() = New_ReadOnly
    PropertyChanged "ReadOnly"
End Property

Public Property Get System() As Boolean
    System = File1.System
End Property
Public Property Let System(ByVal New_System As Boolean)
    File1.System() = New_System
    PropertyChanged "System"
End Property

Public Property Get Selected(ByVal Index As Integer) As Boolean
Attribute Selected.VB_Description = "Returns/sets the selection status of an item in a control."
    If Index < 0 Or Index > ListCount - 1 Then
        Selected = False
    Else
        Selected = IIf(lstIndex = Index, True, False)
    End If
End Property

Public Property Let Selected(ByVal Index As Integer, ByVal New_Selected As Boolean)
On Error Resume Next
    If Index <= -1 Then Exit Property 'head off errors
            If New_Selected = False Then 'if we're unselecting
                If Index > 0 Then 'we can unload all but the first control
                    If Not Shape1(Index) Is Nothing Then Unload Shape1(Index)
                Else
                    Shape1(Index).Visible = False 'so we simply hide the first control
                End If
            Else
                If Not lstIndex = Index Then setSelection Index 'calls a sub to place a new shape control over the selected image
            End If
        If Not lstIndex = Index Then lstIndex = Index: RaiseEvent SelChanged 'respond to the new selection
    PropertyChanged "Selected"
End Property

Private Sub setSelection(Index As Integer) 'matches the selection shape with the selected item in the file list box
On Error Resume Next
Dim lin As Long
Dim row As Long
If Index = -1 Then Exit Sub
        If Me.MultiSelect > 0 Then 'this isn't set up yet to handel multi file selections but this is here in preperation
            If Index <> 0 And Selected(Index) = False Then Load Shape1(Index)
        Else
                If ListIndex > 0 And Not Index = ListIndex Then
                    Unload Shape1(ListIndex) 'unload any other selection shape
                Else
                    If Not ListIndex = -1 Then Shape1(ListIndex).Visible = False 'hide the first one
                End If
            If Index > 0 And Selected(Index) = False Then Load Shape1(Index) 'load a new shape based on the index of the selected thumb
        End If
    'position and show the shape around the thumb
    lin = ((Index \ PartsWide) * (curDrawSize + TextSize.Y)) + TextSize.Y
    row = ((Index Mod PartsWide) * (curDrawSize + 10)) + 10
    Shape1(Index).Visible = True
    Shape1(Index).Move row, lin, curDrawSize, (curDrawSize + TextSize.Y)
    'Ensure we can always see the selected image
        If UserControl.ScaleHeight - (18 + (TextSize.Y * 2)) < Shape1(Index).Height Then
            VS.value = VS.value + Shape1(Index).Top + Shape1(Index).Height - (-Picture1.Top + (UserControl.ScaleHeight - 4))
                If Err Then VS.value = VS.Max
        ElseIf Shape1(Index).Top + Shape1(Index).Height > -Picture1.Top + (UserControl.ScaleHeight - 4) Then
            VS.value = VS.value + Shape1(Index).Top + Shape1(Index).Height - (-Picture1.Top + (UserControl.ScaleHeight - 4))
                If Err Then VS.value = VS.Max
        ElseIf Shape1(Index).Top < -Picture1.Top Then
            VS.value = Shape1(Index).Top - 5
                If VS.value < TextSize.Y Then VS.value = 0
        End If
End Sub

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    p2(0).BackColor() = New_BackColor
    p2(1).BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get Appearance() As AppearanceConst
Attribute Appearance.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D effects."
Attribute Appearance.VB_ProcData.VB_Invoke_Property = "propThumbs"
    Appearance = UserControl.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As AppearanceConst)
Dim BC As Long 'as you may well know the backcolor reverts to white when the appearance is changed
    BC = BackColor 'this was added to reset the color after the change
    UserControl.Appearance() = New_Appearance
    BackColor = BC
    PropertyChanged "Appearance"
End Property

Public Property Get BorderStyle() As BorderStyleConst
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
Attribute BorderStyle.VB_ProcData.VB_Invoke_Property = "propThumbs"
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleConst)
Dim BC As Long
    BC = BackColor
    UserControl.BorderStyle() = New_BorderStyle
    BackColor = BC
    PropertyChanged "BorderStyle"
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = Picture1.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    Picture1.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
    UpdateImages = True
    If Loading = False Then Refresh
End Property

'**************--------resizeing code--------*********************************************************************
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Not AllowSizing Then Exit Sub
    If Not BMove Then
    Dim a As Long
        BMove = True
        SetCapture UserControl.hWnd
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
'On Error Resume Next
    If Not AllowSizing Then Exit Sub
        If cMove Then
            UserControl.Extender.Left = UserControl.Extender.Left + -(intPnt.x - x)
            UserControl.Extender.Top = UserControl.Extender.Top + -(intPnt.Y - Y)
        ElseIf BMove Then 'resize the control
            If UserControl.Parent.ScaleMode = 1 Then x = x * Screen.TwipsPerPixelX: Y = Y * Screen.TwipsPerPixelY
            Select Case UserControl.Extender.align
            Case 0 'if we arn't docked to the parent form
                With UserControl.Extender
                    If SZr.szTopLeft = True Then
                        If Not .Width + -x <= 0 And Not .Height + -Y <= 0 Then .Move .Left - -x, .Top - -Y, .Width + -x, .Height + -Y
                    ElseIf SZr.szTopRight = True Then: If Not .Height + -Y = 0 And Not x <= 0 Then .Move .Left, .Top - -Y, x + (4 * Screen.TwipsPerPixelX), .Height + -Y
                    ElseIf SZr.szBotLeft = True Then: If Not .Width + -x <= 0 And Not Y <= 0 Then .Move .Left - -x, .Top, .Width + -x, Y + (4 * Screen.TwipsPerPixelY)
                    ElseIf SZr.szBotRight = True Then: If Not x <= 0 And Not Y <= 0 Then .Move .Left, .Top, x + (4 * Screen.TwipsPerPixelX), Y + (4 * Screen.TwipsPerPixelY)
                    ElseIf SZr.szTop = True Then: If Not .Height + -Y <= 0 Then .Move .Left, .Top - -Y, .Width, .Height + -Y
                    ElseIf SZr.szBot = True Then: If Not Y <= 0 Then .Height = Y + (4 * Screen.TwipsPerPixelY)
                    ElseIf SZr.szLeft = True Then: If Not .Width + -x <= 0 Then .Move .Left - -x, .Top, .Width + -x, .Height
                    ElseIf SZr.szRight = True Then: If Not x <= 0 Then .Width = x + (4 * Screen.TwipsPerPixelX)
                    End If
                End With
            Case 1 'If we are docked we only size one edge
                    If SplitCoordY = Y Then Exit Sub
                SplitCoordY = Y
                    UserControl.Height = SplitCoordY + (4 * Screen.TwipsPerPixelY)
            Case 2
                    If SplitCoordY = Y Then Exit Sub
                SplitCoordY = Y
                    UserControl.Height = UserControl.Height - SplitCoordY
            Case 3
                    If SplitCoordX = x Then Exit Sub
                SplitCoordX = x
                    If Not SplitCoordX <= 0 Then UserControl.Width = SplitCoordX + (4 * Screen.TwipsPerPixelX)
            Case 4
                    If SplitCoordX = x Then Exit Sub
                SplitCoordX = x
                    If SplitCoordX > UserControl.Parent.Width Then SplitCoordX = UserControl.Parent.Width - 500!
                UserControl.Width = UserControl.Width - SplitCoordX
            End Select
        Else 'if we arn't moving the control then moniter the mouse movements
        'along the edges of the control and
        'set up the mouse pointer based on the edge your on
    UserControl.SetFocus
    MousePointer = 0
        Select Case UserControl.Extender.align
            Case 0
                With UserControl
                    SZr.szTop = False
                    SZr.szBotRight = False
                    SZr.szBotLeft = False
                    SZr.szTopRight = False
                    SZr.szTopLeft = False
                    SZr.szBot = False
                    SZr.szRight = False
                    SZr.szLeft = False
                        If x < 4 Then
                                If x >= 0 Then
                                    SZr.szLeft = True
                                End If
                        ElseIf x < .ScaleWidth Then
                                If x > .ScaleWidth - 4 Then
                                    SZr.szRight = True
                                End If
                        End If
                        If Y < 4 Then
                                If Y >= 0 Then
                                    SZr.szTop = True
                                        If SZr.szLeft = True Then
                                            SZr.szTopLeft = True
                                        ElseIf SZr.szRight = True Then
                                            SZr.szTopRight = True
                                        End If
                                End If
                        ElseIf Y < .ScaleHeight Then
                                If Y > .ScaleHeight - 4 Then
                                    SZr.szBot = True
                                        If SZr.szLeft = True Then
                                            SZr.szBotLeft = True
                                        ElseIf SZr.szRight = True Then
                                            SZr.szBotRight = True
                                        End If
                                End If
                        End If
                End With
                'set the pointer based on the findings above
                If SZr.szRight = True Then
                    If SZr.szTopRight = True Then
                        MousePointer = 6 'nw/se
                    ElseIf SZr.szBotRight = True Then
                        MousePointer = 8 'ne/sw
                    Else
                        MousePointer = 9 'e/w
                    End If
                ElseIf SZr.szLeft = True Then
                    If SZr.szTopLeft = True Then
                        MousePointer = 8 'ne/sw
                    ElseIf SZr.szBotLeft = True Then
                        MousePointer = 6 'nw/se
                    Else
                        MousePointer = 9 'e/w
                    End If
                ElseIf SZr.szTop = True Then
                    MousePointer = 7  'n/s
                ElseIf SZr.szBot = True Then
                    MousePointer = 7  'n/s
                End If
            Case 1: If Y < UserControl.ScaleHeight And Y > UserControl.ScaleHeight - 4 Then MousePointer = 7 'n/s
            Case 2: If Y > -1 And Y < 4 Then MousePointer = 7 'n/s
            Case 3: If x < UserControl.ScaleWidth And x > UserControl.ScaleWidth - 4 Then MousePointer = 9 'e/w
            Case 4: If x > -1 And x < 4 Then MousePointer = 9 'e/w
            End Select
        End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Not AllowSizing Then Exit Sub
        If BMove Then
            ReleaseCapture
            BMove = False
        ElseIf cMove Then
            ReleaseCapture
            cMove = False
        End If
End Sub

'**************-----end resizeing code--------*********************************************************************

Private Sub DrawAntiAliasText(Text As String, x As Single, Y As Single)
   'Dim graphics As Long
   Dim brush As Long
   Dim fontFam As Long, curFont As Long
   Dim rcLayout As RECTF   ' Designates the string drawing bounds
   ' Initializations
   Call GdipCreateFromHDC(Picture1.hdc, graphics) ' Initialize the graphics class - required for all drawing
   Call GdipCreateSolidFill(Black, brush)     ' Create a brush to draw the text with
   ' Create a font family object to allow us to create a font
   ' We have no font collection here, so pass a NULL for that parameter
   Call GdipCreateFontFamilyFromName(StrConv(Picture1.Font.name, vbUnicode), 0, fontFam)
   ' Create the font from the specified font family name
   Call GdipCreateFont(fontFam, TextSize.Y - 2, FontStyleBold, UnitPixel, curFont)
   ' Set up a drawing area
   ' NOTE: Leaving the right and bottom values at zero means there is no boundary
   rcLayout.Left = x
   rcLayout.Top = Y
   Call GdipSetTextRenderingHint(graphics, TextRenderingHintAntiAlias)
   ' We have no string format object, so pass a NULL for that parameter
   Call GdipDrawString(graphics, StrConv(Text, vbUnicode), Len(Text), curFont, rcLayout, 0, brush)
   ' Cleanup
   Call GdipDeleteFont(curFont)     ' Delete the font object
   Call GdipDeleteFontFamily(fontFam)  ' Delete the font family object
   Call GdipDeleteBrush(brush)
   'Call GdipDeleteGraphics(graphics)
End Sub

Private Function RetFileName(path As String, Optional Ext As Boolean = True) As String
Dim Ary() As String
        If Right$(path, 1) = "\" Then path = Left(path, Len(path) - 1)
    Ary = Split(path, "\", , vbBinaryCompare)
        If Ext = False Then
            FileName = Left$(Ary(UBound(Ary)), Len(Ary(UBound(Ary))) - 4)
        Else
            FileName = Ary(UBound(Ary))
        End If
End Function

