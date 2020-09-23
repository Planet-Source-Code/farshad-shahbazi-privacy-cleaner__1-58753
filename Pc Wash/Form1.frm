VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pc Wash"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7125
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   7125
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog Co 
      Left            =   1560
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
      Flags           =   4
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Save log"
      Height          =   375
      Left            =   4080
      TabIndex        =   53
      Top             =   4725
      Width           =   855
   End
   Begin TabDlg.SSTab Tab1 
      Height          =   3615
      Left            =   240
      TabIndex        =   11
      Top             =   720
      Width           =   3885
      _ExtentX        =   6853
      _ExtentY        =   6376
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Quick"
      TabPicture(0)   =   "Form1.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Picture3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Browser"
      TabPicture(1)   =   "Form1.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Windows"
      TabPicture(2)   =   "Form1.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Picture1"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Other"
      TabPicture(3)   =   "Form1.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Picture4"
      Tab(3).ControlCount=   1
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3280
         Left            =   -74980
         ScaleHeight     =   3285
         ScaleWidth      =   3855
         TabIndex        =   25
         Top             =   320
         Width           =   3850
         Begin VB.CheckBox chRar 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Win Rar Recent Files"
            Height          =   255
            Left            =   240
            TabIndex        =   33
            Top             =   1680
            Width           =   1935
         End
         Begin VB.CheckBox chZip 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Win Zip Recent Files"
            Height          =   255
            Left            =   240
            TabIndex        =   32
            Top             =   1320
            Width           =   1935
         End
         Begin VB.CheckBox chVB 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Visual Basic Recent Files"
            Height          =   255
            Left            =   240
            TabIndex        =   31
            Top             =   2760
            Width           =   2175
         End
         Begin VB.CheckBox chCover 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Nero Cover Designer Recent Files"
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   2400
            Width           =   2775
         End
         Begin VB.CheckBox chWM 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Windows Media Player Recent Files"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   240
            Width           =   2895
         End
         Begin VB.CheckBox chPaint 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Paint Recent Files"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   240
            TabIndex        =   28
            Top             =   600
            Width           =   1935
         End
         Begin VB.CheckBox chWord 
            BackColor       =   &H00FFFFFF&
            Caption         =   "WordPad Recent Files"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   240
            TabIndex        =   27
            Top             =   960
            Width           =   2055
         End
         Begin VB.CheckBox chNero 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Nero Recent Files"
            Height          =   255
            Left            =   240
            TabIndex        =   26
            Top             =   2040
            Width           =   1695
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3280
         Left            =   20
         ScaleHeight     =   3285
         ScaleWidth      =   3855
         TabIndex        =   22
         Top             =   320
         Width           =   3850
         Begin VB.PictureBox Picture5 
            BackColor       =   &H00FFFFFF&
            Height          =   2940
            Left            =   240
            ScaleHeight     =   2880
            ScaleWidth      =   3075
            TabIndex        =   35
            Top             =   180
            Width           =   3135
            Begin VB.PictureBox picBoard 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   5895
               Left            =   0
               ScaleHeight     =   5895
               ScaleWidth      =   3135
               TabIndex        =   36
               Top             =   0
               Width           =   3135
               Begin VB.CheckBox ch5 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Clear Temp"
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Left            =   120
                  TabIndex        =   52
                  Top             =   1560
                  Width           =   1335
               End
               Begin VB.CheckBox ch6 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Clear Recent Documents"
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Left            =   120
                  TabIndex        =   51
                  Top             =   1920
                  Width           =   2295
               End
               Begin VB.CheckBox ch7 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Clear Typed Commands In Run"
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Left            =   120
                  TabIndex        =   50
                  Top             =   2280
                  Width           =   2535
               End
               Begin VB.CheckBox ch8 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Clear Searched Files"
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Left            =   120
                  TabIndex        =   49
                  Top             =   2640
                  Width           =   1935
               End
               Begin VB.CheckBox ch1 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Clear History"
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Left            =   120
                  TabIndex        =   48
                  Top             =   120
                  Width           =   1455
               End
               Begin VB.CheckBox ch2 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Clear Temporary Internet Files"
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Left            =   120
                  TabIndex        =   47
                  Top             =   480
                  Width           =   2535
               End
               Begin VB.CheckBox ch3 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Clear Typed Urls In IE"
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Left            =   120
                  TabIndex        =   46
                  Top             =   840
                  Width           =   2055
               End
               Begin VB.CheckBox ch4 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Clear Cookies"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   45
                  Top             =   1200
                  Width           =   1455
               End
               Begin VB.CheckBox ch13 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Win Rar Recent Files"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   44
                  Top             =   4440
                  Width           =   1935
               End
               Begin VB.CheckBox ch12 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Win Zip Recent Files"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   43
                  Top             =   4080
                  Width           =   1935
               End
               Begin VB.CheckBox ch16 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Visual Basic Recent Files"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   42
                  Top             =   5520
                  Width           =   2175
               End
               Begin VB.CheckBox ch15 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Nero Cover Designer Recent Files"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   41
                  Top             =   5160
                  Width           =   2775
               End
               Begin VB.CheckBox ch9 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Windows Media Player Recent Files"
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Left            =   120
                  TabIndex        =   40
                  Top             =   3000
                  Width           =   2895
               End
               Begin VB.CheckBox ch10 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Paint Recent Files"
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Left            =   120
                  TabIndex        =   39
                  Top             =   3360
                  Width           =   2535
               End
               Begin VB.CheckBox ch11 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "WordPad Recent Files"
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Left            =   120
                  TabIndex        =   38
                  Top             =   3720
                  Width           =   2055
               End
               Begin VB.CheckBox ch14 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Nero Recent Files"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   37
                  Top             =   4800
                  Width           =   1695
               End
            End
         End
         Begin VB.VScrollBar VS 
            Height          =   2940
            LargeChange     =   750
            Left            =   3360
            Max             =   2880
            SmallChange     =   375
            TabIndex        =   34
            Top             =   180
            Width           =   255
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3280
         Left            =   -74980
         ScaleHeight     =   3285
         ScaleWidth      =   3855
         TabIndex        =   17
         Top             =   320
         Width           =   3850
         Begin VB.ListBox List3 
            Appearance      =   0  'Flat
            Height          =   615
            Left            =   240
            TabIndex        =   24
            ToolTipText     =   "IE Cookies"
            Top             =   2520
            Width           =   3375
         End
         Begin VB.ListBox List2 
            Appearance      =   0  'Flat
            Height          =   615
            Left            =   240
            TabIndex        =   23
            ToolTipText     =   "IE Cache"
            Top             =   1680
            Width           =   3375
         End
         Begin VB.CheckBox chCook 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Clear Cookies"
            Height          =   255
            Left            =   240
            TabIndex        =   21
            Top             =   1320
            Width           =   1455
         End
         Begin VB.CheckBox chTyped 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Clear Typed Urls In IE"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   240
            TabIndex        =   20
            Top             =   960
            Width           =   2055
         End
         Begin VB.CheckBox chIETemp 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Clear Temporary Internet Files"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   240
            TabIndex        =   19
            Top             =   600
            Width           =   2535
         End
         Begin VB.CheckBox chHis 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Clear History"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3280
         Left            =   -74980
         ScaleHeight     =   3285
         ScaleWidth      =   3855
         TabIndex        =   12
         Top             =   320
         Width           =   3850
         Begin VB.CheckBox chSearched 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Clear Searched Files"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   240
            TabIndex        =   16
            Top             =   1320
            Width           =   1935
         End
         Begin VB.CheckBox chRun 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Clear Typed Commands In Run"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   960
            Width           =   2535
         End
         Begin VB.CheckBox chRec 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Clear Recent Documents"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   600
            Width           =   2295
         End
         Begin VB.CheckBox chTemp 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Clear Temp"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   240
            Width           =   1335
         End
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Select All"
      Height          =   375
      Left            =   5040
      TabIndex        =   1
      Top             =   4725
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4320
      Top             =   1560
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   2880
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.DirListBox Dir1 
      Height          =   315
      Left            =   2880
      TabIndex        =   6
      Top             =   480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   375
      Left            =   195
      TabIndex        =   3
      Top             =   4725
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "About"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   4725
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Do It"
      Height          =   375
      Left            =   6000
      TabIndex        =   0
      Top             =   4725
      Width           =   855
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   3540
      Left            =   4320
      TabIndex        =   8
      Top             =   840
      Width           =   2535
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00DBDBDB&
      BorderWidth     =   5
      X1              =   6840
      X2              =   6840
      Y1              =   840
      Y2              =   4320
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00DBDBDB&
      BorderWidth     =   3
      X1              =   6840
      X2              =   4320
      Y1              =   4360
      Y2              =   4360
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   6960
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   135
      X2              =   6960
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pc Wash"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   600
      Left            =   4680
      TabIndex        =   5
      Top             =   120
      Width           =   1800
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "What do you want to do?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   315
      TabIndex        =   4
      Top             =   240
      Width           =   2160
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pc Wash"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C4C8FF&
      Height          =   600
      Left            =   4680
      TabIndex        =   9
      Top             =   165
      Width           =   1800
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "What do you want to do?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   330
      TabIndex        =   10
      Top             =   270
      Width           =   2160
   End
   Begin VB.Image Image1 
      Height          =   5280
      Left            =   0
      Picture         =   "Form1.frx":037A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7125
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private t As String
Private CookCleaned As Boolean
Private IETempCleaned As Boolean
Private DIRForRM As Collection

Private Type INTERNET_CACHE_ENTRY_INFO
    dwStructSize As Long
    lpszSourceUrlName As Long
    lpszLocalFileName As Long
    CacheEntryType As Long
    dwUseCount As Long
    dwHitRate As Long
    dwSizeLow As Long
    dwSizeHigh As Long
    LastModifiedTime As FILETIME
    ExpireTime As FILETIME
    LastAccessTime As FILETIME
    LastSyncTime As FILETIME
    lpHeaderInfo As Long
    dwHeaderInfoSize As Long
    lpszFileExtension As Long
    dwReserved As Long
    dwExemptDelta As Long
    'szRestOfData() As Byte
End Type
Private Declare Function FindFirstUrlCacheEntry Lib "wininet.dll" Alias "FindFirstUrlCacheEntryA" (ByVal lpszUrlSearchPattern As String, ByVal lpFirstCacheEntryInfo As Long, ByRef lpdwFirstCacheEntryInfoBufferSize As Long) As Long
Private Declare Function FindNextUrlCacheEntry Lib "wininet.dll" Alias "FindNextUrlCacheEntryA" (ByVal hEnumHandle As Long, ByVal lpNextCacheEntryInfo As Long, ByRef lpdwNextCacheEntryInfoBufferSize As Long) As Long
Private Declare Sub FindCloseUrlCache Lib "wininet.dll" (ByVal hEnumHandle As Long)
Private Declare Function DeleteUrlCacheEntry Lib "wininet.dll" Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long


Private Sub LoadHistroy()
    Dim ICEI As INTERNET_CACHE_ENTRY_INFO, ret As Long
    Dim hEntry As Long, Msg As VbMsgBoxResult
    Dim MemBlock As New MemoryBlock
    
    List1.Clear
    'Start enumerating the visited URLs
    FindFirstUrlCacheEntry vbNullString, ByVal 0&, ret
    'If Ret is larger than 0...
    If ret > 0 Then
        '... allocate a buffer
        MemBlock.Allocate ret
        'call FindFirstUrlCacheEntry
        hEntry = FindFirstUrlCacheEntry(vbNullString, MemBlock.Handle, ret)
        'copy from the buffer to the INTERNET_CACHE_ENTRY_INFO structure
        MemBlock.ReadFrom VarPtr(ICEI), LenB(ICEI)
        'Add the lpszSourceUrlName string to the listbox
        If ICEI.lpszSourceUrlName <> 0 Then
         If Left(MemBlock.ExtractString(ICEI.lpszSourceUrlName, ret), 6) = "Cookie" Then
         List3.AddItem MemBlock.ExtractString(ICEI.lpszSourceUrlName, ret)
         Else
         List2.AddItem MemBlock.ExtractString(ICEI.lpszSourceUrlName, ret)
         End If
        End If
    End If
    'Loop until there are no more items
    Do While hEntry <> 0
        'Initialize Ret
        ret = 0
        'Find out the required size for the next item
        FindNextUrlCacheEntry hEntry, ByVal 0&, ret
        'If we need to allocate a buffer...
        If ret > 0 Then
            '... do it
            MemBlock.Allocate ret
            'and retrieve the next item
            FindNextUrlCacheEntry hEntry, MemBlock.Handle, ret
            'copy from the buffer to the INTERNET_CACHE_ENTRY_INFO structure
            MemBlock.ReadFrom VarPtr(ICEI), LenB(ICEI)
            'Add the lpszSourceUrlName string to the listbox
            If ICEI.lpszSourceUrlName <> 0 Then
             If Left(MemBlock.ExtractString(ICEI.lpszSourceUrlName, ret), 6) = "Cookie" Then
             List3.AddItem MemBlock.ExtractString(ICEI.lpszSourceUrlName, ret)
             Else
             List2.AddItem MemBlock.ExtractString(ICEI.lpszSourceUrlName, ret)
             End If
            End If
        'Else = no more items
        Else
            Exit Do
        End If
    Loop
    'Close enumeration handle
    FindCloseUrlCache hEntry
    'Delete our memory block
    Set MemBlock = Nothing
End Sub

Private Function DeleteCacheEntry(SourceUrl As String) As Boolean
    Dim lReturnValue As Long
    
    lReturnValue = DeleteUrlCacheEntry(SourceUrl)
    DeleteCacheEntry = CBool(lReturnValue)
    
End Function


Function ClearD(strPath)
On Error Resume Next
Set DIRForRM = New Collection

If Right(strPath, 1) <> "\" Then
strPath = strPath & "\"
End If

SetAttr strPath, vbArchive

Dir1.Path = strPath
File1.Path = strPath

For i = 0 To File1.ListCount - 1
SetAttr File1.Path & "\" & File1.List(i), vbArchive
Kill File1.Path & "\" & File1.List(i)
 If Dir(File1.Path & "\" & File1.List(i), vbArchive) = "" Then
 List1.AddItem "-->" & File1.Path & "\" & File1.List(i)
 List1.ListIndex = List1.ListCount - 1
 End If
Next
File1.Refresh


a = Dir1.Path
For i = 0 To Dir1.ListCount - 1
Dir1.Path = a
ClearD Dir1.List(i)
SetAttr Dir1.List(i), vbArchive
Dir1.Path = a
DIRForRM.Add Dir1.List(i)
Next
Dir1.Refresh

For i = 1 To DIRForRM.Count
RmDir DIRForRM(i)
Next

End Function

Function Recent()
Dim RecFolder As String
RecFolder = RGGetKeyValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Recent")



ClearD RecFolder

For i = Asc("a") To Asc("z")
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\RecentDocs\", Chr(i)
Next
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\RecentDocs\", "MRUList"
End Function


Function Temp()
Dim TempFolder As String
Dim strFolder As String * 255
Dim intLenght As Integer
intLenght = GetTempPath(255, strFolder)
TempFolder = Left(strFolder, intLenght)


ClearD TempFolder
End Function

Function IETemp()
Dim IEFolder As String

If IETempCleaned = False Then
 For ret = 0 To List2.ListCount - 1
  DeleteCacheEntry List2.List(ret)
  List1.AddItem "-->" & List2.List(ret)
  List1.ListIndex = List1.ListCount - 1
  List2.RemoveItem ret
  List2.AddItem "Removed", ret
 Next ret
IETempCleaned = True
End If

IEFolder = RGGetKeyValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Cache")
ClearD IEFolder

'ClearD "C:\WINDOWS\Local Settings"


End Function

Function Run()
For i = Asc("a") To Asc("j")
se = RGGetKeyValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\RunMRU", Chr(i))
If se <> "" Then List1.AddItem "-->" & se: List1.ListIndex = List1.ListCount - 1
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\RunMRU", Chr(i)
Next
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\RunMRU", "MRUList"
End Function

Function VBR()
For i = 1 To 100
se = RGGetKeyValue(HKEY_CURRENT_USER, "Software\Microsoft\Visual Basic\6.0\RecentFiles", "" & i)
If se <> "" Then List1.AddItem "-->" & se: List1.ListIndex = List1.ListCount - 1
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Visual Basic\6.0\RecentFiles", "" & i

se = RGGetKeyValue(HKEY_CURRENT_USER, "Software\Microsoft\Visual Basic\5.0\RecentFiles", "" & i)
If se <> "" Then List1.AddItem "-->" & se: List1.ListIndex = List1.ListCount - 1
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Visual Basic\5.0\RecentFiles", "" & i

Next
End Function

Function Cover()
For i = 1 To 100
se = RGGetKeyValue(HKEY_CURRENT_USER, "Software\Ahead\Cover Designer\Recent File List", "file" & i)
If se <> "" Then List1.AddItem "-->" & se: List1.ListIndex = List1.ListCount - 1
DeleteValue HKEY_CURRENT_USER, "Software\Ahead\Cover Designer\Recent File List", "file" & i
Next
End Function

Function nero()
For i = 1 To 100
se = RGGetKeyValue(HKEY_CURRENT_USER, "Software\Ahead\Nero - Burning Rom\Recent File List", "file" & i)
If se <> "" Then List1.AddItem "-->" & se: List1.ListIndex = List1.ListCount - 1
DeleteValue HKEY_CURRENT_USER, "Software\Ahead\Nero - Burning Rom\Recent File List", "file" & i
Next
End Function

Function Rar()
For i = 0 To 100
se = RGGetKeyValue(HKEY_CURRENT_USER, "Software\WinRAR\ArcHistory", "" & i)
If se <> "" Then List1.AddItem "-->" & se: List1.ListIndex = List1.ListCount - 1
DeleteValue HKEY_CURRENT_USER, "Software\WinRAR\ArcHistory", "" & i

se = RGGetKeyValue(HKEY_CURRENT_USER, "Software\WinRAR\DialogEditHistory\ArcName", "" & i)
If se <> "" Then List1.AddItem "-->" & se: List1.ListIndex = List1.ListCount - 1
DeleteValue HKEY_CURRENT_USER, "Software\WinRAR\DialogEditHistory\ArcName", "" & i

se = RGGetKeyValue(HKEY_CURRENT_USER, "Software\WinRAR\DialogEditHistory\ExtrPath", "" & i)
If se <> "" Then List1.AddItem "-->" & se: List1.ListIndex = List1.ListCount - 1
DeleteValue HKEY_CURRENT_USER, "Software\WinRAR\DialogEditHistory\ExtrPath", "" & i

Next
End Function

Function Zip()
For i = 1 To 100
se = RGGetKeyValue(HKEY_CURRENT_USER, "Software\Nico Mak Computing\WinZip\filemenu", "filemenu" & i)
If se <> "" Then List1.AddItem "-->" & se: List1.ListIndex = List1.ListCount - 1
DeleteValue HKEY_CURRENT_USER, "Software\Nico Mak Computing\WinZip\filemenu", "filemenu" & i
Next
End Function

Function Word()
For i = 1 To 100
se = RGGetKeyValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Applets\WordPad 2.0\Recent File List", "File" & i)
If se <> "" Then List1.AddItem "-->" & se: List1.ListIndex = List1.ListCount - 1
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Applets\WordPad 2.0\Recent File List", "File" & i
Next
End Function

Function Paint()
For i = 1 To 100
se = RGGetKeyValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Applets\Paint\Recent File List", "File" & i)
If se <> "" Then List1.AddItem "-->" & se: List1.ListIndex = List1.ListCount - 1
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Applets\Paint\Recent File List", "File" & i
Next
End Function

Function WM()
For i = 0 To 100
se = RGGetKeyValue(HKEY_CURRENT_USER, "Software\Microsoft\MediaPlayer\Player\RecentURLList", "URL" & i)
If se <> "" Then List1.AddItem "-->" & se: List1.ListIndex = List1.ListCount - 1
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\MediaPlayer\Player\RecentURLList", "URL" & i
Next
End Function

Function Typed()
For i = 1 To 100
se = RGGetKeyValue(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\TypedURLs\", "url" & i)
If se <> "" Then List1.AddItem "-->" & se: List1.ListIndex = List1.ListCount - 1
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\TypedURLs\", "url" & i
Next
End Function

Function Searched()
For i = Asc("a") To Asc("j")
se = RGGetKeyValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Doc Find Spec MRU\", Chr(i))
If se <> "" Then List1.AddItem "-->" & se: List1.ListIndex = List1.ListCount - 1
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Doc Find Spec MRU\", Chr(i)
Next
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Doc Find Spec MRU\", "MRUList"
End Function

Function Cookies()
Dim CoFolder As String

If CookCleaned = False Then
 For ret = 0 To List3.ListCount - 1
  DeleteCacheEntry List3.List(ret)
  List1.AddItem "-->" & List3.List(ret)
  List1.ListIndex = List1.ListCount - 1
  List3.RemoveItem ret
  List3.AddItem "Removed", ret
 Next ret
 CookCleaned = True
End If

CoFolder = RGGetKeyValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Cookies")
ClearD CoFolder
End Function

Function HisTory()
On Error Resume Next
Dim HisFolder As String

Dim a As UrlHistory
Set a = New UrlHistory
a.ClearHistory

HisFolder = RGGetKeyValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "History")
ClearD HisFolder

End Function

Private Sub Command1_Click()
Me.MousePointer = vbHourglass
DoEvents
List1.Clear
If chHis.Value = 1 Or ch1.Value = 1 Then List1.AddItem "->Deleting History": List1.ListIndex = List1.ListCount - 1: HisTory
If chIETemp.Value = 1 Or ch2.Value = 1 Then List1.AddItem "->Deleting IE Cache": List1.ListIndex = List1.ListCount - 1: IETemp
If chTyped.Value = 1 Or ch3.Value = 1 Then List1.AddItem "->Deleting Typed urls": List1.ListIndex = List1.ListCount - 1: Typed
If chCook.Value = 1 Or ch4.Value = 1 Then List1.AddItem "->Deleting Cookies": List1.ListIndex = List1.ListCount - 1: Cookies
If chTemp.Value = 1 Or ch5.Value = 1 Then List1.AddItem "->Deleting Temp Files": List1.ListIndex = List1.ListCount - 1: Temp
If chRec.Value = vbChecked Or ch6.Value = 1 Then List1.AddItem "->Deleting Recent Documents": List1.ListIndex = List1.ListCount - 1: Recent
If chRun.Value = 1 Or ch7.Value = 1 Then List1.AddItem "->Deleting Typed Commands In Run": List1.ListIndex = List1.ListCount - 1: Run
If chSearched.Value = 1 Or ch8.Value = 1 Then List1.AddItem "->Deleting Searched Files": List1.ListIndex = List1.ListCount - 1: Searched
If chWM.Value = 1 Or ch9.Value = 1 Then List1.AddItem "->Deleting WM Recent Files": List1.ListIndex = List1.ListCount - 1: WM
If chPaint.Value = 1 Or ch10.Value = 1 Then List1.AddItem "->Deleting Paint Recent Files": List1.ListIndex = List1.ListCount - 1: Paint
If chWord.Value = 1 Or ch11.Value = 1 Then List1.AddItem "->Deleting Word Pad Recent Files": List1.ListIndex = List1.ListCount - 1: Word
If chZip.Value = 1 Or ch12.Value = 1 Then List1.AddItem "->Deleting Win Zip Recent Files": List1.ListIndex = List1.ListCount - 1: Zip
If chRar.Value = 1 Or ch13.Value = 1 Then List1.AddItem "->Deleting Win Rar Recent Files": List1.ListIndex = List1.ListCount - 1: Rar
If chNero.Value = 1 Or ch14.Value = 1 Then List1.AddItem "->Deleting Nero Recent Files": List1.ListIndex = List1.ListCount - 1: nero
If chCover.Value = 1 Or ch15.Value = 1 Then List1.AddItem "->Deleting Nero Cover Designer Recent Files": List1.ListIndex = List1.ListCount - 1: Cover
If chVB.Value = 1 Or ch16.Value = 1 Then List1.AddItem "->Deleting Visual Basic Recent Files": List1.ListIndex = List1.ListCount - 1: VBR

List1.AddItem "->Completed<-": List1.ListIndex = List1.ListCount - 1:
Me.MousePointer = vbNormal
End Sub

Private Sub Command2_Click()
List1.Clear
t = 0
Timer1.Enabled = True
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Command4_Click()
Select Case Tab1.Caption
Case "Quick"
 ch1.Value = 1
 ch2.Value = 1
 ch3.Value = 1
 ch4.Value = 1
 ch5.Value = 1
 ch6.Value = 1
 ch7.Value = 1
 ch8.Value = 1
 ch9.Value = 1
 ch10.Value = 1
 ch11.Value = 1
 ch12.Value = 1
 ch13.Value = 1
 ch14.Value = 1
 ch15.Value = 1
 ch16.Value = 1
Case "Browser"
 chHis.Value = 1
 chIETemp.Value = 1
 chTyped.Value = 1
 chCook.Value = 1
Case "Windows"
 chTemp.Value = 1
 chRec.Value = 1
 chRun.Value = 1
 chSearched.Value = 1
Case "Other"
 chWM.Value = 1
 chPaint.Value = 1
 chWord.Value = 1
 chZip.Value = 1
 chRar.Value = 1
 chNero.Value = 1
 chCover.Value = 1
 chVB.Value = 1
End Select
 
End Sub

Private Sub Command5_Click()
On Error GoTo hell
Co.ShowSave
strlog = "PC Wash   Version 1.20" & vbCrLf & vbCrLf
For i = 0 To List1.ListCount - 1
strlog = strlog & List1.List(i) & vbCrLf
Next
Open Co.FileName For Output As #1
Print #1, strlog
Close #1
MsgBox "Log Saved.", vbInformation
hell:
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Form_Load()
LoadHistroy
End Sub

Private Sub Tab1_Click(PreviousTab As Integer)
 ch1.Value = 0
 ch2.Value = 0
 ch3.Value = 0
 ch4.Value = 0
 ch5.Value = 0
 ch6.Value = 0
 ch7.Value = 0
 ch8.Value = 0
 ch9.Value = 0
 ch10.Value = 0
 ch11.Value = 0
 ch12.Value = 0
 ch13.Value = 0
 ch14.Value = 0
 ch15.Value = 0
 ch16.Value = 0
 chHis.Value = 0
 chIETemp.Value = 0
 chTyped.Value = 0
 chCook.Value = 0
 chTemp.Value = 0
 chRec.Value = 0
 chRun.Value = 0
 chSearched.Value = 0
 chWM.Value = 0
 chPaint.Value = 0
 chWord.Value = 0
 chZip.Value = 0
 chRar.Value = 0
 chNero.Value = 0
 chCover.Value = 0
 chVB.Value = 0
End Sub

Private Sub Timer1_Timer()
If t = 0 Then
List1.AddItem "                  Pc Wash"
t = t + 1
Exit Sub
End If

If t = 1 Then
List1.AddItem "                Version 1.20"
t = t + 1
Exit Sub
End If

If t = 2 Then
List1.AddItem "         By Farshad Shahbazi"
t = t + 1
Exit Sub
End If

If t = 3 Then
List1.AddItem "       Shahbazi66@yahoo.com"
Timer1.Enabled = False
End If
End Sub

Private Sub VS_Change()
picBoard.Top = -VS.Value
End Sub

Private Sub VS_Scroll()
picBoard.Top = -VS.Value
End Sub
