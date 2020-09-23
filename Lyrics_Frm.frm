VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "Msinet.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Lyrics_Frm 
   BackColor       =   &H00523939&
   Caption         =   "Lyric Finder"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9870
   LinkTopic       =   "Form1"
   ScaleHeight     =   8160
   ScaleWidth      =   9870
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3735
      Top             =   2025
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   3705
      Top             =   1275
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Frame Alpha_Frame 
      Appearance      =   0  'Flat
      BackColor       =   &H00523939&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   105
      TabIndex        =   8
      Top             =   525
      Width           =   9630
      Begin VB.OptionButton Alpha_Opt 
         Caption         =   "1-9"
         Height          =   420
         Index           =   0
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         Width           =   450
      End
   End
   Begin VB.CommandButton SearchA_but 
      Caption         =   "Search"
      Height          =   315
      Left            =   6885
      TabIndex        =   7
      Top             =   120
      Width           =   1050
   End
   Begin VB.TextBox txtLyrics 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      ForeColor       =   &H0000C000&
      Height          =   7050
      Left            =   4005
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   6
      Top             =   990
      Width           =   5730
   End
   Begin MSComctlLib.ListView RefList_Lvw 
      Height          =   4080
      Left            =   105
      TabIndex        =   4
      Top             =   990
      Width           =   3825
      _ExtentX        =   6747
      _ExtentY        =   7197
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   49152
      BackColor       =   0
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Artist"
         Object.Width           =   6174
      EndProperty
   End
   Begin VB.TextBox Song_Tbox 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H0000C000&
      Height          =   300
      Left            =   4065
      TabIndex        =   2
      Text            =   "Up"
      Top             =   120
      Width           =   2730
   End
   Begin VB.TextBox Artist_Tbox 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H0000C000&
      Height          =   315
      Left            =   660
      TabIndex        =   0
      Text            =   "Shania Twain"
      Top             =   105
      Width           =   2730
   End
   Begin MSComctlLib.ListView Songs_Lvw 
      Height          =   2895
      Left            =   105
      TabIndex        =   5
      Top             =   5130
      Width           =   3825
      _ExtentX        =   6747
      _ExtentY        =   5106
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   49152
      BackColor       =   0
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Songs"
         Object.Width           =   6174
      EndProperty
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00523939&
      Caption         =   "Song"
      ForeColor       =   &H00C0C000&
      Height          =   195
      Left            =   3510
      TabIndex        =   3
      Top             =   165
      Width           =   435
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00523939&
      Caption         =   "Artist"
      ForeColor       =   &H00C0C000&
      Height          =   240
      Left            =   135
      TabIndex        =   1
      Top             =   150
      Width           =   390
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnusavelyrics 
         Caption         =   "&Save Lyrics"
      End
      Begin VB.Menu mnuprintlyrics 
         Caption         =   "&Print Lyrics"
      End
      Begin VB.Menu mnusepa 
         Caption         =   "-"
      End
      Begin VB.Menu mnuabout 
         Caption         =   "About"
      End
      Begin VB.Menu mnusepb 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "Lyrics_Frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Original idea came from an app
'kind of like this submitted by James Balducci (2002)
'so thanks to him for this idea.

'any problems, bitching or even nice things
'can be directed to steve@finishingplus.com
'Have A Good Time Steve Grimes aka KingFisH.


'KNOWN BUGS:
' Some artists are listed under the main site but
' no songs actually exist in the main site you
' using explorer you would get redirected to a new
' Web site. I completly ignore it and just return
' No Songs Found. A Good example is Neil Diamond.


Dim sWebSiteRoot As String
Dim bIgnoreClick As Boolean
Private Sub Form_Load()
Dim iLoop As Integer

'Im Starting From This Web Site
'But You Can Start From A Number Of Them And It
'Should Work Because The Program Just Follows The 'href=[somesite.com]/SomeArtist/SomeSong.Html'
'Reference As We Narrow Down To The Lyrics
sWebSiteRoot = "http://www.azlyrics.com/"

'Add All The Alpha Directory Buttons
Alpha_Opt(0).Width = Alpha_Frame.Width / 27
For iLoop = 1 To 26
    Load Alpha_Opt(iLoop)
    With Alpha_Opt(iLoop)
        .Left = Alpha_Opt(iLoop - 1).Left + Alpha_Opt(iLoop - 1).Width
        .Visible = True
        .Caption = Chr(iLoop + 64)
    End With
Next 'iLoop
    
End Sub
Private Function GetArtistList_Alpha(sAlpha As String, Optional sFilter As String = "") As Boolean
Dim sTab As String
Dim sPage As String
Dim RefArr
Dim iLoop As Integer
Dim lItm As ListItem
Dim iFltrLoop As Integer
Dim iMatch As Integer
Dim arrFilter

'Click On A Letter And Get The Ball Rolling

RefList_Lvw.ListItems.Clear
Songs_Lvw.ListItems.Clear
txtLyrics.Text = ""


sTab = LCase(sAlpha)
If sTab = "1-9" Then sTab = "19" 'If Other Starting Site, You May Want To Verify
                                 'What ##.Html They Use For Non Alpha Artists
                                 


sPage = GetPage(sWebSiteRoot & sTab & ".html")

If sPage <> "" Then 'OK DOKE
    If FindAndStripBefore(sPage, "</form>") = True Then  'Ok Good
        
        RefArr = FindHREFS(sPage, "<A href=", "</a><BR>", , sFilter)
        If UBound(RefArr, 2) > 0 Then
            For iLoop = 1 To UBound(RefArr, 2)
                GetArtistList_Alpha = True
                Set lItm = RefList_Lvw.ListItems.Add(, , RefArr(1, iLoop))
                lItm.Tag = RefArr(0, iLoop)
                lItm.ToolTipText = RefArr(0, iLoop)
            Next 'iLoop
        Else
           ' Msgbox "No Reference Found <href>"
        End If
    
    Else
       ' Msgbox "Cannot Strip Page, No " & "</Table>"
    End If
Else
   ' Msgbox "Internet Access My Not Be Available"
End If


End Function
Private Sub mnuabout_Click()

MsgBox "Contact steve@finishingplus.com With Problems, Bugs Or Kudos" & vbCrLf & vbCrLf & "You May Use This Software Any Way You See Fit. Have Fun. Just Give Credit Where Credit Is Due", , "FisHFindeR 1.2 by KingFisH"


End Sub
Private Sub mnuexit_Click()

Unload Me

End Sub
Private Sub mnuprintlyrics_Click()
Dim sBuff As String
Dim arrToCenter
Dim lArrLoop As Long

'I Liked The Lyrics Centered On The Page So I Went Through All
'This To Get It To Look Like That. You Could Very Well Use The
'Following To Keep It Simple:
'   Printer.Print txtLyrics.Text
'   Printer.EndDoc

If txtLyrics.Text <> "" Then
    sBuff = RefList_Lvw.SelectedItem.Text & vbCrLf & vbCrLf & Songs_Lvw.SelectedItem.Text & vbCrLf & vbCrLf & vbCrLf & txtLyrics.Text
    arrToCenter = Split(sBuff, vbCrLf)
    Printer.CurrentY = 720 '1/2 Inch Top Margin
    For lArrLoop = 0 To UBound(arrToCenter)
        Printer.CurrentX = (Printer.Width / 2) - (Printer.TextWidth(arrToCenter(lArrLoop)) / 2)
        Printer.Print arrToCenter(lArrLoop)
    Next
    Printer.EndDoc 'Initiates Printing To The Default Printer
Else: End If

End Sub
Private Sub mnusavelyrics_Click()
Dim iFree As Integer
Dim sBuff As String
Dim sDirRet As String

'Save Lyrics To Text File


If txtLyrics <> "" Then
    With CommonDialog1
        .DialogTitle = "Saving " & Songs_Lvw.SelectedItem.Text & " Lyrics"
        .DefaultExt = ".Txt"
        .FileName = RefList_Lvw.SelectedItem.Text & "-" & Songs_Lvw.SelectedItem.Text & ".Txt"
RE_TRY:
        .ShowSave
        If .FileName <> "" Then
            sDirRet = Dir(.FileName)
            If sDirRet <> "" Then
                Select Case MsgBox("Delete The Existing File?", vbYesNoCancel + vbDefaultButton3, "Confirm Overwrite")
                    Case vbYes
                        On Error Resume Next
                            Kill .FileName
                            If Err.Number <> 0 Then
                                MsgBox "The File Cannot Be Deleted, It May Be In Use", vbOKOnly, "File In Use"
                                Exit Sub
                            Else: End If
                        On Error GoTo 0
                    Case vbNo
                        GoTo RE_TRY:
                    Case Else
                        Exit Sub
                End Select
            Else: End If
            iFree = FreeFile
            sBuff = txtLyrics.Text
            Open .FileName For Binary Access Write As iFree
                Put iFree, 1, sBuff
            Close iFree
        Else: End If
    End With
Else: End If

End Sub
Private Sub RefList_Lvw_ItemClick(ByVal Item As MSComctlLib.ListItem)

Songs_Lvw.ListItems.Clear
If GetSongsList_Alpha(Item.Tag) = False Then
    MsgBox "No Songs By This Artist Are Available"
Else: End If

End Sub
Private Sub SearchA_but_Click()
Dim sArtist As String
Dim sTitle As String
Dim sAlphaTab As String
Dim arrArtist()
Dim iLoop As Integer
Dim lItm As ListItem

sArtist = Trim(Artist_Tbox)
sTitle = Trim(Song_Tbox)

'{The} Generally Isnt Use In The Artist Name So Remove It.
sArtist = Trim(Replace(sArtist, "THE ", "", , , vbTextCompare))

For iLoop = 0 To 128 'Remove All Non Used Chrs()
    Select Case iLoop
        Case 0 To 31, 33 To 47, 58 To 64, 123 To 128   'Remove Them
            sArtist = Replace(sArtist, Chr(iLoop), "")
            sTitle = Replace(sTitle, Chr(iLoop), "")
        Case Else
            '0-9 , A-Z or a-z
    End Select
Next

RefList_Lvw.ListItems.Clear
Songs_Lvw.ListItems.Clear
txtLyrics.Text = ""


bIgnoreClick = True 'Had To Add This So The Alpha_Opt_Click() Event Would Not Fire
    sAlphaTab = LCase(Left(sArtist, 1)) 'Must Use Lower Case Letters
    Select Case Asc(sAlphaTab) 'Just Changing The Alpha Button To Selected
        Case 48 To 57 'Numbers
            Alpha_Opt(0).Value = 1
        Case Else 'Letters
            Alpha_Opt(Asc(sAlphaTab) - 96).Value = True
    End Select
bIgnoreClick = False


GetArtistList_Alpha sAlphaTab, sArtist

If RefList_Lvw.ListItems.Count > 0 Then
    If sTitle <> "" Then
TOPOFLOOP:
        For Each lItm In RefList_Lvw.ListItems
            lItm.Selected = True
            If GetSongsList_Alpha(RefList_Lvw.SelectedItem.Tag, sTitle) = False Then
                RefList_Lvw.ListItems.Remove lItm.Index
                GoTo TOPOFLOOP
            Else: End If
        Next
        If Songs_Lvw.ListItems.Count = 1 Then
            Songs_Lvw.ListItems(1).Selected = True
            Songs_Lvw_ItemClick Songs_Lvw.ListItems(1)
        Else: End If
    Else
        If RefList_Lvw.ListItems.Count = 1 Then
            RefList_Lvw.ListItems(1).Selected = True
            RefList_Lvw_ItemClick RefList_Lvw.ListItems(1)
        Else: End If
    End If
Else
    MsgBox "No Matching Artists For: " & sArtist
End If

End Sub
Private Sub Alpha_Opt_Click(Index As Integer)

If bIgnoreClick = False Then
    GetArtistList_Alpha LCase(Alpha_Opt(Index).Caption)
Else: End If

End Sub
Private Function FindHREFS(sPage As String, TxtStart As String, TxtEnd As String, Optional TxtEndb As String, Optional sFilter As String) As Variant
Dim lCurPos As Long
Dim lMidPos As Long
Dim lEndPos As Long
Dim Arr() As String
Dim sTemp As String
Dim sHref As String
Dim sCap As String
Dim lRefStart As Long
Dim lRefEnd As Long
Dim iFltrLoop As Integer
Dim iMatch As Integer
Dim arrFilter
Dim lNextCRPos As Long

'This Function Just Extracts And Puts Into And Array
'Every Instance It Finds Between {TxtStart} And {TxtEnd}
'In Raw Html
'I Use It To Get <href> Links And Captions
'TxtStart = "<A href=" And TxtEnd = "</a><BR>"


sFilter = Trim(sFilter)
If sFilter <> "" Then
    arrFilter = Split(sFilter, Chr(32))
Else: End If

'Returnes An 0-1 Element Size Array
'element (0,#&) = RefWebPage
'element (1,#&) = RefCaption

ReDim Preserve Arr(1, 0)
sFilter = Trim(sFilter)
lCurPos = 1
lMidPos = 1
lEndPos = 1

Do
    lCurPos = InStr(lCurPos, sPage, TxtStart, vbTextCompare)
    If lCurPos < 1 Then GoTo EL_DUNNO:
    lNextCRPos = InStr(lCurPos, sPage, vbCrLf, vbTextCompare)
    lEndPos = InStr(lCurPos, sPage, TxtEnd, vbTextCompare)
    If lEndPos < 1 Then
        lEndPos = InStr(lCurPos, sPage, TxtEndb, vbTextCompare)
    Else: End If
    If lEndPos < 1 Then GoTo EL_DUNNO:
    If lEndPos > lNextCRPos Then
        lEndPos = lCurPos + Len(TxtStart) + 1
        GoTo DO_NEXT:
    Else: End If
    sTemp = Mid(sPage, lCurPos + Len(TxtStart) + 1, lEndPos - (lCurPos + Len(TxtStart)) - 1)
    sTemp = Replace(sTemp, "</B>", "", , , vbTextCompare)
    lMidPos = InStr(1, sTemp, ">", vbTextCompare)
    If lMidPos > 0 Then 'Has A Ref And Has A Caption
        sHref = Left(sTemp, lMidPos - 2)
        lRefEnd = InStr(1, sHref, Chr(34), vbTextCompare)
        If lRefEnd <> 0 Then
            sHref = Left(sHref, lRefEnd - 1)
        Else: End If
        If InStr(1, sHref, "http:", vbTextCompare) Then
            'Using Outside Web Sight, Ok Though Just Follow The Link
        ElseIf Left(sHref, 3) = "../" Then 'You Must Have Clicked On An Artist
            sHref = GetHttpRoot(RefList_Lvw.SelectedItem.Tag, sHref)
        Else
            If InStr(1, sHref, "/") Then
                sHref = sWebSiteRoot & sHref
            Else
                sHref = GetHttpRoot(RefList_Lvw.SelectedItem.Tag, sHref)
            End If
        End If
        sCap = Trim(Mid(sTemp, lMidPos + 1))
    Else
        sHref = ""
        sCap = sTemp
    End If
    RemoveBraces sCap
    If sCap = "" Then GoTo DO_NEXT:
    
    If InStr(1, sHref, ".html", vbTextCompare) Then
        If sFilter = "" Then
            ReDim Preserve Arr(1, UBound(Arr, 2) + 1)
            Arr(0, UBound(Arr, 2)) = sHref
            Arr(1, UBound(Arr, 2)) = sCap
        Else
            iMatch = 0
            For iFltrLoop = 0 To UBound(arrFilter)
                If InStr(1, sHref, arrFilter(iFltrLoop), vbTextCompare) Or InStr(1, sCap, arrFilter(iFltrLoop), vbTextCompare) Then
                    iMatch = iMatch + 1
                Else: End If
            Next
            
            If iMatch = (UBound(arrFilter) + 1) Then
                ReDim Preserve Arr(1, UBound(Arr, 2) + 1)
                Arr(0, UBound(Arr, 2)) = sHref
                Arr(1, UBound(Arr, 2)) = sCap
            Else: End If
        End If
    Else: End If
DO_NEXT:
    lCurPos = lCurPos + (lEndPos - lCurPos)
Loop

EL_DUNNO:

FindHREFS = Arr

End Function
Private Function GetHttpRoot(sFullHttp As String, sHref As String) As String
Dim sSlash As String
Dim lFound As Long

'Returns The Root Web Site Path

If InStr(1, sHref, "/") Then
    lFound = InStr(8, sFullHttp, "/")
    sSlash = Left(sFullHttp, lFound)
    sHref = Replace(sHref, "../", "")
    If Left(sHref, 1) = "/" Then sHref = Mid(sHref, 2)
    If Right(sSlash, 1) <> "/" Then sSlash = sSlash & "/"
    GetHttpRoot = sSlash & sHref
Else
    'Need To Pull The Entire Http Info
    lFound = InStrRev(sFullHttp, "/")
    sSlash = Left(sFullHttp, lFound)
    GetHttpRoot = sSlash & sHref
End If


End Function
Private Function FindAndStripAfter(sPage As String, sFindWhat As String) As Boolean
Dim lFound As Long
FindAndStripAfter = True
el_TOPO:

lFound = InStr(1, sPage, sFindWhat, vbTextCompare)
If lFound > 0 Then
    FindAndStripAfter = True
    sPage = Mid(sPage, lFound + Len(sFindWhat)) '+ Len(sFindWhat)
    If Right(sFindWhat, 1) = "<" Then
        sPage = "<" & sPage
    Else: End If
    GoTo el_TOPO
Else
    'FindAndStripAfter = False
End If

End Function
Private Function FindAndStripBefore(sPage As String, sFindWhat As String) As Boolean
Dim lFound As Long

'Usually There Is A Lot Of Junk In The Begining Of
'Some Of The Lyric Html Pages, So I Delete Andthing
'Usually Before </Form>

lFound = InStr(1, sPage, sFindWhat, vbTextCompare)
If lFound > 0 Then
    FindAndStripBefore = True
    sPage = Mid(sPage, lFound)  '+ Len(sFindWhat)
Else
    FindAndStripBefore = False
End If

End Function
Private Function GetPage(sHttp As String) As String
Dim bRet() As Byte
Dim sData As String
Dim lLoop As Long

'Could Getting An Html Page Be Any Easier.
bRet() = Inet1.OpenURL(sHttp, 1)
For lLoop = 0 To UBound(bRet) - 1
    sData = sData + Chr(bRet(lLoop))
Next
GetPage = sData

End Function
Private Function GetSongsList_Alpha(sHttpPath As String, Optional sFilter As String = "") As Boolean
Dim sPage As String
Dim RefArr
Dim iLoop As Long
Dim lItm As ListItem

'Find Any Songs Under A Clicked Artist Name

txtLyrics.Text = ""
sPage = GetPage(sHttpPath)

RefArr = FindHREFS(sPage, "<A href=", "</a><BR>", "</a></h3>", sFilter)
If UBound(RefArr, 2) > 0 Then
    GetSongsList_Alpha = True
    Songs_Lvw.ListItems.Clear
    For iLoop = 1 To UBound(RefArr, 2)
        Set lItm = Songs_Lvw.ListItems.Add(, , RefArr(1, iLoop))
        lItm.ToolTipText = RefArr(0, iLoop)
        lItm.Tag = RefArr(0, iLoop)
    Next 'iLoop
Else
    'Msgbox "No Reference Found <href>"
End If

End Function
Private Function RemoveBraces(sPage As String)
Dim lStart As Long
Dim lEnd As Long
Dim sTemp As String
Dim sLookForThis As String

'Not Really Braces But Greater/Less Than Operators <>
'Removes Anything and Everything Between < and >
'Example: sPage = "<Head><Title>Hello World</Title><Head>"
'Returnes: "Hello World"

lStart = 1
lEnd = 1

'Remove Head Tags
lStart = InStr(1, sPage, "<HEAD>", vbTextCompare)
lEnd = InStr(1, sPage, "</HEAD>", vbTextCompare)
If lStart <> 0 And lEnd <> 0 Then
    sTemp = Mid(sPage, lStart, (lEnd - lStart) + 1)
    sPage = Replace(sPage, sTemp, "")
    sPage = Trim(sPage)
Else: End If
'

lStart = 1
lEnd = 1
Do
    lStart = InStr(lStart, sPage, "<")
    If lStart = 0 Then GoTo QUIT_NOW:
    lEnd = InStr(lStart, sPage, ">")
    If lEnd = 0 Then
        lEnd = Len(sPage)
    Else: End If
    sTemp = Mid(sPage, lStart, (lEnd - lStart) + 1)
    sPage = Replace(sPage, sTemp, "")
    sPage = Trim(sPage)
Loop

QUIT_NOW:
End Function
Private Sub Songs_Lvw_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim sPage As String
Dim RefArr
Dim iLoop As Long
Dim lItm As ListItem
Dim sTitle As String

txtLyrics.Text = ""
sPage = GetPage(Item.Tag)
FindAndStripAfter sPage, "LYRICS</B><BR><BR>"
sPage = Replace(sPage, vbCrLf, "")
sPage = Replace(sPage, "<BR>", vbCrLf, , , vbTextCompare)
RemoveBraces sPage
txtLyrics.Text = sPage

End Sub
