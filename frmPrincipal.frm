VERSION 5.00
Begin VB.Form frmPrincipal 
   Caption         =   "BuscaTexto 1.0"
   ClientHeight    =   5640
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   13335
   Icon            =   "frmPrincipal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   13335
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Informações para a varredura:"
      Height          =   1335
      Left            =   7320
      TabIndex        =   10
      Top             =   120
      Width           =   5895
      Begin VB.CommandButton cmdIniciar 
         BackColor       =   &H0080FF80&
         Caption         =   "Clique aqui para iniciar a varredura"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   840
         Width           =   3975
      End
      Begin VB.TextBox txtTexto 
         Height          =   405
         Left            =   1800
         TabIndex        =   1
         Top             =   360
         Width           =   3975
      End
      Begin VB.Label Label1 
         Caption         =   "Informe o texto a ser buscado nos arquivos:"
         Height          =   435
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1770
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Fechar"
      Height          =   375
      Left            =   12000
      TabIndex        =   5
      Top             =   5160
      Width           =   1215
   End
   Begin VB.ListBox List2 
      Height          =   1425
      Left            =   120
      TabIndex        =   4
      Top             =   3600
      Width           =   13095
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   13095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Informe o diretório a ser escaneado:"
      Height          =   1335
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   7095
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   375
         Left            =   6480
         TabIndex        =   0
         Top             =   480
         Width           =   495
      End
      Begin VB.Label txtCaminho 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   6255
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "www.pilandia.com.br | pilandia@gmail.com"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   5400
      Width           =   3000
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Desenvolvido por Vinicius Hernandes"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   5160
      Width           =   2655
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Arquivos que contém o texto procurado:"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   3360
      Width           =   2835
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Arquivos lidos:"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   1020
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Caixa diálogo p/ selecionar diretótio----------------
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260
Private Declare Function SHBrowseForFolder Lib "shell32" _
(lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" _
(ByVal pidList As Long, _
ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" _
(ByVal lpString1 As String, ByVal _
lpString2 As String) As Long
Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

'Busca dos arquivos-------------------
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Const MAXDWORD = &HFFFF
Const INVALID_HANDLE_VALUE = -1
Const FILE_ATTRIBUTE_ARCHIVE = &H20
Const FILE_ATTRIBUTE_DIRECTORY = &H10
Const FILE_ATTRIBUTE_HIDDEN = &H2
Const FILE_ATTRIBUTE_NORMAL = &H80
Const FILE_ATTRIBUTE_READONLY = &H1
Const FILE_ATTRIBUTE_SYSTEM = &H4
Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

Dim Arquivo As String
Dim LArq As Boolean
Dim strLinha As String
Dim Mycheck As Boolean
Dim VariávelPesquisa As String
Dim Texto As String

Function StripNulls(OriginalStr As String) As String
    If (InStr(OriginalStr, Chr(0)) > 0) Then
        OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
    End If
    StripNulls = OriginalStr
End Function

Function FindFilesAPI(path As String, SearchStr As String, FileCount As Integer, DirCount As Integer)
    'KPD-Team 1999
    'E-Mail: [email]KPDTeam@Allapi.net[/email]
    'URL: [url]http://www.allapi.net/[/url]

    Dim FileName As String ' Walking filename variable...
    Dim DirName As String ' SubDirectory Name
    Dim dirNames() As String ' Buffer for directory name entries
    Dim nDir As Integer ' Number of directories in this path
    Dim i As Integer ' For-loop counter...
    Dim hSearch As Long ' Search Handle
    Dim WFD As WIN32_FIND_DATA
    Dim Cont As Integer
    If Right(path, 1) <> "\" Then path = path & "\"
    ' Search for subdirectories.
    nDir = 0
    ReDim dirNames(nDir)
    Cont = True
    hSearch = FindFirstFile(path & "*", WFD)
    If hSearch <> INVALID_HANDLE_VALUE Then
        Do While Cont
        DoEvents
        DirName = StripNulls(WFD.cFileName)
        ' Ignore the current and encompassing directories.
        If (DirName <> ".") And (DirName <> "..") Then
            ' Check for directory with bitwise comparison.
            If GetFileAttributes(path & DirName) And FILE_ATTRIBUTE_DIRECTORY Then
                dirNames(nDir) = DirName
                DirCount = DirCount + 1
                nDir = nDir + 1
                ReDim Preserve dirNames(nDir)
            End If
        End If
        Cont = FindNextFile(hSearch, WFD) 'Get next subdirectory.
        Loop
        Cont = FindClose(hSearch)
    End If
    ' Walk through this directory and sum file sizes.
    hSearch = FindFirstFile(path & SearchStr, WFD)
    Cont = True
    If hSearch <> INVALID_HANDLE_VALUE Then
        While Cont
            DoEvents
            FileName = StripNulls(WFD.cFileName)
            If (FileName <> ".") And (FileName <> "..") Then
                FindFilesAPI = FindFilesAPI + (WFD.nFileSizeHigh * MAXDWORD) + WFD.nFileSizeLow
                FileCount = FileCount + 1
                List1.AddItem path & FileName
                List1.ListIndex = List1.ListCount - 1
                LArq = ChecaArquivo(path & FileName)
                If LArq = True Then
                    List2.AddItem path & FileName
                End If
            End If
            Cont = FindNextFile(hSearch, WFD) ' Get next file
        Wend
        Cont = FindClose(hSearch)
    End If
    ' If there are sub-directories...
    If nDir > 0 Then
        ' Recursively walk into them...
        For i = 0 To nDir - 1
            FindFilesAPI = FindFilesAPI + FindFilesAPI(path & dirNames(i) & "\", SearchStr, FileCount, DirCount)
        Next i
    End If
End Function


Private Sub cmdIniciar_Click()
    Dim SearchPath As String, FindStr As String
    Dim FileSize As Long
    Dim NumFiles As Integer, NumDirs As Integer
    
    If Trim(txtCaminho) = "" Then
        MsgBox "Informe o diretório que deseja escanear!", vbInformation
        Command1.SetFocus
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    List1.Clear
    List2.Clear
    SearchPath = txtCaminho
    FindStr = "*.*"
    FileSize = FindFilesAPI(SearchPath, FindStr, NumFiles, NumDirs)
    'Text3.Text = NumFiles & " Files found in " & NumDirs + 1 & " Directories"
    'Text4.Text = "Size of files found under " & SearchPath & " = " & Format(FileSize, "#,###,###,##0") & " Bytes"
    Screen.MousePointer = vbDefault
    
    MsgBox "Escaneamento finalizado!", vbInformation
End Sub

Private Sub Command1_Click()
'Opens a Treeview control that displays the directories in a computer

    Dim lpIDList As Long
    Dim sBuffer As String
    Dim szTitle As String
    Dim tBrowseInfo As BrowseInfo

    szTitle = "This is the title"
    With tBrowseInfo
        .hWndOwner = Me.hWnd
        .lpszTitle = lstrcat(szTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With

    lpIDList = SHBrowseForFolder(tBrowseInfo)

    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        txtCaminho = sBuffer
        cmdIniciar.SetFocus
    End If
End Sub

Function ChecaArquivo(arq As String) As Boolean
    On Error GoTo Erro
    Open arq For Input As #1
    ChecaArquivo = False
       
    While Not EOF(1)
        Line Input #1, strLinha
    
        VariávelPesquisa = UCase$(strLinha)
        Texto = "*" & UCase$(txtTexto.Text) & "*"
        Mycheck = VariávelPesquisa Like Texto
        If Mycheck = True Then
            ChecaArquivo = True
            Close #1
            Exit Function
        End If
    
    Wend
    Close #1
    Exit Function
Erro:
    ChecaArquivo = False
End Function

Private Sub Command2_Click()
    Unload Me
End Sub
