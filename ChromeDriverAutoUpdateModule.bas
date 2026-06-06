Option Explicit
'====================================================================================================
' ChromeDriver 自動更新モジュール
'   - chrome.exe のビルドに合わせて Chrome for Testing の chromedriver(win64)を取得・配置する
'   - 依存:参照設定「Microsoft Scripting Runtime」 / JsonConverter(VBA-JSON)モジュール
'   - 対象:VBA7(Office 2010 以降)。LongPtr により 32bit / 64bit いずれの Office でも動作
'====================================================================================================

Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
    (ByVal pCaller As LongPtr, _
     ByVal szURL As String, _
     ByVal szFileName As String, _
     ByVal dwReserved As Long, _
     ByVal lpfnCB As LongPtr) As Long

Private Declare PtrSafe Function SHCreateDirectoryEx Lib "shell32.dll" Alias "SHCreateDirectoryExA" _
    (ByVal hwnd As LongPtr, _
     ByVal pszPath As String, _
     ByVal psa As LongPtr) As Long

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Type VersionType
    Major As Long
    Minor As Long
    Build As Long
    Revision As Long
    BuildVersion As String      ' Major.Minor.Build(リビジョンを除いたビルドまで)
    RevisionVersion As String   ' Major.Minor.Build.Revision(フルバージョン文字列)
End Type

Private Const ZIP_FILE As String = "chromedriver.zip"
Private workPath As String
Private objFso As New Scripting.FileSystemObject

Public Function ChromeDriverAutoUpdate(Optional ByVal ForcedExecution As Boolean = False) As Boolean
'====================================================================================================
' chrome.exe と chromedriver.exe のビルドを比較し、必要なら chromedriver を自動更新する
' ForcedExecution = True のときはビルド一致でも強制的に更新処理を行う
'====================================================================================================
    Dim chromePath As String
    Dim chromeFullpath As String
    Dim chromeVersion As VersionType
    Dim chromedriverPath As String
    Dim chromedriverFullPath As String
    Dim objFolder As Scripting.Folder
    Dim lngRevision As Long
    Dim bestRevision As Long
    Dim bestFolderName As String
    Dim foldersToDelete As Collection
    Dim p As Variant

    On Error GoTo ErrLabel

    ' ---ダウンロード用フォルダを作成(Python/Selenium に合わせたパス)
    workPath = Environ("USERPROFILE") & "\.cache\selenium\seleniumbasic"
    Select Case SHCreateDirectoryEx(0&, workPath, 0&)
        Case 0, 183
            ' 0=作成成功 / 183=ERROR_ALREADY_EXISTS(作成済み)
        Case Else
            MsgBox "ダウンロード用フォルダを作成できませんでした" & vbCrLf & workPath, vbCritical
            ChromeDriverAutoUpdate = False
            Exit Function
    End Select

    ' ---chrome 本体のフォルダを探す
    Select Case True
        Case objFso.FolderExists(Environ("ProgramW6432") & "\Google\Chrome\Application")
            chromePath = Environ("ProgramW6432") & "\Google\Chrome\Application"
        Case objFso.FolderExists(Environ("ProgramFiles") & "\Google\Chrome\Application")
            chromePath = Environ("ProgramFiles") & "\Google\Chrome\Application"
        Case objFso.FolderExists(Environ("LOCALAPPDATA") & "\Google\Chrome\Application")
            chromePath = Environ("LOCALAPPDATA") & "\Google\Chrome\Application"
        Case Else
            MsgBox "'chrome' フォルダが見つかりません", vbCritical
            ChromeDriverAutoUpdate = False
            Exit Function
    End Select

    If objFso.FileExists(chromePath & "\chrome.exe") Then
        chromeFullpath = chromePath & "\chrome.exe"
    Else
        MsgBox "'chrome.exe' が見つかりません", vbCritical
        ChromeDriverAutoUpdate = False
        Exit Function
    End If

    ' ---SeleniumBasic(chromedriver の配置先)フォルダを探す
    Select Case True
        Case objFso.FolderExists(Environ("ProgramW6432") & "\SeleniumBasic")
            chromedriverPath = Environ("ProgramW6432") & "\SeleniumBasic"
        Case objFso.FolderExists(Environ("ProgramFiles") & "\SeleniumBasic")
            chromedriverPath = Environ("ProgramFiles") & "\SeleniumBasic"
        Case objFso.FolderExists(Environ("LOCALAPPDATA") & "\SeleniumBasic")
            chromedriverPath = Environ("LOCALAPPDATA") & "\SeleniumBasic"
        Case Else
            MsgBox "'SeleniumBasic' のフォルダが見つかりません", vbCritical
            ChromeDriverAutoUpdate = False
            Exit Function
    End Select
    chromedriverFullPath = chromedriverPath & "\chromedriver.exe"   ' ※初回は未配置のこともある

    ' ---chrome.exe のバージョンを取得
    If GetChromeVersion(chromeFullpath, chromeVersion) = False Then
        MsgBox "'chrome.exe' のバージョンが取得できませんでした", vbCritical
        ChromeDriverAutoUpdate = False
        Exit Function
    End If

    ' ---既存 chromedriver が現行 chrome と同じビルドなら更新不要(強制実行時は除く)
    If Not ForcedExecution Then
        If objFso.FileExists(chromedriverFullPath) Then
            If objFso.GetFileVersion(chromedriverFullPath) Like chromeVersion.BuildVersion & ".*" Then
                ChromeDriverAutoUpdate = True
                Exit Function
            End If
        End If
    End If

    ' ---chrome のビルドに合う chromedriver をダウンロード(キャッシュ済みはスキップ)
    If ChromedriverCheck(chromeVersion) = False Then
        MsgBox "'chromedriver.exe' のダウンロードに失敗しました", vbCritical
        ChromeDriverAutoUpdate = False
        Exit Function
    End If

    ' ---使用する chromedriver を選ぶ(ビルド一致かつリビジョンが chrome 以下で最大のもの)
    bestRevision = -1
    bestFolderName = ""
    Set foldersToDelete = New Collection
    For Each objFolder In objFso.GetFolder(workPath).SubFolders
        If objFolder.Name Like chromeVersion.BuildVersion & ".*" Then
            If objFso.FileExists(objFolder.Path & "\chromedriver.exe") Then
                lngRevision = CLng(Split(objFolder.Name, ".")(3))
                If lngRevision <= chromeVersion.Revision And lngRevision > bestRevision Then
                    bestRevision = lngRevision
                    bestFolderName = objFolder.Name
                End If
            End If
        Else
            ' 別ビルドの不要フォルダ ※列挙中に削除するとスキップが起きるので一旦控える
            foldersToDelete.Add objFolder.Path
        End If
    Next

    For Each p In foldersToDelete
        objFso.DeleteFolder p, True
    Next

    If bestFolderName = "" Then
        MsgBox "chrome (" & chromeVersion.RevisionVersion & ") に適合する chromedriver が見つかりませんでした", vbCritical
        ChromeDriverAutoUpdate = False
        Exit Function
    End If

    ' ---選んだ chromedriver を配置先へ上書きコピー
    objFso.GetFile(workPath & "\" & bestFolderName & "\chromedriver.exe").Copy chromedriverPath & "\chromedriver.exe", True

    ' ---更新不要だった場合も含めて正常終了は True
    ChromeDriverAutoUpdate = True
    Exit Function
ErrLabel:
    MsgBox "chromedriver の入替に失敗しました" & vbCrLf & "[" & Err.Description & "]" & vbCrLf & _
           "※実行中の chromedriver があれば終了してから再実行してください", vbCritical
    ChromeDriverAutoUpdate = False
End Function

Private Function GetChromeVersion(ByVal chromeFullpath As String, ByRef chromeVersion As VersionType) As Boolean
'====================================================================================================
' chrome.exe のファイルバージョンを取得して VersionType に展開する(PowerShell 不要)
'====================================================================================================
    Dim parts() As String

    On Error GoTo ErrLabel
    ' ---初期値
    chromeVersion.Major = 1
    chromeVersion.Minor = 0
    chromeVersion.Build = 0
    chromeVersion.Revision = 0

    ' ---ファイルバージョンを直読み(コンソールが出ない・速い・引用符の心配なし)
    chromeVersion.RevisionVersion = Trim(objFso.GetFileVersion(chromeFullpath))

    ' ---念のため正規表現で形式を確認
    With CreateObject("VBScript.RegExp")
        .Pattern = "^\d+\.\d+\.\d+(\.\d+)?$"
        If Not .Test(chromeVersion.RevisionVersion) Then
            GetChromeVersion = False
            Exit Function
        End If
    End With

    ' ---分解して格納
    parts = Split(chromeVersion.RevisionVersion, ".")
    chromeVersion.Major = CLng(parts(0))
    chromeVersion.Minor = CLng(parts(1))
    chromeVersion.Build = CLng(parts(2))
    If UBound(parts) >= 3 Then chromeVersion.Revision = CLng(parts(3))
    chromeVersion.BuildVersion = Join(Array(chromeVersion.Major, chromeVersion.Minor, chromeVersion.Build), ".")

    GetChromeVersion = True
    Exit Function
ErrLabel:
    MsgBox "chrome.exe のバージョン情報取得に失敗しました" & vbCrLf & "[" & Err.Description & "]", vbCritical
    GetChromeVersion = False
End Function

Private Function ChromedriverCheck(ByRef chromeVersion As VersionType) As Boolean
'====================================================================================================
' Chrome for Testing の JSON から chrome のビルドに一致する chromedriver(win64)を取得する
' ※リビジョンが chrome 以下のものだけ対象にする(それより新しい版は選ばれないため取得しない)
'====================================================================================================
    Dim objHttp As Object
    Dim objRet As Object
    Dim objVersion As Object
    Dim chromedriver As Variant
    Dim ver As String

    Const JSON_ENDPOINTS_URL As String = "https://googlechromelabs.github.io/chrome-for-testing/known-good-versions-with-downloads.json"
    Const TARGET_PLATFORM As String = "win64"   ' ※Windows11 以降は 64bit のみのため決め打ち

    On Error GoTo ErrLabel
    Set objHttp = CreateObject("MSXML2.XMLHTTP.6.0")
    objHttp.Open "GET", JSON_ENDPOINTS_URL, False
    objHttp.setRequestHeader "Cache-Control", "no-cache"   ' キャッシュ由来の古い情報を避ける
    objHttp.Send
    Set objRet = JsonConverter.ParseJson(objHttp.responseText)

    For Each objVersion In objRet("versions")
        ver = objVersion("version")
        ' ---ビルド一致 かつ リビジョンが chrome 以下のものだけ対象
        If ver Like chromeVersion.BuildVersion & ".*" Then
            If CLng(Split(ver, ".")(3)) <= chromeVersion.Revision Then
                ' ---未取得のものだけダウンロード
                If objFso.FolderExists(workPath & "\" & ver) = False Then
                    ' ---古い版は downloads に chromedriver が無いことがあるので存在確認
                    If objVersion("downloads").Exists("chromedriver") Then
                        For Each chromedriver In objVersion("downloads")("chromedriver")
                            If chromedriver("platform") = TARGET_PLATFORM Then
                                If DownloadChromedriver(chromedriver("url"), ver) = False Then
                                    ChromedriverCheck = False
                                    Exit Function
                                End If
                            End If
                        Next
                    End If
                End If
            End If
        End If
    Next

    ChromedriverCheck = True
    Exit Function
ErrLabel:
    MsgBox "chromedriver.exe の更新に失敗しました" & vbCrLf & "[" & Err.Description & "]", vbCritical
    ChromedriverCheck = False
End Function

Private Function DownloadChromedriver(ByVal url As String, ByVal targetVersion As String) As Boolean
'====================================================================================================
' zip をダウンロードして解凍し、chromedriver.exe をバージョンフォルダ直下へ配置する
'====================================================================================================
    Dim downloadPath As String
    Dim zipFullPath As String
    Dim newDriverPath As String
    Dim objFolder As Scripting.Folder
    Dim subPaths As Collection
    Dim p As Variant
    Dim waited As Long

    On Error GoTo ErrLabel
    downloadPath = workPath & "\" & targetVersion
    zipFullPath = workPath & "\" & ZIP_FILE

    ' ---バージョン別フォルダを作成
    Select Case SHCreateDirectoryEx(0&, downloadPath, 0&)
        Case 0, 183
            ' OK(作成成功 or 作成済み)
        Case Else
            MsgBox "ChromeDriver 用フォルダを作成できませんでした" & vbCrLf & downloadPath, vbCritical
            DownloadChromedriver = False
            Exit Function
    End Select

    ' ---zip をダウンロード
    If URLDownloadToFile(0&, url, zipFullPath, 0&, 0&) <> 0 Then
        MsgBox "ChromeDriver をダウンロードできませんでした" & vbCrLf & url, vbCritical
        DownloadChromedriver = False
        Exit Function
    End If

    ' ---zip を解凍(CopyHere は非同期 + サイレント指定)
    With CreateObject("Shell.Application")
        ' &H4=進捗UI非表示 &H10=上書き確認なし &H200=フォルダ作成確認なし &H400=エラーUI非表示
        .Namespace((downloadPath)).CopyHere .Namespace((zipFullPath)).Items, &H4 Or &H10 Or &H200 Or &H400
    End With

    ' ---解凍は非同期なので、目的の exe が現れるまで待つ(最大30秒)
    Do
        newDriverPath = SearchFilesRecursively(downloadPath & "\", "chromedriver.exe")
        If newDriverPath <> "" Then Exit Do
        Sleep 200
        waited = waited + 200
    Loop While waited < 30000

    If newDriverPath = "" Then
        MsgBox "chromedriver.exe の解凍に失敗しました" & vbCrLf & downloadPath, vbCritical
        DownloadChromedriver = False
        Exit Function
    End If

    ' ---chromedriver.exe をバージョンフォルダ直下へ移動(既に直下にあれば移動不要)
    If StrComp(newDriverPath, downloadPath & "\chromedriver.exe", vbTextCompare) <> 0 Then
        objFso.MoveFile newDriverPath, downloadPath & "\"
    End If

    ' ---解凍で出来た空サブフォルダを削除 ※列挙中に削除するとスキップするので一旦控える
    Set subPaths = New Collection
    For Each objFolder In objFso.GetFolder(downloadPath).SubFolders
        subPaths.Add objFolder.Path
    Next
    For Each p In subPaths
        objFso.DeleteFolder p, True
    Next

    ' ---zip を削除
    If objFso.FileExists(zipFullPath) Then objFso.DeleteFile zipFullPath, True

    DownloadChromedriver = True
    Exit Function
ErrLabel:
    MsgBox "chromedriver の解凍・配置に失敗しました" & vbCrLf & "[" & Err.Description & "]", vbCritical
    DownloadChromedriver = False
End Function

Private Function SearchFilesRecursively(ByVal folderPath As String, ByVal fileName As String) As String
'====================================================================================================
' folderPath を起点にサブフォルダまで再帰探索し、fileName のフルパスを返す(無ければ "")
'====================================================================================================
    Dim subFolder As Scripting.Folder
    Dim objFile As Scripting.File
    Dim found As String

    ' ---このフォルダ直下を探す
    For Each objFile In objFso.GetFolder(folderPath).Files
        If StrComp(objFile.Name, fileName, vbTextCompare) = 0 Then
            SearchFilesRecursively = objFile.Path
            Exit Function
        End If
    Next objFile

    ' ---サブフォルダを再帰探索(見つかった戻り値をそのまま採用する)
    For Each subFolder In objFso.GetFolder(folderPath).SubFolders
        found = SearchFilesRecursively(subFolder.Path, fileName)
        If found <> "" Then
            SearchFilesRecursively = found
            Exit Function
        End If
    Next subFolder

    SearchFilesRecursively = ""
End Function

