' ==============================================================================
' video-shiwake
' ==============================================================================
' OS: Windows 10, Windows 11
' Charset: UTF-8
' EOL: CRLF
' ==============================================================================
' 動画ファイルの「メディアの作成日時」を読み取り、月毎のフォルダに移動します。
' フォルダ移動の際には、ファイル名の先頭に作成日時を付与します。
' ------------------------------------------------------------------------------
' 当プログラムのご利用は自己責任でお願いします。
' 当プログラムがいかなる問題を引き起こしたとしても、作成者はその責任を負いません。
' ==============================================================================

' ==============================================================================
' 定数
' ------------------------------------------------------------------------------
' 入力元フォルダ
dim INPUT_DIR: INPUT_DIR = "D:\Videos\_unsorted"
' 出力先フォルダ
dim OUTPUT_DIR: OUTPUT_DIR = "D:\Videos\"
' タイムスタンプが記録されていない場合に出力するフォルダ
dim UNSORTABLE_DIR_NAME: UNSORTABLE_DIR_NAME = "_unsortable"

' ==============================================================================
' Main
' ------------------------------------------------------------------------------
dim mRegexp: set mRegexp = CreateObject("VBScript.RegExp")
mRegexp.Pattern = "[^0-9/ :]"
mRegexp.Global = True

call scanMediaFiles()

set mRegexp = Nothing
WScript.Quit

' ==============================================================================
' 動画ファイルの仕分け処理
' ------------------------------------------------------------------------------
sub scanMediaFiles()
  dim originalName
  dim created
  dim dt
  dim strYearMonth
  dim strDate
  dim strTime
  dim newName
  dim newPath
  dim Shell
  dim Folder

  dim fso: set fso = CreateObject("Scripting.FileSystemObject")
  dim files: set files = fso.GetFolder(INPUT_DIR).files
  dim file
  for each file in files
    originalName = Replace(file, INPUT_DIR & "\", "")

    set Shell = CreateObject("Shell.Application")
    set Folder = Shell.Namespace(fso.GetFile(file).ParentFolder.path)

    ' メディアの作成日時を取得
    ' [Note] Windows 10 では `208` で動くが、Windows 7 などでは違う数値だったらしい
    created = Folder.GetDetailsOf(Folder.ParseName(fso.GetFile(file).name), 208)
    dt = convDate(created)

    if IsEmpty(dt) then
      ' メディアの作成日時が取得できない場合
      newFile = OUTPUT_DIR & "\" & UNSORTABLE_DIR_NAME & "\" & originalName
      call makeDir(UNSORTABLE_DIR_NAME)
      call moveFile(file, newFile)
    else
      ' メディアの作成日時が取得できた場合
      strYearMonth = Year(dt) _
        & "-" & Right("0" & Month(dt), 2)
      strDate = strYearMonth _
        & "-" & Right("0" & Day(dt), 2)
      strTime = Right("0" & Hour(dt), 2) _
        & "-" & Right("0" & Minute(dt), 2)
      newName = strDate & "-" & strTime & "-" & originalName
      newFile = OUTPUT_DIR & "\" & strYearMonth & "\" & newName
      ' Make a directory and move a file
      call makeDir(strYearMonth)
      call moveFile(file, newFile)
    end if

    set Folder = Nothing
    set Shell = Nothing
  next
end sub

' ==============================================================================
' 撮影日(文字列)をDate型へ変換
' ------------------------------------------------------------------------------
function convDate(strTimestamp)
  dim text: text = Trim(mRegexp.Replace(strTimestamp, ""))
  if Len(text) = 0 then
    convDate = Empty
  else
    convDate = CDate(text)
  end if
end function

' ==============================================================================
' フォルダの作成
' ------------------------------------------------------------------------------
sub makeDir(folder)
  dim fso: set fso = CreateObject("Scripting.FileSystemObject")
  dim mkdirPath: mkdirPath = OUTPUT_DIR & "\" & folder
  if not fso.FolderExists(mkdirPath) Then
    fso.CreateFolder(mkdirPath)
  end if
  set fso = Nothing
end sub

' ==============================================================================
' ファイルの移動
' ------------------------------------------------------------------------------
sub moveFile(pathFrom, pathTo)
  with CreateObject("Scripting.FileSystemObject")
    if .FileExists(pathTo) Then
      ' 移動先にファイルが存在する場合
      if .GetFile(pathFrom).Path = .GetFile(pathTo).Path then
        ' 移動元と移動先が同じなら何もしない
      Elseif .GetFile(pathFrom).Size = .GetFile(pathTo).Size then
        ' ファイル名もサイズが同じなら既に仕分け済みであると判定し、移動せず削除
        .DeleteFile pathFrom
      Else
        ' ファイル名が重複している場合、エラーメッセージを出力
        WScript.StdOut.WriteLine "moveFile - [Error] Dupulicated: " & pathTo
      end if
    else
      ' ファイルを移動
      call .MoveFile(pathFrom, pathTo)
    end if
  end with
end sub
