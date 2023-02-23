' ==============================================================================
' video-shiwake
' ==============================================================================
' OS: Windows 10
' Charset: UTF-8
' EOL: CRLF
' ==============================================================================
' 動画ファイルの「メディアの作成日時」を読み取り、月毎のフォルダに移動します。
' フォルダ移動の際には、ファイル名の先頭に作成日時を付与します。
' 当プログラムは Windows 10 以外のOSでは期待通りに動作しない可能性があります。
' ------------------------------------------------------------------------------
' ご利用は自己責任でお願いします。
' 当プログラムがいかなる問題を引き起こしたとしても、開発者はその責任を負いません。
' ==============================================================================
' INPUT DIRECTORY
dim inputDir: inputDir = "D:\Videos\_unsorted"
' OUTPUT DIRECTORY
dim outputDir: outputDir = "D:\Videos\"
' ==============================================================================

dim mRegexp: set mRegexp = CreateObject("VBScript.RegExp")
mRegexp.Pattern = "[^0-9/ :]"
mRegexp.Global = True
call scanMediaFiles()
set mRegexp = Nothing
WScript.Quit

sub scanMediaFiles()
  dim fileName
  dim created
  dim dt
  dim strYearMonth
  dim strDate
  dim strTime
  dim newName
  dim newPath
  dim SHell
  dim Folder
  with CreateObject("Scripting.FileSystemObject")
    dim files: set files = .GetFolder(inputDir).files
    dim file
    for each file in files
      fileName = Replace(file, inputDir & "\", "")

      set SHell = CreateObject("Shell.Application")
      set Folder = SHell.Namespace(.GetFile(file).ParentFolder.path)

      ' [Note] Windows 10 では `208` で動くが、Windows 7 などでは違う数値だったらしい
      created = Folder.GetDetailsOf(Folder.ParseName(.GetFile(file).name), 208)

      dt = cleanDT(created)
      if IsEmpty(dt) = False then
        strYearMonth = Year(dt) _
          & "-" & Right("0" & Month(dt), 2)
        strDate = strYearMonth _
          & "-" & Right("0" & Day(dt), 2)
        strTime = Right("0" & Hour(dt), 2) _
          & "-" & Right("0" & Minute(dt), 2)
        newName = strDate & "-" & strTime & "-" & fileName
        newFile = outputDir & "\" & strYearMonth & "\" & newName
        ' Make a directory
        call makeDir(strYearMonth)
        ' Move file
        if .FileExists(newFile) Then
          if .GetFile(file).Path = .GetFile(newFile).Path then
            ' through
          Elseif .GetFile(file).Size = .GetFile(newFile).Size then
            .DeleteFile file
          Else
            WScript.StdOut.WriteLine "moveFile - [Error] Dupulicated: " & newFile
          end if
        else
          call .MoveFile(file, newFile)
        end if
      end if

      set Folder = Nothing
      set SHell = Nothing
    next
  end with
end sub

function cleanDT(strTimestamp)
  dim text: text = Trim(mRegexp.Replace(strTimestamp, ""))
  if Len(text) = 0 then
    cleanDT = Empty
  else
    cleanDT = CDate(text)
  end if
end function

function makeDir(folder)
  dim fso: set fso = CreateObject("Scripting.FileSystemObject")
  dim mkdirPath: mkdirPath = outputDir & "\" & folder
  if not fso.FolderExists(mkdirPath) Then
    fso.CreateFolder(mkdirPath)
  end if
  set fso = Nothing
end function
