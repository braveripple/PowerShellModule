
Function Make-ZIP {
<#
.Synopsis
   ZIPファイルの作成
.DESCRIPTION
   ZIPファイルを作成します。
   対象ファイルをパイプライン経由で渡すこともできます。
.EXAMPLE
   Make-ZIP -ZipFilePath "item.zip" -TargetItems @("item1.txt","item2.txt","item3.txt")
.EXAMPLE
   Get-ChildItem -Path "item*.txt" | Make-ZIP -ZipFilePath "item.zip"
#>
  [CmdletBinding()]
  [OutputType([System.IO.FileInfo[]])]
  Param (
    [Parameter(Mandatory=$True)]
      [String]$ZipFilePath,
   
    [Parameter(Mandatory=$True,ValueFromPipeline=$True)]
      [ValidateScript({
        # パラメーターをデータ型の配列で宣言している場合、
        # $_にはデータ型が入る
        If (Test-Path -LiteralPath $_) { return $True }
        Throw "ファイルが存在しません：" + (Get-AbsolutePath $_)
      })]
      [String[]]$TargetItems
  )
  Begin {
    # Zipファイル作成（同一名称ファイルが存在する場合には予め削除する）
    If(Test-Path -Path $ZipFilePath) {
      Remove-Item -Path $ZipFilePath
    }
    $ZipFile = ([char]80 + [char]75 + [char]5 + [char]6).ToString() + ([char]0).ToString()*18 | `
      New-Item -Path $ZipFilePath -Type File
      
    # Zipコンテナ作成
    $Shell = New-Object -ComObject Shell.Application
    $ZipContainer = $Shell.NameSpace($ZipFile.FullName)
    
    $ArchivedItems = New-Object System.Collections.Generic.List[System.IO.FileInfo]
  }
  Process {
    ForEach($tItem in ($TargetItems | %{ Get-Item -Path $_ } )) {
      ForEach($aItem in $ArchivedItems) {
        # 同一名称のアイテムが存在した場合、zipファイルを削除して例外をスロー
        If($tItem.Name -eq $aItem.Name) {
          Throw "ファイル名またはフォルダ名が重複しています：" `
            + $aItem.FullName + "," + $tItem.FullName
        }
      }
      $ArchivedItems.Add($tItem) | Out-Null
      # Zipコンテナにファイルをコピー
      $ZipContainer.CopyHere($tItem.FullName)
      While($True) {
        # 非同期処理のため、完了するまで一定時間待つ
        If($ArchivedItems.Count -eq $ZipContainer.Items().Count) {
          Break
        }
        Start-Sleep -Milliseconds 100
      }
    }
  }
  End {
    Return $ArchivedItems.ToArray()
  }
}

Export-ModuleMember -Function Make-ZIP
