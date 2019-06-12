class CardWirthReader : System.IDisposable {

  [System.IO.BinaryReader] $reader
  [System.Text.Encoding] $encoding

  CardWirthReader() {
    $this.encoding = [System.Text.Encoding]::GetEncoding("Shift_JIS")
  }

  [void] Open($file) {
    $this.reader = [System.IO.BinaryReader]::new([System.IO.File]::Open($file, [System.IO.FileMode]::Open))
  }
  [void] Close() {
    if ($this.reader -ne $null) {
      $this.reader.Close()
    }
  }
  [void] Dispose() {
    $this.Close()
  }

  [String] ReadString() {
    $strSize = $this.DWord()
    if ($strSize -eq 0) {
      return $null
    }
    $rawData = $this.reader.ReadBytes($strSize)
    return $this.encoding.GetString($rawData)
  }
  [Byte[]] ReadImage() {
    $imageSize = $this.DWord()
    if ($imageSize -eq 0) {
      return $null
    }
    return $this.reader.ReadBytes($imageSize)
  }
  [Boolean] ReadBoolean() {
    $byte = $this.Byte()
    $a = $byte -eq 1
    return $a
  }
  [Int32] DWord() {
    return $this.reader.ReadInt32()
  }
  [Int32] Byte() {
    return $this.reader.ReadBytes(1)[0]
  }
}

#$file = "G:\Game\cwnext160_14\CardWirthNext\Scenario\解析\ゴブリンの洞窟\Summary.wsm"
$file = "G:\Game\cwnext160_14\CardWirthNext\Scenario\解析\ゴブリンの洞窟Next\Summary.wsm"
[CardWirthReader] $cwReader = [CardWirthReader]::new()

try {
  
  $cwReader.Open($file)
  $image = $cwReader.ReadImage()
  $title = $cwReader.ReadString()
  $description = $cwReader.ReadString()
  $author = $cwReader.ReadString()
  $requiredCoupons = $cwReader.ReadString()
  $requiredCouponsNum = $cwReader.DWord()

  $areaVersion = $cwReader.DWord()
  if ($areaVersion -le 19999) {
    $version = 0
    $areaId = $areaVersion - 10000
  }
  elseif ($areaVersion -le 39999) {
    $version = 2
    $areaId = $areaVersion - 20000
  }
  elseif ($areaVersion -le 49999) {
    $version = 4
    $areaId = $areaVersion - 40000
  }
  else {
    $version = 7
    $areaId = $areaVersion - 70000
  }

  # データバージョン7(Next)はデータ構造が異なるため以降のデータは読めない

  $stepsNum = $cwReader.DWord()
  $step = @();
  for ($i = 0; $i -lt $stepsNum; $i++) {
    $stepName = $cwReader.ReadString()
    $stepDefault = $cwReader.ReadString()
    $stepValues = @()
    for ($j = 0; $j -lt 10; $j++) {
      $stepValues += $cwReader.ReadString()
    }
    $step += ([PSCustomObject]@{
        Name    = $stepName
        Default = $stepDefault
        Values  = $stepValues
      })
  }

  $flagsNum = $cwReader.DWord()

  $flags = @()
  for ($i = 0; $i -lt $flagsNum; $i++) {
    $flagName = $cwReader.ReadString()
    $flagDefault = $cwReader.ReadBoolean()
    $flagTrueValue = $cwReader.ReadString()
    $flagFalseValue = $cwReader.ReadString()
    $flags += ([PSCustomObject]@{
        Name       = $flagName
        Default    = $flagDefault
        TrueValue  = $flagTrueValue
        FalseValue = $flagFalseValue
      })
  }

  $cwReader.DWord()

  if ($version -gt 0) {
    $levelMin = $cwReader.DWord()
    $levelMax = $cwReader.DWord()
  }
  else {
    $levelMin = 0
    $levelMax = 0
  }

  [PSCustomObject]@{
    title              = $title;
    description        = $description
    author             = $author
    requiredCoupons    = $requiredCoupons
    requiredCouponsNum = $requiredCouponsNum
    areaVersion        = $areaVersion
    version            = $version
    areaId             = $areaId
    stepNum            = $stepsNum
    step               = $step
    flagsNum           = $flagsNum
    flags              = $flags
    levelMin           = $levelMin
    levelMax           = $levelMax
  }

}
catch {
  Write-Host "エラーになった"
}
finally {
  $cwReader.Close()
}

#[System.IO.MemoryStream] $ms = [System.IO.MemoryStream]::new($image)
#[System.Drawing.Bitmap] $bitmap = [System.Drawing.Bitmap]::new($ms)
#$ms.Close()
#$bitmap.Save("")