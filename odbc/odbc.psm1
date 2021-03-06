
Write-Verbose 'Loading odbc.psm1'

Function Get-OdbcConnection {
  <#
  .SYNOPSIS
   OdbcConnectionの接続
  .DESCRIPTION
   System.Data.Odbc.OdbcConnectionを作成し、引数のODBC接続文字列を元に接続します
  .EXAMPLE
   $Connection = Get-OdbcConnection($ConnectionString)
  .PARAMETER ConnectionString
   ODBC接続文字列
  #>
  [CmdletBinding()]
  Param (
    [Parameter(Mandatory=$True)]
      [String]$ConnectionString
  )
  
  Process {
    $Connection = New-Object System.Data.Odbc.OdbcConnection($ConnectionString)
    $Connection.Open()
    return $Connection
  }
}

Function Get-OdbcDataWithConnection {
  <#
  .SYNOPSIS
   SQL(SELECT文)でデータを取得
  .DESCRIPTION
   SQL(SELECT文)でデータを取得します。
   引数に接続済みのOdbcConnectionが必要です。
  .EXAMPLE
   
  .PARAMETER Connection
   OdbcConnection
   接続済みのODBCコネクション
  .PARAMETER Query
   SQL(SELECT文)
  .PARAMETER Parameter
   OdbcParameter
   SQLパラメータ。複数指定可能。
  #>
  [CmdletBinding()]
  Param (
    [Parameter(Mandatory=$True)]
      [System.Data.Odbc.OdbcConnection]$Connection,
    [Parameter(Mandatory=$True,ValueFromPipeline=$True)]
      [String]$Query,
    [Parameter()]
      [System.Data.Odbc.OdbcParameter[]]$Parameter = @()
  )
  
  Process {

    $cmd = New-Object System.Data.Odbc.OdbcCommand
    $cmd.Connection = $Connection
    $cmd.CommandText = $Query
    $Parameter | % {
      $cmd.Parameters.Add($_) | Out-Null
    }
    $da = New-Object System.Data.Odbc.OdbcDataAdapter
    $da.SelectCommand = $cmd
    
    $DataSet = New-Object System.Data.DataSet
    $nRecs = $da.Fill($DataSet)
    $nRecs | Out-Null
    $DataSet.Tables[0]

  }
}

Function Invoke-OdbcQueryWithConnection {
  <#
  .SYNOPSIS
   SQLの実行
  .DESCRIPTION
   SQL文を実行します。
   引数に接続済みのOdbcConnectionが必要です。
  .EXAMPLE
   
  .PARAMETER Connection
   OdbcConnection
   接続済みのODBCコネクション
  .PARAMETER Query
   SQL文
  .PARAMETER Parameter
   OdbcParameter
   SQLパラメータ。複数指定可能。
  .PARAMETER Transaction
   OdbcTransaction
  #>
  [CmdletBinding()]
  Param (
    [Parameter(Mandatory=$True)]
      [System.Data.Odbc.OdbcConnection]$Connection,
    [Parameter(Mandatory=$True,ValueFromPipeline=$True)]
      [String]$Query,
    [Parameter()]
      [System.Data.Odbc.OdbcParameter[]]$Parameter = @(),
    [Parameter()]
      [System.Data.Odbc.OdbcTransaction]$Transaction = $Null
  )
  
  Process {
    $cmd = New-Object System.Data.Odbc.OdbcCommand
    $cmd.Connection = $Connection
    If ($Transaction) {
      $cmd.Transaction = $Transaction
    }
    $cmd.CommandText = $Query
    $Parameter | % {
      $cmd.Parameters.Add($_) | Out-Null
    }
    
    $nRecs = $cmd.ExecuteNonQuery();
    $nRecs
  }  
}

Function Get-OdbcSchemaWithConnection {
  [CmdletBinding()]
  Param (
    [Parameter(Mandatory=$True)]
      [System.Data.Odbc.OdbcConnection]$Connection,
    [Parameter()]
      [String]$Name = "MetaDataCollections",
    [Parameter()]
      [Array]$Parameter = @()
  )
  
  Process {
    $DataTable = $Connection.GetSchema($Name, $Parameter)
    $DataTable
  }
}

Function Get-OdbcData {
  <#
  .SYNOPSIS
   SQL(SELECT文)でデータを取得
  .DESCRIPTION
   SQL(SELECT文)でデータを取得します。
  .EXAMPLE
   
  .PARAMETER ConnectionString
   ODBC接続文字列
  .PARAMETER Query
   SQL(SELECT文)
  .PARAMETER Parameter
   OdbcParameter
   SQLパラメータ。複数指定可能。
  #>
  [CmdletBinding()]
  Param (
    [Parameter(Mandatory=$True)]
      [String]$ConnectionString,
    [Parameter(Mandatory=$True,ValueFromPipeline=$True)]
      [String]$Query,
    [Parameter()]
      [System.Data.Odbc.OdbcParameter[]]$Parameter = @()
  )

  Process {
    $Connection = $null
    try{
      $Connection = Get-OdbcConnection($ConnectionString)
    }catch{
      Write-Error("接続に失敗しました。接続文字列を確認してください。")
      Write-Error($_.Exception)
      Exit 1
    }
    
    try{
    
      Get-OdbcDataWithConnection -Connection $Connection -Query $Query -Parameter $Parameter
      
    }catch{
      Write-Error("データの取得に失敗しました。")
      Write-Error($_.Exception)
    } finally {
      try {
        $Connection.Close()
      } catch {
      }
    }
  }
}

Function Invoke-OdbcQuery {
  <#
  .SYNOPSIS
   SQL(SELECT文)でデータを取得
  .DESCRIPTION
   SQL文を実行します。
   引数のODBC接続文字列を使用して１回だけ接続が行われ、
   SQL実行後は即コミットされます。
  .EXAMPLE
   
  .PARAMETER ConnectionString
   ODBC接続文字列
  .PARAMETER Query
   SQL文
  .PARAMETER Parameter
   OdbcParameter
   SQLパラメータ。複数指定可能。
  #>
  [CmdletBinding()]
  Param (
    [Parameter(Mandatory=$True)]
      [String]$ConnectionString,
    [Parameter(Mandatory=$True,ValueFromPipeline=$True)]
      [String]$Query,
    [Parameter()]
      [System.Data.Odbc.OdbcParameter[]]$Parameter = @()
  )

  Process {
    $Connection = $null
    try{
      $Connection = Get-OdbcConnection($ConnectionString)
    }catch{
      Write-Error("接続に失敗しました。接続文字列を確認してください。")
      Write-Error($_.Exception)
      Exit 1
    }

    try {
    
      $nRecs = Invoke-OdbcQueryWithConnection -Connection $Connection -Query $Query -Parameter $Parameter
      $nRecs
    
    } catch {
      Write-Error("SQLの実行に失敗しました。")
      Write-Error($_.Exception)
    } finally {
      try {
        $Connection.Close()
      } catch {
      }
    }
  }
}

Function Make-OdbcParameter {
  <#
  .SYNOPSIS
   OdbcParameterの作成
  .DESCRIPTION
   OdbcParameterを作成します。
  .EXAMPLE
   
  .PARAMETER Name
   パラメータの名前
  .PARAMETER Type
   パラメータの型
  .PARAMETER Value
   パラメータの値
  #>
  [CmdletBinding()]
  Param(
    [Parameter(Mandatory=$True)]
      [String]$Name,
    [Parameter(Mandatory=$True)]
      [System.Object]$Type,
    [Parameter(Mandatory=$True)]
      [System.Object]$Value
  )
  Process {
    $param = New-Object System.Data.Odbc.OdbcParameter($Name, $Type)
    $param.Value = $Value
    return $param
  }
}
