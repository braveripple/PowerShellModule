#Requires -Version 2.0

Write-Verbose 'Loading odbc.psm1'

Function Get-OdbcConnection {
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

Function Get-OdbcDataSet {
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
    return $DataSet

  }
}

Function Execute-OdbcQuery {
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

# SELECT文を実行して結果を返す
Function Get-OdbcData {
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

    $DataSet = Get-OdbcDataSet -Connection $Connection -Query $Query -Parameter $Parameter
    $DataSet.Tables[0]

  }
}

# INSERT/UPDATE/DELETE文の実行、即コミット
Function Invoke-OdbcQuery {
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

    $nRecs = Execute-OdbcQuery -Connection $Connection -Query $Query -Parameter $Parameter
    $nRecs
  }
}

Function Get-OdbcSchema {
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
    $DataTable = $Connection.GetSchema($Name,$Parameter)
    return $DataTable

  }
}

Function Make-OdbcParameter {
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
#New-Alias -Name dbbb -Value Invoke-SelectQuery -Description "quick wmi alias" -Option ReadOnly
