<#
.SYNOPSIS

Configuration as Code - Azure DevOps Scafold and Pipeline
Copyright (C) 2021  Sjoerd de Valk

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License version 3 as published by
the Free Software Foundation.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.  If not, see <https://www.gnu.org/licenses/>.
All helper functions for use by the default Powershell 5.1+ implementation are loaded here, to be used by a main function
.DESCRIPTION

Do not adjust any function. Just add new ones. Clean-up later.
.EXAMPLE

Load this script file as follows: . .\Powershell-HelperFunctions.ps1
#>
if ($IsWindows) {
  $global:folderseparator = "\"
}
else {
  $global:folderseparator = "/"
}

<#
Reference: https://gist.githubusercontent.com/alexbevi/34b700ff7c7c53c7780b/raw/8925255eb7be0cf4db180b79b86a315b1ca1077c/Execute-With-Retry.ps1
This function can be used to pass a ScriptBlock (closure) to be executed and returned.
The operation retried a few times on failure, and if the maximum threshold is surpassed, the operation fails completely.
Params:
    Command         - The ScriptBlock to be executed
    RetryDelay      - Number (in seconds) to wait between retries
                      (default: 5)
    MaxRetries      - Number of times to retry before accepting failure
                      (default: 5)
    VerboseOutput   - More info about internal processing
                      (default: false)
Examples:
Start-RetryScriptBlock { $connection.Open() }
$result = Start-RetryScriptBlock -ScriptBlock { $command.ExecuteReader() } -SecondsDelay 1 -Retries 2
#>
function Start-RetryScriptBlock {
  [CmdletBinding()]
  param(
    [Parameter(ValueFromPipeline, Mandatory)]
    $ScriptBlock,
    $SecondsDelay = 5,
    $Retries = 1,
    [array]$ExcludeErrorMatches = @(),
    $VerboseOutput = $true,
    $Indent = ""
  )

  $currentRetry = 0
  $success = $false
  $cmd = $ScriptBlock.ToString()

  do {
    try {
      $result = . $ScriptBlock
      $success = $true
      if ($VerboseOutput -eq $true) {
        $write = Write-Host "$($Indent)Successfully executed [$cmd]" -ForegroundColor "Green"
      }

      $Error.Clear()
      return $result
    }
    catch [System.Exception] {
      $currentRetry = $currentRetry + 1
      $errorexcluded = $false

      foreach ($ExcludeErrorMatch in $ExcludeErrorMatches) {
        if ($_.ToString() -match $ExcludeErrorMatch) {
          $errorexcluded = $true
        }
      }

      if (!$errorexcluded) {
        if ($VerboseOutput -eq $true) {
          $write = Write-Host "$($Indent)Failed to execute [$cmd]: $($_.ToString())" -ForegroundColor "Red"
        }

        if ($currentRetry -gt $Retries) {
          if ($global:stacktracemode) {
            $write = Write-Host "$($Indent)Retry limit exceeded. Could not execute [$cmd]. The error: $($_.FullyQualifiedErrorId) || $($_.ScriptStackTrace) || $($_.Exception.ToString()) || $($_.ToString())" -ForegroundColor "Red"
            throw "Retry limit exceeded. Could not execute [$cmd]. The error: $($_.FullyQualifiedErrorId) || $($_.ScriptStackTrace) || $($_.Exception.ToString()) || $($_.ToString())"
          }
          else {
            $write = Write-Host "$($Indent)Retry limit exceeded. Could not execute [$cmd]. The error: $($_.ToString())" -ForegroundColor "Red"
            throw "Retry limit exceeded. Could not execute [$cmd]. The error: " + $_.ToString()
          }
        }
        else {
          if ($VerboseOutput -eq $true) {
            $write = Write-Host "$($Indent)Waiting $SecondsDelay second(s) before attempt #$currentRetry of [$cmd]" -ForegroundColor "Yellow"
          }
          Start-Sleep -s $SecondsDelay
        }
      }
      else {
        $Error.Clear()
        $success = $true
      }
    }
  } while (!$success);
}


Function ConvertFrom-Cli {
  [CmdletBinding()]
  param (
    [Parameter(ValueFromPipeline)] [string] $line
  )
  begin {
    # Collect all lines in the input
    $lines = @()
  }

  process {
    # 'process' is run once for each line in the input pipeline.
    $lines += $line
  }

  end {
    # Azure Cli errors and warnings change output colors permanently.
    # Reset the shell colors after each operation to keep consistent.
    [Console]::ResetColor()

    # If the 'az' process exited with a non-zero exit code we have an error.
    # The 'az' error message is already printed to console, and is not a part of the input.
    if ($LASTEXITCODE) {
      Write-Error "az exited with exit code $LASTEXITCODE" -ErrorAction 'Stop'
    }

    $inputJson = $([string]::Join("`n", $lines));
    # We expect a Json result from az cli if we have no error. The json result CAN be $null.
    $result = ConvertFrom-Json $inputJson
    return $result
  }
}
