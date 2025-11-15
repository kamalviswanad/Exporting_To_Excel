function Export-InventoryToExcel {
    [CmdletBinding(DefaultParameterSetName='Path')]
    param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [PSCustomObject[]]$InputObject,

        [Parameter(Mandatory=$true)]
        [string]$Path,
    )

    begin {
        # Check if the required module is available
        if (-not (Get-Module -Name ImportExcel -ListAvailable)) {
            Write-Error "The 'ImportExcel' module is required but not installed. Please run: Install-Module -Name ImportExcel"
            break
        }
    }

    process {
        $groupedData = $InputObject | Group-Object -Property DataType
        foreach ($group in $groupedData) {
            $currentSheet = $group.Name
            Write-Host "Exporting $($group.Count) records for data type '$currentSheet'..."
            $group.Group | Export-Excel -Path $Path -WorksheetName $currentSheet -AutoFilter -AutoSize -Append 
        }
    }

    end {
        Write-Host "Successfully exported inventory data to: $Path"
    }
}



function Get-ServerInventory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string[]]$ComputerName,

        [Parameter(Mandatory=$false)]
        [System.Management.Automation.PSCredential]$Credential = [System.Management.Automation.PSCredential]::Empty
    )

    foreach ($server in $ComputerName) {
        try {
            
            $remoteData = Invoke-Command -ComputerName $server -Credential $Credential -ScriptBlock {
                
            $command = @(Get-Process | Sort-Object CPU | Select-Object -First 3)
 
    if($command.count -gt 0){
        for($i=0;$i -lt $command.count;$i++){
                [PSCustomObject]@{
                    index          = $i
                    PSComputerName = $env:COMPUTERNAME
                    DataType       = "ProcessName"
                    Value          = $command[$i].ProcessName
                    CPU         = $command[$i].CPU
                    Timestamp      = (Get-Date)
                                 }
                     }
            }else{
                Write-Host "the number is equal or less than 0"
                 }   
            }

            # Return the output
            Write-Output $remoteData

        } catch {
            Write-Error "Error connecting to or running command on $server : $($_.Exception.Message)"
        }
    }
}


function FinalOutput{
Get-ServerInventory -ComputerName @("Server1","Server2","Server3") | Select-Object PSComputerName, DataType, Value, SourceKey, Timestamp | Export-InventoryToExcel -Path "Filepath"
}
