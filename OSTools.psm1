$messages = DATA {
    # culture="de-de"
    ConvertFrom-StringData @"
        Connecting = Gonna talk to
        Failed = That did not work
"@
}
#Import-LocalizedData -BindingVariable messages
function Get-SysDiskDetails {
    [CmdletBinding()]
    param (
        [Alias("Name","Hostname")]
        [Parameter(
            Mandatory=$true,
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true
        )]
        [string[]]$ComputerName = "localhost"
    )
    PROCESS {
        foreach ($computer in $ComputerName) {

            Write-Verbose "$($messages.Connecting) $computer"

            try{
                $os     = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $computer -ErrorAction Stop
                $cs     = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $computer
                $bios   = Get-WmiObject -Class Win32_BIOS -ComputerName $computer
                $disks  = Get-WmiObject -Class Win32_LogicalDisk -ComputerName $computer

                #First Create Nested Object
                $diskcollection = @()
                foreach ($disk in $disks) {
                    $diskprops = @{
                        'DriveLetter'   =   $disk.deviceid;
                        'DriveType'     =   $disk.DriveType;
                        'Size'          =   $disk.Size
                    }
                    $diskobj        =   New-Object -TypeName psobject -Property $diskprops
                    $diskcollection +=  $diskobj
                }

                $props  = @{
                    'ComputerName'  =   $cs.__SERVER;
                    'OSVersion'     =   $os.version;
                    'SPVersion'     =   $os.servicepackmajorversion;
                    'OSBuild'       =   $os.buildnumber;
                    'Manufacturer'  =   $cs.Manufacturer;
                    'Model'         =   $cs.model;
                    'BIOSSerial'    =   $bios.serialnumber;
                    'Disks'         =   $diskcollection
                }
                $obj = New-Object -TypeName psobject -Property $props
                $obj.psobject.typenames.insert(0,'CriticalSolutions.OSTools.SysDiskDetails')
                Write-Output $obj
            } 
            catch{
                Write-Warning "$($messages.failed) $computer"
            }
        }
    } 
}

function Get-SystemDetails {
    <#
    .SYNOPSIS
        Gets basic system information from one or more computers via WMI
    .DESCRIPTION
        See synopsis. This is not complex.
    .PARAMETER ComputerName
        The name of the Host to Query
    .EXAMPLE
        Get-SystemDetails -Computername DC
        Gets system info from the computer named DC.
    .INPUTS
        Inputs (if any)
    .OUTPUTS
        Output (if any)
    .NOTES
        General notes
    #>
    [CmdletBinding()]
    param (
        [Alias("Name","Hostname")]
        [Parameter(
            Mandatory=$true,
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true
        )]
        [string[]]$ComputerName = "localhost"
    )
    PROCESS {
        foreach ($computer in $ComputerName) {
            $os     = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $computer
            $cs     = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $computer
            $bios   = Get-WmiObject -Class Win32_BIOS -ComputerName $computer

            $props  = @{
                'ComputerName'  =   $cs.__SERVER;
                'OSVersion'     =   $os.version;
                'SPVersion'     =   $os.servicepackmajorversion;
                'OSBuild'       =   $os.buildnumber;
                'Manufacturer'  =   $cs.Manufacturer;
                'Model'         =   $cs.model;
                'BIOSSerial'    =   $bios.serialnumber
            }
            $obj = New-Object -TypeName psobject -Property $props
            $obj.psobject.typenames.insert(0,'CriticalSolutions.OSTools.SystemDetails')
            Write-Output $obj
        }
    }
}#

function Get-DiskDetails {
    <#
    .SYNOPSIS
        Gets information on local disks.
    .DESCRIPTION
        See Synopsis, uses WMI.
    .PARAMETER ComputerName
        One or more Computers or IP's to query.
    .EXAMPLE
        Get-DiskDetails -ComputerName DC,CLIENT
        Get disk space details from Computers DC and CLIENT.
    .INPUTS
        Inputs (if any)
    .OUTPUTS
        Output (if any)
    .NOTES
        General notes
    #>
    [CmdletBinding()]
    param (
        [Alias("Name","Hostname")]
        [Parameter(
            Mandatory=$true,
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true
        )]
        [string[]]$ComputerName = "localhost"
    )
    PROCESS {
        foreach ($computer in $ComputerName) {
            $disks      = Get-WmiObject -Class Win32_LogicalDisk -Filter "DriveType='3'" -ComputerName $computer | Where-Object {$_.size -ne $null}
            $ram        = (Get-WmiObject Win32_PhysicalMemory -ComputerName $computer | Measure-Object -Property capacity -Sum).sum /1gb
            foreach ($disk in $disks) {
                $props  = @{
                    'ComputerName'  =   $computer;
                    'Drive'         =   $disk.DeviceID;
                    'FreeSpace'     =   "{0:N2}" -f ($disk.FreeSpace / 1GB);
                    'Size'          =   "{0:N2}" -f ($disk.size / 1GB);
                    'FreePercent'   =   "{0:N2}" -f ($disk.FreeSpace / $disk.size * 100 -as [int]);
                    'Collected'     =   (Get-Date)
                    'Ram (GB)'      =   $ram
                }
                $obj = New-Object -TypeName psobject -Property $props
                $obj.psobject.typenames.insert(0,'Report.OSTools.DiskDetails')
                Write-Output $obj
            }
        }
    }
} #

function Save-DiskDetailsToDatabase {
    [CmdletBinding()]
    param (
        [Parameter(
            Mandatory=$true,
            ValueFromPipeline=$true
        )]
        [object[]]$inputobject
    )
    BEGIN {
        $connection = New-Object -TypeName System.Data.SqlClient.SqlConnection
        $connection.ConnectionString = "Server=CSNSQLU;Database=DiskDetails;User Id=pshell;Password=Wcdi2010!@#;"
        $connection.Open() | Out-Null
    }
    PROCESS{
        $command = New-Object -TypeName System.Data.SqlClient.SqlCommand
        $command.Connection = $connection
        
        $sql = "DELETE FROM DiskData WHERE ComputerName = '$($inputobject.ComputerName)' AND DriveLetter = '$($inputobject.Drive)'"
        
        Write-Debug "Executing $sql"  
        $command.CommandText = $sql
        $command.ExecuteNonQuery() | Out-Null

        $sql = "INSERT INTO DiskData (
            ComputerName,
            DriveLetter,
            FreeSpace,
            Size,
            FreePercent
            )
            VALUES(
            '$($inputobject.computername)',
            '$($inputobject.Drive)',
            '$($inputobject.FreeSpace)',
            '$($inputobject.Size)',
            '$($inputobject.FreePercent)'
            )"
        Write-Debug "Executing $sql"
        
        $command.CommandText = $sql
        $command.ExecuteNonQuery() | Out-Null
    }
    END {
        $connection.Close()
    }
}

function Get-ComputerNamesForDiskDetailsFromDatabase {
    
    $connection = New-Object -TypeName System.Data.SqlClient.SqlConnection
    $connection.ConnectionString = "Server=CSNSQLU;Database=DiskDetails;User Id=pshell;Password=Wcdi2010!@#;"
    $connection.Open() | Out-Null

    $command = New-Object -TypeName System.Data.SqlClient.SqlCommand
    $command.Connection = $connection
    
    $sql = "SELECT ComputerName FROM DiskData"
    Write-Debug "Executing $sql"
    $command.CommandText = $sql

    $reader = $command.ExecuteReader()

    while ($reader.read()) {
        $computername = $reader.GetSqlString(0).Value
        Write-Output $computername
    }


    $connection.Close()
}

function Set-ComputerState {
    [CmdletBinding(
        SupportsShouldProcess=$true,# Turns on Whatif
        ConfirmImpact='High' # and confirm.
    )]
    param (
        [Parameter(
            Mandatory=$true,
            ValueFromPipeline=$true
        )]
        [string]$ComputerName,

        # Action
        [Parameter(
            Mandatory=$true
        )]
        [ValidateSet(
            'LogOff',
            'PowerOff',
            'Shutdown',
            'Restart'
        )]
        [string]$Action,

        [switch]$Force
    )

    PROCESS{
        foreach ($computer in $ComputerName) {
            
            switch ($Action) {
                'LogOff'    { $x = 0 }
                'Shutdown'  { $x = 1 }
                'Restart'   { $x = 2 }
                'PowerOff'  { $x = 8 }
            }
            if ($Force) {
                $x += 4
            }

            $os = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $computer -EnableAllPrivileges

            if ($pscmdlet.ShouldProcess("$Action $computer")) {
                $os.Win32Shutdown($x) | Out-Null
            }
        }
    }
}

function Convert-WuaResultCodeToName {
    #doug rios modified
    param( [Parameter(Mandatory=$true)]
    [int] $ResultCode
    )
    $Result = $ResultCode
    switch($ResultCode) {
        2 {$Result = "Succeeded"}
        3 {$Result = "Succeeded With Errors"}
        4 {$Result = "Failed"}
    }
    return $Result
}
function Get-WuaHistory {
    
    # Get a WUA Session
    $session = (New-Object -ComObject 'Microsoft.Update.Session')
    
    # Query the latest History starting with the first record
    $history = $session.QueryHistory("",0,100) | `
    
    ForEach-Object {
        $Result = Convert-WuaResultCodeToName -ResultCode $_.ResultCode
        # Make the properties hidden in com properties visible.
        $_ | Add-Member -MemberType NoteProperty -Value $Result -Name Result
        $Product = $_.Categories | Where-Object {$_.Type -eq 'Product'} | Select-Object -First 1 -ExpandProperty Name
        $KBs = $_.Title | Select-String -Pattern "KB\d*"
        $KB = $KBs.Matches.Value
        $_ | Add-Member -MemberType NoteProperty -Value $_.UpdateIdentity.UpdateId -Name UpdateId
        $_ | Add-Member -MemberType NoteProperty -Value $_.UpdateIdentity.RevisionNumber -Name RevisionNumber
        $_ | Add-Member -MemberType NoteProperty -Value $Product -Name Product -PassThru
        $_ | Add-Member -MemberType NoteProperty -Value $KB -Name KBArticle -PassThru
        Write-Output $_
    }
    #Remove null records and only return the fields we want
    $history |
    Where-Object {![String]::IsNullOrWhiteSpace($_.title)} |
    Select-Object Result, Date, Title, KBArticle, SupportUrl, Product, UpdateId, RevisionNumber
}