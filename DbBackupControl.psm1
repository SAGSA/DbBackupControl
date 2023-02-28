function Remove-DbBackup{
<#
.SYNOPSIS
    Removes MSSQL and other sql Server backups from disk or ftp.
.DESCRIPTION
    Provides wide functionality to delete SQL backups from disk or ftp. Implemented functionality to block deletion of .bak files if .diff or .trn depend on them 
.PARAMETER Path
    Specifies the name of the base level folder to search for backup files. The search for backup files will be recursive from this location if the Deph parameter is greater than 0.
.PARAMETER DbName
    Specify databases to be deleted
.PARAMETER KeepVersions
    Specifies the number of the most recent backups to keep.
.PARAMETER KeepVersionsDiff
    Specifies the number of most recent differential MSSQL backup copies to keep. File extension must be .diff
.PARAMETER KeepVersionsTrn
    Specifies the number of most recent MSSQL transaction log backups to keep. File extension must be .trn
.PARAMETER KeepVersionsWeekly
    Specifies the number of the most recent weekly backups to keep.
.PARAMETER DayOfWeek
    Specify the day of the week. Only backups made on the specified day will be saved. By default sunday.
.PARAMETER KeepVersionsMonthly
    Specifies the number of the most recent monthly backups to keep.
.PARAMETER DayOfMonth
    Specify the day of the month. Only backups made on the specified day will be saved. By default 28.
.PARAMETER KeepVersionsYearly
    Specifies the number of the most recent yearly backups to keep.
.PARAMETER DayOfYear
    Specify the day of the year. Only backups made on the specified day will be saved. By default 365.
.PARAMETER CheckArchiveBit
    If this switch is enabled, the filesystem Archive bit is checked before deletion. If this bit is set (which translates to "it has not been backed up to another location yet", the file won't be deleted.
    Doesn't matter for files that are on the ftp server
.PARAMETER FtpCredential
    Specify ftp credentials if the backup files are located on ftp.

.EXAMPLE
    Remove-DbBackup -Path "C:\MSSQL\BACKUP"
    All backup files in "C:\MSSQL\BACKUP" will be removed. Files in the subdirectory will not be affected
.EXAMPLE
    Remove-DbBackup -Path 'C:\MSSQL\BACKUP' -DbName "db1","db2" -KeepVersions 7
    Only the 7 most recent versions of db1 and db2 will be kept
.EXAMPLE
    Remove-DbBackup -Path 'C:\MSSQL\BACKUP' -DbName "db1","db2" -KeepVersions 7 -WhatIf
    Same as example #2, but doesn't actually remove any files. The function will instead show you what would be done.
    This is useful when first experimenting with using the function.
.EXAMPLE
    Remove-DbBackup -Path 'C:\MSSQL\BACKUP' -DbName "db1","db2" -KeepVersions 7 -KeepVersionsWeekly 4 -DayOfWeek Friday,Monday
    The 7 most recent versions of db1 and db2 will be kept. The last 4 copies made on Monday and Friday will also be saved.
.EXAMPLE    
    $FtpCredential=Get-Credential
    Remove-DbBackup -Path 'C:\MSSQL\BACKUP',"ftp://MSSQL/BACKUP" -KeepVersions 30 -KeepVersionsWeekly 4 -KeepVersionsMonthly 12 -KeepVersionsYearly 5 -Deph 3 -FtpCredential $FtpCredential
    After executing this command, the following will be saved:
        - 30 most recent backups
        - 4  latest weekly backups (by default day of week is sunday)
        - 12 last monthly backups (by default day of month is 28)
        - 5  last yearly  backups (by default day of year is 365)
.EXAMPLE
    Remove-DbBackup -Path 'C:\MSSQL\BACKUP' -KeepVersionsMonthly 12 -DayOfMonth 1,28 -KeepVersionsYearly 6 -DayOfYear 1,365
    After executing this command, the following will be saved:
    - 12 last monthly backups made 1 and 28 day of month
    - 6 last yearly bakups made 1 and 365 day of year
.EXAMPLE
    Remove-DbBackup -Path mega: -Deph 3 -KeepVersions 3
    Only the 3 latest versions of the database copy on mega cloud storage will be kept. You need to download rclone to the C:\Windows\PsScript folder and configure it. https://rclone.org/
.NOTES
    Author: SAGSA
    https://github.com/SAGSA/DbBackupControl
    Requires: Powershell 2.0
#>    
    [cmdletbinding(SupportsShouldProcess=$true)]
    param(
        [ValidateScript({($_ -match "^ftp://") -or $_ -match "^\w+:"})]
        [string[]]$Path,
        [ValidateScript({-not ($_ -match "\s+")})]
        [string[]]$DbName,
        [ValidateScript({$_ -ge 1})]
        [int]$KeepVersions,
        [ValidateScript({$_ -ge 1})]
        [int]$KeepVersionsDiff,
        [ValidateScript({$_ -ge 1})]
        [int]$KeepVersionsTrn,
        [ValidateScript({$_ -ge 1})]
        [int]$KeepVersionsWeekly,
        [ValidateSet("Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday")]
        [string[]]$DayOfWeek,
        [ValidateScript({$_ -ge 1})]
        [int]$KeepVersionsMonthly,
        [ValidateRange(1,31)]
        [int[]]$DayOfMonth,
        [ValidateScript({$_ -ge 1})]
        [int]$KeepVersionsYearly,
        [ValidateRange(1,365)]
        [int[]]$DayOfYear,
        [switch]$CheckArchiveBit,
        [ValidateRange(1,5)]
        [int]$Deph,
        [string]$JsonConfigPath,
        $FtpCredential


    )

        if ($PSBoundParameters["JsonConfigPath"] -eq $null -and $PSBoundParameters["Path"] -eq $null){
            Write-Error "Parameters Path or JsonConfigPath are required" -ErrorAction Stop
        }
        if ($PSBoundParameters["JsonConfigPath"]){
            if($PSBoundParameters.Keys.Count -gt 1){
                Write-Error "Only the JsonConfigPath option can be accepted" -ErrorAction Stop
            }
        }
        if($PSBoundParameters["DayOfWeek"] -or $PSBoundParameters["KeepVersionsWeekly"]){
            if(-not $PSBoundParameters["KeepVersionsWeekly"]){
                Write-Error "Parameter KeepVersionsWeekly must be present" -ErrorAction Stop
            }
            if ($PSBoundParameters["KeepVersionsWeekly"] -and (-not $PSBoundParameters["DayOfWeek"])){
                $PSBoundParameters["DayOfWeek"]=[string[]]"Sunday"
            }
        }
        if($PSBoundParameters["DayOfMonth"] -or $PSBoundParameters["KeepVersionsMonthly"]){
            if(-not $PSBoundParameters["KeepVersionsMonthly"]){
                Write-Error "Parameter KeepVersionsMonthly must be present" -ErrorAction Stop
            }
            if ($PSBoundParameters["KeepVersionsMonthly"] -and (-not $PSBoundParameters["DayOfMonth"])){
                $PSBoundParameters["DayOfMonth"]=[int[]]28
            }
        }
        if($PSBoundParameters["DayOfYear"] -or $PSBoundParameters["KeepVersionsYearly"]){
            if(-not $PSBoundParameters["KeepVersionsYearly"]){
                Write-Error "Parameter KeepVersionsYearly must be present" -ErrorAction Stop
            }
            if ($PSBoundParameters["KeepVersionsYearly"] -and (-not $PSBoundParameters["DayOfYear"])){
                $PSBoundParameters["DayOfYear"]=[int[]]365
            }
        }
        if ($Path -match "ftp://"){
            if(-not $PSBoundParameters["FtpCredential"]){
                Write-Error "Parameter FtpCredential must be present" -ErrorAction Stop
            }
        }
        if ($PSBoundParameters["JsonConfigPath"]){
            Write-Error "This functionality is not ready yet. Install a new version of the module and try again" -ErrorAction Stop
        }
        $KeepConfig="KeepVersions","KeepVersionsWeekly","DayOfWeek","KeepVersionsMonthly","DayOfMonth","KeepVersionsYearly","DayOfYear","KeepVersionsDiff","KeepVersionsTrn","CheckArchiveBit"
        $Parameters=$PSBoundParameters
        $AllKeepSettings=@()
        $DefaultKeepSettings=$null
        
        if ($Parameters["Dbname"] -eq $null){
            $BaseObjConfig=New-Object -TypeName psobject 
            $BaseObjConfig | Add-Member -MemberType NoteProperty -Name BaseName -Value "DefaultSettings"
                $Parameters.keys | Where-Object {$KeepConfig -match $_}  | foreach {
                $BaseObjConfig | Add-Member -MemberType NoteProperty -Name $_ -Value $Parameters[$_]  
            }
            $BaseObjConfig | Add-Member -MemberType NoteProperty -Name IsDefaultSettings -Value $true 
            $DefaultKeepSettings=$BaseObjConfig
        } 
        else{
            $Parameters["DbName"] | foreach{
                    $BaseName=$_
                    $BaseObjConfig=New-Object -TypeName psobject 
                    $BaseObjConfig | Add-Member -MemberType NoteProperty -Name BaseName -Value $BaseName
                    $Parameters.Keys | Where-Object {$KeepConfig -match $_} | foreach {
                        $Key=$_
                        $BaseObjConfig | Add-Member -MemberType NoteProperty -Name $Key -Value $Parameters[$Key]
                    
                    }
                $BaseObjConfig | Add-Member -MemberType NoteProperty -Name IsDefaultSettings -Value $False   
                $AllKeepSettings+=$BaseObjConfig
            }  
        }
        
    
    RemoveOldDumps -DumpPaths $Path -AllKeepSettings $AllKeepSettings -DefaultKeepSettings $DefaultKeepSettings -RecurseDeph $Deph -Credential $FtpCredential -WhatIf:$([bool]$WhatIfPreference.IsPresent)
    
    
}
function StartRcloneApiServer{
    [cmdletbinding()]
    param(
        [string]$RclonePath="$env:SystemRoot\psscript\rclone.exe",
        [parameter(Mandatory=$true)]
        $RcloneCredential
    )
    try{
        
        if(-not (Test-Path -Path $RclonePath)){
            Write-Error "Incorrect path $RclonePath Rclone Not found" -ErrorAction Stop
        }
        $WRcloneCommandLine="'"+'"'+$($RclonePath -replace "\\","\\")+'"'+" rcd%"+"'"
        Get-WmiObject -Query "Select * FROM win32_process WHERE CommandLine like $WRcloneCommandLine" | foreach{
                Stop-Process -Id $_.ProcessId -Force -ErrorAction Stop
        }
        if($Global:RcloneServer -ne $null){
            $Global:RcloneServer.PowerShell.Dispose()
            Remove-Variable -Scope Global -Name RcloneServer -WhatIf:$false   
        }
        $RcloneUser=$RcloneCredential.UserName
        $RclonePassword=GetPlainTextPassword -SecString $RcloneCredential.Password
        $SessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
        Get-Command -CommandType Function -Name InvokeExe | foreach {
            $SessionStateFunction = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList $_.name, $_.Definition         
            Write-Verbose "Add script Function $($_.name)"
            $SessionState.Commands.Add($SessionStateFunction)
                
        }
        $RunspacePool = [runspacefactory]::CreateRunspacePool(1,1,$SessionState,$Host)
        $RunspacePool.Open()

        [ScriptBlock]$SbRunspace={
        param(
            [string]$FilePath,
            [string]$RcloneUser,
            [string]$RclonePassword
        )

            [string]$RcPassParam=$("--rc-pass="+"$RclonePassword")
            $RcloneArgs=@(
                "rcd",
                "--rc-user=$RcloneUser",
                "--rc-pass=$RclonePassword"
            )
            
            $Res=InvokeExe -ExeFile $FilePath -Args $RcloneArgs
            $Res
        }
        $PowerShell = [powershell]::Create()
        [void]$PowerShell.AddScript($SbRunspace)
        $ParamList=@{}
        $ParamList.Add("FilePath",$(get-variable -Name RclonePath -ValueOnly))
        $ParamList.Add("RcloneUser",$(get-variable -Name RcloneUser -ValueOnly))
        $ParamList.Add("RclonePassword",$(get-variable -Name RclonePassword -ValueOnly))
        [void]$PowerShell.AddParameters($ParamList)
        $PowerShell.Runspacepool = $RunspacePool
        $State = $PowerShell.BeginInvoke()
        $temp = '' | Select PowerShell,State,StartTime
        $temp.powershell=$PowerShell
        $temp.state=$State
        $temp.StartTime=get-date
    
        New-Variable -Scope Global -Name RcloneServer -Value $temp -WhatIf:$false | Out-Null
    }
    catch{
        Write-Error $_
    }
    
    
}
function Remove-ItemRclone{
    [cmdletbinding(SupportsShouldProcess=$true)]
    param(
        [parameter(Mandatory=$true)]
        [string]$Path,
        [string]$RclonePath="$env:SystemRoot\psscript\rclone.exe",
        [parameter(Mandatory=$true)]
        $RcloneCredential
    )
    $RcloneArgs=@(
        "lsjson",
        $Path   
    )
    
    if($Path -match "(.+:)(.*)$"){
        $Fs=$Matches[1]
        $Remote=$Matches[2]
    }
    else{
        Write-Error "Incorrect path $Path" -ErrorAction Stop
    }
    $ApiParam='{"fs": '+'"'+$Fs+'", "remote": '+'"'+$Remote+'"'+'}'
    if ($RcloneServer.State.IsCompleted -ne $false){
        #StartRcloneApiServer -ErrorAction Stop
        if ($RcloneServer.State.IsCompleted -ne $false){
            Write-Error "Rclone Server api not working" -ErrorAction Stop
        }
        
    }
    if ($([bool]$WhatIfPreference.IsPresent) -eq $false){
        Write-Verbose "Delete file $Path"
        $Res=Invoke-WebRequest -Uri "http://localhost:5572/operations/deletefile" -Method Post -Body $ApiParam -UseBasicParsing -ContentType "application/json" -Credential $RcloneCredential -ErrorAction Stop
        if ($Res.StatusCode -ne 200){
            Write-Error "Invoke-WebRequest error $($Res.StatusCode)" -ErrorAction Stop
        }
    }
    else{
        Write-Host "WhatIf: Remove-ItemRclone -Path $Path"
    }
    
}
function Get-ChildItemRclone{
    [cmdletbinding()]
    param(
        [parameter(Mandatory=$true)]
        [string]$Path,
        [string]$RclonePath="$env:SystemRoot\psscript\rclone.exe",
        [parameter(Mandatory=$true)]
        $RcloneCredential
    )
    $RcloneArgs=@(
        "lsjson",
        $Path   
    )
    
    if($Path -match "(.+:)(.*)$"){
        $Fs=$Matches[1]
        $Remote=$Matches[2]
    }
    else{
        Write-Error "Incorrect path $Path" -ErrorAction Stop
    }
    $ApiParam='{"fs": '+'"'+$Fs+'", "remote": '+'"'+$Remote+'"}'
    if ($RcloneServer.State.IsCompleted -ne $false){
        #StartRcloneApiServer -ErrorAction Stop
        if ($RcloneServer.State.IsCompleted -ne $false){
            Write-Error "Rclone Server api not working" -ErrorAction Stop
        }
        
    }
    $Res=Invoke-WebRequest -Uri "http://localhost:5572/operations/list" -Method Post -Body $ApiParam -UseBasicParsing -ContentType "application/json" -Credential $RcloneCredential -ErrorAction Stop
    if ($Res.StatusCode -ne 200){
        Write-Error "Invoke-WebRequest error $($Res.StatusCode)" -ErrorAction Stop
    }
    $ItemsInfo=($Res.content | ConvertFrom-Json).list
    $ItemsInfo | foreach {
        $Path=$_.Path
        $Name=$_.name
        $IsContainer=$_.IsDir
        $FullName=$Fs+$Path
        $Item=New-Object -TypeName psobject -Property @{
                Name=$Name
                FullName=$FullName
                PSIsContainer=$IsContainer
        }
        $Item
    }
    
    <#$Res=InvokeExe -ExeFile $RclonePath -Args $RcloneArgs -Encoding 65001
    if ($Res.exitcode -ne 0){
        Write-Error "StdErr: $($Res.StdErr) StdOut: $($Res.StOut)" -ErrorAction Stop
    }
    $Items=$Res.stdout
    if(-not ($Res.stdout -match "]$")){
        $Items=($Res.stdout -replace "]")+"]"
    }
    $ItemsInfo=$Items | ConvertFrom-Json 
    $ItemsInfo | foreach {
        $Name=$_.name
        $IsContainer=$_.IsDir
        if($Path -match ".+:$"){
            $FullName="$Path"+$Name
        }
        elseif($Path -match "$Name$"){
            $FullName=$Path
        }
        else{
            $FullName="$Path\"+$Name
        }
        
        
        $Item=New-Object -TypeName psobject -Property @{
                Name=$Name
                FullName=$FullName
                PSIsContainer=$IsContainer
        }
        $Item
    }#>
}
function New-FakeBackup{
<#
.SYNOPSIS
    Creates the specified number of false backups
.DESCRIPTION
    Creates the specified number of false backups 
    This is useful when first experimenting with using the function.
.EXAMPLE
    New-FakeBackup -Path "C:\FakeDump" -BaseName test1 -FakeCount 1200 
    This command will create 1200 false backups in the "C:\FakeDump" folder

.NOTES
    Author: SAGSA
    https://github.com/SAGSA/DbBackupControl
    Requires: Powershell 2.0
#>
    [CmdletBinding()]
    param(
        [parameter(Mandatory=$true)]
        [ValidateScript({$_ -match "^\w:\\"})]
        [string]$Path,
        [parameter(Mandatory=$true)]
        [string]$BaseName,
        [parameter(Mandatory=$true)]
        [int]$FakeCount,
        [ValidateRange(1,6)]
        [int]$DiffCountDay,
        [ValidateRange(1,24)]
        [int]$LogCountDay
    )
    $CurrentDate=Get-Date
    function CreateFakeFile{
        param(
            $Date,
            [string]$BaseName,
            [string]$Extension="backup"
        )
        $FileDate=Get-Date $Date -Format yyyy_MM_dd_HHmmss
        $FileName=$BaseName+"_backup_"+$FileDate+".$Extension"
        $FilePath=Join-Path -Path $Path -ChildPath $FileName
        New-Item -ItemType File -Path $FilePath
    }
    
    1..$FakeCount | foreach {
        $FakeCount-=1
        
        $FakeDate=(Get-Date -Hour 0 -Minute 0 -Second 0 -Day $CurrentDate.Day -Month $CurrentDate.Month -Year $CurrentDate.Year).AddDays(-$FakeCount) 
        CreateFakeFile -BaseName $BaseName -Date $FakeDate
        
        [int]$DiffCount=0
        if ($DiffCountDay -ge 1){
            
            1..$DiffCountDay | foreach {
                $DiffCount+=18
                $FakeDateDiff=$FakeDate.AddHours($DiffCount) 
                CreateFakeFile -BaseName $BaseName -Date $FakeDateDiff -Extension "diff"
            }
        }
        [int]$LogCount=0
        if ($LogCountDay -ge 1){
            
            1..$LogCountDay | foreach {
                $LogCount+=1
                $FakeDatelog=$FakeDate.AddHours($LogCount) 
                CreateFakeFile -BaseName $BaseName -Date $FakeDatelog -Extension "trn"
            }
        }
        
    }
    
     
}
function Create-RemoveDbBackupJsonConfig{
    [cmdletbinding(SupportsShouldProcess=$true)]
    param(
        [parameter(Mandatory=$true)]
        [ValidateScript({($_ -match "^ftp://") -or $_ -match "^\w:\\"})]
        [string[]]$Path,
        [ValidateScript({-not ($_ -match "\s+")})]
        [string[]]$DbName,
        [ValidateScript({$_ -ge 1})]
        [int]$KeepVersions,
        [ValidateScript({$_ -ge 1})]
        [int]$KeepVersionsDiff,
        [ValidateScript({$_ -ge 1})]
        [int]$KeepVersionsTrn,
        [ValidateScript({$_ -ge 1})]
        [int]$KeepVersionsWeekly,
        [ValidateSet("Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday")]
        [string[]]$DayOfWeek,
        [ValidateScript({$_ -ge 1})]
        [int]$KeepVersionsMonthly,
        [ValidateRange(1,31)]
        [int[]]$DayOfMonth,
        [ValidateScript({$_ -ge 1})]
        [int]$KeepVersionsYearly,
        [ValidateRange(1,365)]
        [int[]]$DayOfYear,
        [switch]$CheckArchiveBit,
        [ValidateRange(1,5)]
        [int]$Deph,
        [parameter(Mandatory=$true)]
        [string]$JsonConfigPath,
        $FtpCredential


    )

        if($PSBoundParameters["DayOfWeek"] -or $PSBoundParameters["KeepVersionsWeekly"]){
            if(-not $PSBoundParameters["KeepVersionsWeekly"]){
                Write-Error "Parameter KeepVersionsWeekly must be present" -ErrorAction Stop
            }
            if ($PSBoundParameters["KeepVersionsWeekly"] -and (-not $PSBoundParameters["DayOfWeek"])){
                $PSBoundParameters["DayOfWeek"]=[string[]]"Sunday"
            }
        }
        if($PSBoundParameters["DayOfMonth"] -or $PSBoundParameters["KeepVersionsMonthly"]){
            if(-not $PSBoundParameters["KeepVersionsMonthly"]){
                Write-Error "Parameter KeepVersionsMonthly must be present" -ErrorAction Stop
            }
            if ($PSBoundParameters["KeepVersionsMonthly"] -and (-not $PSBoundParameters["DayOfMonth"])){
                $PSBoundParameters["DayOfMonth"]=[int[]]28
            }
        }
        if($PSBoundParameters["DayOfYear"] -or $PSBoundParameters["KeepVersionsYearly"]){
            if(-not $PSBoundParameters["KeepVersionsYearly"]){
                Write-Error "Parameter KeepVersionsYearly must be present" -ErrorAction Stop
            }
            if ($PSBoundParameters["KeepVersionsYearly"] -and (-not $PSBoundParameters["DayOfYear"])){
                $PSBoundParameters["DayOfYear"]=[int[]]365
            }
        }
        if ($Path -match "ftp://"){
            if(-not $PSBoundParameters["FtpCredential"]){
                Write-Error "Parameter FtpCredential must be present" -ErrorAction Stop
            }
        }

        $KeepConfig="KeepVersions","KeepVersionsWeekly","DayOfWeek","KeepVersionsMonthly","DayOfMonth","KeepVersionsYearly","DayOfYear","KeepVersionsDiff","KeepVersionsTrn"
        $Parameters=$PSBoundParameters
        $AllKeepSettings=@()
        $DefaultKeepSettings=$null
        
        if ($Parameters["Dbname"] -eq $null){
            $BaseObjConfig=New-Object -TypeName psobject 
            $BaseObjConfig | Add-Member -MemberType NoteProperty -Name BaseName -Value "DefaultSettings"
                $Parameters.keys.Where({$KeepConfig -match $_})  | foreach {
                $BaseObjConfig | Add-Member -MemberType NoteProperty -Name $_ -Value $Parameters[$_]  
            }
            $BaseObjConfig | Add-Member -MemberType NoteProperty -Name IsDefaultSettings -Value $true 
            $DefaultKeepSettings=$BaseObjConfig
        } 
        else{
            $Parameters["DbName"] | foreach{
                    $BaseName=$_
                    $BaseObjConfig=New-Object -TypeName psobject 
                    $BaseObjConfig | Add-Member -MemberType NoteProperty -Name BaseName -Value $BaseName
                    $Parameters.Keys.Where({$KeepConfig -match $_}) | foreach {
                        $Key=$_
                        $BaseObjConfig | Add-Member -MemberType NoteProperty -Name $Key -Value $Parameters[$Key]
                    
                    }
                $BaseObjConfig | Add-Member -MemberType NoteProperty -Name IsDefaultSettings -Value $False   
                $AllKeepSettings+=$BaseObjConfig
            }  
        }
        
    
    CreateJsonConfig -FilePath $JsonConfigPath -DumpPaths $Path -AllKeepSettings $AllKeepSettings -DefaultKeepSettings $DefaultKeepSettings -RecurseDeph $Deph -FtpCredential $FtpCredential
   
    
    
}
function InvokeExe{
    <#
    .NOTES
        Author: SAGSA
        https://github.com/SAGSA/PostgresCmdlets
        Requires: Powershell 2.0
    #>
    [cmdletbinding()]
        param(
            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [String]$ExeFile,
            [Parameter(Mandatory=$false)]
            [String[]]$Args,
            [hashtable]$EnvVar,
            [switch]$VerboseOutput,
            [string]$LogPath,
            [Parameter(Mandatory=$false)]
            [String]$Verb,
            [int]$Encoding
        )    
        if (!([string]::IsNullOrEmpty($PSBoundParameters["LogPath"])))
        {
            New-Item -ItemType File -Path $LogPath -ErrorAction Stop -Force | Out-Null
        }
        $oPsi = New-Object -TypeName System.Diagnostics.ProcessStartInfo
        [string[]]$MUILanguages=(Get-WmiObject -query "select MUILanguages from win32_operatingsystem" ).MUILanguages
        if ($PSBoundParameters['Encoding'] -ne $null)
        {
            $ProcessEncoding=[System.Text.Encoding]::GetEncoding($Encoding)
        }
        elseif($MUILanguages -eq "ru-RU")
        {
            $ProcessEncoding=[System.Text.Encoding]::GetEncoding(1251)
        }
        $oPsi.StandardOutputEncoding=$ProcessEncoding
        $oPsi.StandardErrorEncoding=$ProcessEncoding
        $oPsi.CreateNoWindow = $true
        $oPsi.UseShellExecute = $false
        $oPsi.RedirectStandardOutput = $true
        $oPsi.RedirectStandardError = $true
        if ($PSBoundParameters["EnvVar"] -ne $null)
        {
            $EnvVar.Keys | foreach {
                [string]$Key=$_
                [string]$Value=$EnvVar[$Key]
                $oPsi.EnvironmentVariables.Add($Key,$Value)
            }
        }
        $oPsi.FileName = $ExeFile
    
        if (! [String]::IsNullOrEmpty($Args)) 
        {
            $oPsi.Arguments = $Args
        }
        if (! [String]::IsNullOrEmpty($Verb)) 
        {
            $oPsi.Verb = $Verb
        }
    
        $oProcess = New-Object -TypeName System.Diagnostics.Process
        $oProcess.StartInfo = $oPsi

        $oStdOutBuilder = New-Object -TypeName System.Text.StringBuilder
        $oStdErrBuilder = New-Object -TypeName System.Text.StringBuilder
        $StdOutObject=New-Object -TypeName psobject
        $StdErrObject=New-Object -TypeName psobject
        $StdOutObject | Add-Member -MemberType NoteProperty -Name StrBuilder -Value $oStdOutBuilder
        $StdOutObject | Add-Member -MemberType NoteProperty -Name LogPath -Value $LogPath
        $StdOutObject | Add-Member -MemberType NoteProperty -Name VerboseOutput -Value $($PSBoundParameters["VerboseOutput"].IsPresent)
        $StdErrObject | Add-Member -MemberType NoteProperty -Name StrBuilder -Value $oStdErrBuilder
        $StdErrObject | Add-Member -MemberType NoteProperty -Name LogPath -Value $LogPath
        $StdErrObject | Add-Member -MemberType NoteProperty -Name VerboseOutput -Value $($PSBoundParameters["VerboseOutput"].IsPresent)

        $sScripBlock = {
            if (!([String]::IsNullOrEmpty($EventArgs.Data))) 
            {
                
                if (!($Event.MessageData.VerboseOutput -eq $true) -and [string]::IsNullOrEmpty($event.MessageData.LogPath))
                {
                    $Event.MessageData.StrBuilder.AppendLine($EventArgs.Data)    
                }
                else
                {
                    if (!([string]::IsNullOrEmpty($event.MessageData.LogPath)))
                    {
                        $($EventArgs.Data) | Out-File -FilePath $($event.MessageData.LogPath) -Append -Force -WhatIf:$false -Confirm:$false  -ErrorAction Stop 
                    }
                    if ($Event.MessageData.VerboseOutput -eq $true)
                    {
                        Write-Verbose "$($EventArgs.Data)" -Verbose     
                    }    
                }
                
                
                
                    
  
            }
        }
        $oStdOutEvent = Register-ObjectEvent -InputObject $oProcess -Action $sScripBlock -EventName 'OutputDataReceived' -MessageData $StdOutObject
        $oStdErrEvent = Register-ObjectEvent -InputObject $oProcess -Action $sScripBlock -EventName 'ErrorDataReceived' -MessageData $StdErrObject
        Unregister-Event -SourceIdentifier ProcessExitedEvent -Confirm:$false -WhatIf:$false -ErrorAction SilentlyContinue
        Remove-Event -SourceIdentifier ProcessExitedEvent -ErrorAction SilentlyContinue -WhatIf:$false -Confirm:$false
        Register-ObjectEvent -InputObject $oProcess -EventName 'Exited' -SourceIdentifier ProcessExitedEvent
        
        [Void]$oProcess.Start()
         
        $oProcess.BeginOutputReadLine()
        $oProcess.BeginErrorReadLine()
        $ProcessClose=$false
        try
        {
                if ($PSBoundParameters["VerboseOutput"].isPresent -or !([string]::IsNullOrEmpty($PSBoundParameters["LogPath"])))
                {  
                    do 
                    {
                        Start-Sleep -Milliseconds 5
                    }while(!($oProcess.HasExited))    
                    $ProcessClose=$true
                }
                else
                {
         
                        Wait-Event -SourceIdentifier ProcessExitedEvent -ErrorAction Stop | Out-Null
                        $ProcessClose=$true
                
                }
        }
        finally
        {
                if (!($ProcessClose))
                {
                    Write-Verbose "Try stop process $($oProcess.ID) $($oProcess.name)"
                    Stop-Process -Id $($oProcess.ID) -WhatIf:$false -Confirm:$false -Force    
                }
                
                Unregister-Event -SourceIdentifier $oStdOutEvent.Name -Confirm:$false -WhatIf:$false
                Unregister-Event -SourceIdentifier $oStdErrEvent.Name -Confirm:$false -WhatIf:$false
                Unregister-Event -SourceIdentifier ProcessExitedEvent -Confirm:$false -WhatIf:$false
                Remove-Event -SourceIdentifier ProcessExitedEvent -Confirm:$false -WhatIf:$false -ErrorAction SilentlyContinue
        }

    
        
        
        $oResult = New-Object -TypeName PSObject -Property (@{
            "ExeFile"  = $ExeFile;
            "Args"     = $Args -join " ";
            "ExitCode" = $oProcess.ExitCode;
            "StdOut"   = $StdOutObject.StrBuilder.ToString().Trim();
            "StdErr"   = $StdErrObject.StrBuilder.ToString().Trim();
        })

        return $oResult
}
function CreateEventlogSource {
    [cmdletbinding()]
    param(
        [string]$EventlogSource
    )
    try{
        if ( -not [System.Diagnostics.EventLog]::SourceExists($EventlogSource) ){
            [System.Diagnostics.EventLog]::CreateEventSource($EventlogSource, "Application")
        } else{
            Write-Verbose "$EventlogSource : eventlog source already exists"  
        }
    } catch{
        Write-Error $_
    }
    
}
Function ParseParam{
    [cmdletbinding()]
    param(
    [parameter(Mandatory=$true)]
    [string]$ParamString,
    [parameter(Mandatory=$true)]
    [string]$BaseName
    )

    [string[]]$PermitParams="DayOfMonth","KeepVersions","KeepVersionsDiff","KeepVersionsTrn","KeepVersionsWeekly","KeepVersionsMonthly","KeepVersionsYearly","DayOfWeek","DayOfYear","CheckArchiveBit"
    [string[]]$SwitchParam="CheckArchiveBit"
    $ArrayHashTableParam=@()
    $ArrayParamString=(((($ParamString -replace "\s+"," ") -replace "\s+$","") -replace "^-"," -") -replace " -"," --") -split "\s-"
    $HashTableParam=@{}
    $ArrayParamString | foreach {
    
        if ($_ -match "^-(.+?)\s(.+)$"){
            $ParseParam=$Matches[1]
            $ParseValue=$Matches[2]
                if ($ParseValue -match ","){
                    if ($ParseParam -ne "Query"){
                        $ArrayParseValue=$ParseValue -split ","
                        $ParseValue=$ArrayParseValue
                    }
                
                }
            $HashTableParam.Add($ParseParam,$ParseValue)
        
        
    
        }
        elseif ($_ -match "-(.+\S)"){
            if ($SwitchParam -eq $Matches[1]){
                $HashTableParam.Add($Matches[1],$True)
            }
            else{
                $HashTableParam.Add($Matches[1],$null)  
            }     
        }
    # End Foreach
    }
    $ObjectParam=New-Object -TypeName psobject -Property $HashTableParam
    $DifObj=$ObjectParam | Get-Member -MemberType NoteProperty | foreach {$_.name}
    $CompareParam=Compare-Object -ReferenceObject $PermitParams -DifferenceObject $DifObj
    if ($CompareParam | where-object {$_.sideindicator -eq "=>"}){
        Write-Error "$BaseName :Incorrect Parameter $(($CompareParam | Where-Object {$_.SideIndicator -eq "=>"}).inputobject). Allowed param: $PermitParams Check configuration" -ErrorAction Stop
    }
    $ObjectParam | Add-Member -MemberType NoteProperty -Name BaseName -Value $BaseName -ErrorAction Stop
    $ObjectParam
    #End Function
}
function ChekClearDumpSettings{
    param(
        [parameter(Mandatory=$true)]
        [psobject]$Settings
    )
    $AllowedDayOfWeek="Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"
    $WeekParam="KeepVersionsWeekly","DayOfWeek"
    $MonthParam="KeepVersionsMonthly","DayOfMonth"
    $YearParam="KeepVersionsYearly","DayOfYear"
    [string[]]$OnlyNumbersParam="KeepVersions","KeepVersionsWeekly","KeepVersionsYearly","DayOfYear","KeepVersionsMonthly","DayOfMonth","KeepVersionsDiff","KeepVersionsTrn"
    if ($Settings.KeepVersions -eq $null){
        Write-Error "KeepVersions is null or empty" -ErrorAction Stop
    }
    if ($Settings.DayOfWeek -ne $null){
        [string[]]$Settings.DayOfWeek | foreach {
            $SettingsDayOfWeek=$_
            if (!($AllowedDayOfWeek -eq $SettingsDayOfWeek)){
                Write-Error "$($Settings.BaseName): Wrong parameter DayOfWeek: $SettingsDayOfWeek Allowed: {$AllowedDayOfWeek}" -ErrorAction Stop
            }
        }
        
    }
    $Settings.psobject.Properties  | foreach {
        if ($OnlyNumbersParam -eq $_.name){
            if (!($_.value -match "^\d+$")){
                Write-Error "$($Settings.BaseName): Wrong parameter $($_.name): {$($_.value)} Only numbers allowed" -ErrorAction Stop       
            }

        }
        
    }
    $Settings.psobject.Properties | foreach {
        if ($WeekParam -eq $_.name){
            if ($Settings.KeepVersionsWeekly -eq $null -or $Settings.DayOfWeek -eq $null){
                Write-Error "$($Settings.BaseName): Both parameters must be specified: {$WeekParam}" -ErrorAction Stop
            }
            [int]$KeepWeekly=$Settings.KeepVersionsWeekly
            if (!(($KeepWeekly -ge 1) -and ($KeepWeekly -le 2000))){
                Write-Error "$($Settings.BaseName): Wrong parameter KeepVersionsWeekly:{$($Settings.KeepVersionsWeekly)} Allowed range 1-2000" -ErrorAction Stop
            }
        }
        if ($MonthParam -eq $_.name){
            if ($Settings.KeepVersionsMonthly -eq $null -or $Settings.DayOfMonth -eq $null){
                Write-Error "$($Settings.BaseName): Both parameters must be specified: {$MonthParam}" -ErrorAction Stop
            }
            [int]$KeepMonthly=$Settings.KeepVersionsMonthly
            if (!($KeepMonthly -ge 1 -and $KeepMonthly -le 2000)){
                Write-Error "$($Settings.BaseName): Wrong parameter KeepVersionsMonthly:{$($Settings.KeepVersionsMonthly)} Allowed range 1-2000" -ErrorAction Stop
            }
            [int[]]$DayOfMonth=$Settings.DayOfMonth
            $DayOfMonth | foreach {
                if (!($_ -ge 1 -and $_ -le 31)){
                    Write-Error "$($Settings.BaseName): Wrong parameter DayOfMonth:{$($Settings.DayOfMonth)} Allowed range 1-31" -ErrorAction Stop
                }   
            }
            
        }
        if ($YearParam -eq $_.name){
            if ($Settings.KeepVersionsYearly -eq $null -or $Settings.DayOfYear -eq $null){
                Write-Error "$($Settings.BaseName): Both parameters must be specified: {$YearParam}" -ErrorAction Stop
            }
            [int]$KeepYearly=$Settings.KeepVersionsYearly 
            if (!($KeepYearly -ge 1 -and $KeepYearly -le 2000)){
                Write-Error "$($Settings.BaseName): Wrong parameter KeepVersionsYearly:{$($Settings.KeepVersionsYearly)} Allowed range 1-2000" -ErrorAction Stop
            }
            [int[]]$DayOfYear=$Settings.DayOfYear
            $DayOfYear | foreach {
                if (!($_ -ge 1 -and $_ -le 366)){
                    Write-Error "$($Settings.BaseName): Wrong parameter DayOfYear:{$($Settings.DayOfYear)} Allowed range 1-366" -ErrorAction Stop
                }
            }
            
        }
    }


}
function ListFtpDirectory{
    [cmdletbinding()]
    param(
        [parameter(Mandatory=$true)]
        [string]$Url, 
        [int]$Deph,
        [parameter(Mandatory=$true)]        
        $Credential
    )
    if (!($Url -match "^ftp://")){
        Write-Error "Incorrect url: $Url" -ErrorAction Stop
    }
    if ($Credential -eq $null){
        Write-Error "Credential is null" -ErrorAction Stop
    }
    
    $listRequest = [Net.WebRequest]::Create($url)
    $listRequest.Method = [System.Net.WebRequestMethods+Ftp]::ListDirectoryDetails
    $listRequest.Credentials = $Credential
    
    $lines = New-Object System.Collections.ArrayList

    try{
        $listResponse = $listRequest.GetResponse()
        $listStream = $listResponse.GetResponseStream()
        $listReader = New-Object System.IO.StreamReader($listStream)
        while (!$listReader.EndOfStream){
            $line = $listReader.ReadLine()
            $lines.Add($line) | Out-Null
           
        }
        $listReader.Dispose()
        $listStream.Dispose()
        if (($listResponse | Get-Member | Where-Object {$_.name -eq "Dispose"})){
            $listResponse.Dispose()   
        }
        
    }
    catch{

        Write-Error $_ -ErrorAction Stop
    }
    
    foreach ($line in $lines){
        $Tokens = $line.Split(" ", 9, [StringSplitOptions]::RemoveEmptyEntries)
        $Name = $tokens[8]
        $Permissions = $tokens[0]
        
        if ($permissions[0] -eq 'd'){
            if (!($url -match "/$")){
                
                $Url+="/"
            }
            $DirectoryUrl=$Url+$name +"/"
            $Item=New-Object -TypeName psobject -Property @{
                Name=$Name
                FullName=$DirectoryUrl
                PSIsContainer=$true
            }
            Write-Verbose "Directory $Name"
            if ($PSBoundParameters['Deph'] -ge 1){
                $Deph-=1
                ListFtpDirectory -Url $DirectoryUrl -Credential $Credential -Deph $Deph
                $Deph+=1
            }
            
            
            
        }
        else{
            if ($lines.Count -ge 1){
                $FileUrl = ($url + $name)    
            }
            else{
                $FileUrl=$Url
            }
            
            $Item=New-Object -TypeName psobject -Property @{
                Name=$Name
                FullName=$FileUrl
                PSIsContainer=$false
            }
            Write-Verbose "File $name"
        }
        $Item
        
    }
}
function TestFtpPath{
    [cmdletbinding()]
    param(
        [parameter(Mandatory=$true)]
        [string]$Path,
        [parameter(Mandatory=$true)]        
        $Credential
    )
    if (!($Path -match "^ftp://")){
        Write-Error "Incorrect url: $Path" -ErrorAction Stop
    }
    if ($Credential -eq $null){
        Write-Error "Credential is null" -ErrorAction Stop
    }
    
    try{
        $Request = [Net.WebRequest]::Create($Path)
        $Request.Method = [System.Net.WebRequestMethods+Ftp]::GetFileSize
        $Request.Credentials = $Credential
        $Response=$Request.GetResponse()
        if (($Response | Get-Member | Where-Object {$_.name -eq "Dispose"})){
            $Response.Dispose()   
        }
        $true    
    }
    catch{
        $response = $_.Exception.InnerException.Response;
        if ($response.StatusCode -eq [Net.FtpStatusCode]::ActionNotTakenFileUnavailable){
            $false
        }
        else{
            Write-Error $_ -ErrorAction Stop
        }
    }


}
function RemoveDump{
    [cmdletbinding(SupportsShouldProcess=$true)]
    param(
        [parameter(Mandatory=$true)]
        [string]$Path,
        [parameter(Mandatory=$true)]
        [string]$StorageType,
        $Credential,
        $RcloneCredential
    )
    try{
        if ($StorageType -eq "Disk"){
            if (Test-Path $Path){
                #Write-Debug -Message debug -Debug
                Remove-Item -Path $Path -WhatIf:$([bool]$WhatIfPreference.IsPresent) -ErrorAction Stop
            }
            else{
                Write-Verbose "$Path already removed"
            }
        } 
        elseif($StorageType -eq "Cloud"){
            Remove-ItemRclone -Path $Path -RcloneCredential $RcloneCredential -WhatIf:$([bool]$WhatIfPreference.IsPresent) -ErrorAction Stop
        }
        else
        {
            #Write-Debug -Message dbg -Debug
            Write-Verbose "RemoveFtpItem -Path $Path"
            RemoveFtpItem -Path $Path -Credential $Credential  -WhatIf:$([bool]$WhatIfPreference.IsPresent)  -ErrorAction Stop
        }
    } catch{
        $PSCmdlet.WriteError($_)
    }
    

}
function RemoveFtpItem{
    [cmdletbinding(SupportsShouldProcess=$true)]
    param(
        [parameter(Mandatory=$true)]
        [string]$Path,
        $Credential
    )
    try{
        if (!($Path -match "^ftp://")){
            Write-Error "Incorrect url: $Path" -ErrorAction Stop
        }
        if ($Credential -eq $null){
            Write-Error "Credential is null" -ErrorAction Stop
        }
        if (TestFtpPath -Path $Path -Credential $Credential){
            $FtpItem=ListFtpDirectory -Url $Path -Credential $credential
        }
         
        if ($FtpItem){
            if ($FtpItem.PSIsContainer){
                Write-Error "This path is Directory. Incorrect path: $Path" -ErrorAction Stop
            }
            if ($([bool]$WhatIfPreference.IsPresent) -eq $false){
                $Request = [Net.WebRequest]::Create($Path)
                $Request.Method = [System.Net.WebRequestMethods+Ftp]::DeleteFile
                $Request.Credentials = $Credential
                $Response=$Request.GetResponse()
                if (($Response | Get-Member | Where-Object {$_.name -eq "Dispose"})){
                    $Response.Dispose()   
                }
                 
            }
            else{
                Write-Verbose "Try RemoveFtpItem -Path $Path" -Verbose    
            }
        }
        else{
            Write-Verbose "Already deleted: $Path"
        }
            
    }
    catch{
        Write-Error $_
    }
    
}
function GetDirRecurse{
    [cmdletbinding()]
    param(
        [parameter(Mandatory=$true)]
        [string]$Path,
        [parameter(Mandatory=$true)]
        [ValidateRange(0,5)]
        [int]$Deph,
        $Credential,
        $RcloneCredential
    )
    
    if ($PSBoundParameters['Deph'] -gt 0){
        if ($Path -match "^.+:\\"){
           $Dirs=Get-ChildItem -Path $Path
        }
        elseif($Path -match "^ftp://"){
            $Dirs=ListFtpDirectory -Url $Path -Credential $Credential
        }
        else{
           $Dirs=Get-ChildItemRclone -Path $Path -RcloneCredential $RcloneCredential
        }
         
        $NestedDirs=$Dirs | Where-Object {$_.PSIsContainer}
        if ($NestedDirs){
            foreach ($NestedDir in $NestedDirs){
                if ($NestedDir.fullname -ne $null){
               
                    if ($PSBoundParameters['Deph'] -gt 1){
                        $Deph-=1
                        GetDirRecurse -Path $NestedDir.fullname -Deph $Deph -Credential $Credential -RcloneCredential $RcloneCredential
                        $Deph+=1
                    }
                }
            
            }
         
        }
        $NestedDirs
    }

}
function GetDumps{
    [cmdletbinding()]
    param(
        [parameter(Mandatory=$true)]
        [string]$BackupPath,
        [ValidateRange(0,100)]
        [int]$Deph=0,
        $Credential,
        $RcloneCredential
    )
    [string[]]$OtherMssqlExtensions="diff","trn"
    $AllPaths=@()
    if ($BackupPath -match "^.+:\\"){
        $AllPaths+=Get-Item -Path $BackupPath -ErrorAction Stop
        $Storage="disk"
    }
    elseif($BackupPath -match "^ftp://"){
        $AllPaths+=New-Object -TypeName psobject -Property @{
            FullName=$BackupPath   
        }
        $Storage="ftp"
    }
    else{
        $AllPaths+=New-Object -TypeName psobject -Property @{
            FullName=$BackupPath   
        }
        $Storage="cloud"
    }
    <#if ($BackupPath -match "^ftp://"){
        $AllPaths+=New-Object -TypeName psobject -Property @{
            FullName=$BackupPath   
        }
        $Storage="ftp"
    }
    else{
        $AllPaths+=Get-Item -Path $BackupPath -ErrorAction Stop
        $Storage="disk"
    }#>
    
    if($Deph -ge 1){
        $AllPaths+=GetDirRecurse -Path $BackupPath -Deph $Deph  -Credential $Credential -RcloneCredential $RcloneCredential 
    }
    
    $AllPaths | foreach {
        $Path=$_.fullname
        if ($BackupPath -match "^.+:\\"){
            Write-Verbose "Get-ChildItem -Path $Path"
            $DumpItems=Get-ChildItem -Path $Path 
        }
        elseif($BackupPath -match "^ftp://"){
            Write-Verbose "ListFtpDirectory -Url $Path"
            $DumpItems=ListFtpDirectory -Url $Path -Credential $Credential
        }
        else{
            Write-Verbose "Get-ChildItemRclone -Path $Path"
            $DumpItems=Get-ChildItemRclone -Path $Path -RcloneCredential $RcloneCredential
        }
        <#if ($Path -match "^ftp://"){
            Write-Verbose "ListFtpDirectory -Url $Path"
            $DumpItems=ListFtpDirectory -Url $Path -Credential $Credential
        }
        else{
            Write-Verbose "Get-ChildItem -Path $Path"
            $DumpItems=Get-ChildItem -Path $Path    
        }#>
        
        $AllBackups=$DumpItems | Where-Object {$_.PSIsContainer -eq $false}
        $OutObjects=@()
        $AllBackups | foreach {
            $FileName=$_.Name
            $FullName=$_.FullName
            $FileAttributes=$_.Attributes
            #$LastWriteTime=$_.LastWriteTime
            #$Length=$_.Length
            if ($FileName -match "^(.+)_(.+)_([\d]{4})_([\d]{2})_([\d]{2})_([\d]{2})([\d]{2})([\d]{2}).*\.(.+)$"){
                
                #"(.+)_([\d]{8})_([\d]{4})\..+$"
                $BaseName=$Matches[1]
                $Year=$Matches[3]
                $Month=$Matches[4]
                $Day=$Matches[5]
                $Hour=$Matches[6]
                $Minute=$Matches[7]
                $Second=$Matches[8]
                $Extension=$Matches[9]
                #$Date=$Matches[2]
                #$Time=$Matches[3]
                <#if ($Date -match "^(\d\d)(\d\d)(\d\d\d\d)$"){
                    $Day=$Matches[1]
                    $Month=$Matches[2]
                    $Year=$Matches[3]
                }
                else{
                    Write-Error "Incorrect Date string $Date" -ErrorAction Stop
                }
                if ($Time -match "^(\d\d)(\d\d)$"){
                    $Hour=$Matches[1]
                    $Minute=$Matches[2]
                }
                else{
                    Write-Error "Incorrect time string $Time" -ErrorAction Stop
                }#>
                $BackupCreateDate=Get-Date -Day $Day -Month $Month -Year $Year -Hour $Hour -Minute $Minute -Second $Second
                #$DumpItem=Get-Item -Path $FullName
                $IsLogOrDiff=$false
                if ($OtherMssqlExtensions -eq $Extension){
                    $IsLogOrDiff=$true
                }
                $ArchiveBitIsPresent=$false
                if ($Storage -eq "Disk"){
                    if ((($FileAttributes -band [io.fileattributes]::Archive).value__ -eq 32) -or (($FileAttributes -band [io.fileattributes]::Archive) -eq 32)){
                        $ArchiveBitIsPresent=$true
                    }
              
                }
                $OutObject=New-Object -TypeName psobject -Property @{
                    "BaseName"=$($BaseName.ToLower());
                    "FullName"=$FullName;
                    "FileName"=$FileName;
                    "CreateDate"=$BackupCreateDate;
                    "Extension"=$Extension;
                    "Storage"=$Storage;
                    "IsLogOrDiff"=$IsLogOrDiff;
                    "ArchiveBitIsPresent"=$ArchiveBitIsPresent
                    #"Length"=$DumpItem.Length

                }
                $OutObjects+=$OutObject

            }
            else{
                Write-Verbose "Skip $FullName incorrect name format. Correct format ^basename_.+_yyyy_MM_dd_HHmmss.*\.(.+)$" -Verbose
            } 
        }
        if ($OutObjects.Count -eq 0){
            Write-Verbose "Backup not found in folder $Path"
        }
        else{
            $OutObjects | Sort-Object -Property CreateDate
        }    
    }
    
    
}
function GetOldDumps{
    [cmdletbinding()]
    param(
        [parameter(Mandatory=$true)]
        [PsObject[]]$Dumps,
        [parameter(Mandatory=$true)]
        [psobject]$KeepSettings
    )
    
    $KeepParam=@{
        "IsWeekly"="KeepVersionsWeekly"
        "IsMonthly"="KeepVersionsMonthly"
        "IsYearly"="KeepVersionsYearly"
        "IsDiff"="KeepVersionsDiff"
        "IsTrn"="KeepVersionsTrn"
    }
    [string[]]$DumpTags="IsDaily","IsDiff","IsTrn","IsWeekly","IsMonthly","IsYearly"
    $TaggedDumps =TagsDump -Dumps $Dumps -Settings $KeepSettings
    
    $AllRemovedBackupBases=@()
    
    $DumpTags | foreach {
            $DumpTag=$_
            [array]$MustBeRemovedBases=$TaggedDumps  | Where-Object {$_.IsDiff -ne $true -and $_.IsTrn -ne $true} | Sort-Object -Property CreateDate -Descending | Select-Object -Skip $([int]$KeepSettings.KeepVersions)
            if ($DumpTag -eq "IsDaily"){
                 [array]$MustBeRemovedBases=$MustBeRemovedBases | Where-Object {$($_.$DumpTag) -eq $true}
            }
            elseif($DumpTag -eq "IsDiff"){
                [int]$Skip=$KeepSettings.$($KeepParam["$DumpTag"])
                [array]$MustBeRemovedBases=$TaggedDumps | Where-Object {$($_.$DumpTag) -eq $true} | Sort-Object -Property CreateDate -Descending  | Select-Object -Skip $Skip
            }
            elseif($DumpTag -eq "IsTrn"){
                [int]$Skip=$KeepSettings.$($KeepParam["$DumpTag"])
                [array]$MustBeRemovedBases=$TaggedDumps | Where-Object {$($_.$DumpTag) -eq $true} | Sort-Object -Property CreateDate -Descending  | Select-Object -Skip $Skip
            }
            else{
                [int]$Skip=$KeepSettings.$($KeepParam["$DumpTag"])
                [array]$MustBeRemovedBases=$MustBeRemovedBases | Where-Object {$($_.$DumpTag) -eq $true} | Select-Object -Skip $Skip
                
            }
            #Write-Verbose "$($BaseName+':') Found $($MustBeRemovedBases.count) old copy"
            $AllRemovedBackupBases+=$MustBeRemovedBases 
    }
    $AllRemovedBackupBases
}
function TagsDump{
    [cmdletbinding()]
    param(
        [parameter(Mandatory=$true)]
        [psobject[]]$Dumps,
        [parameter(Mandatory=$true)]
        [psobject]$Settings
    )
    
    [int]$KeepDaysSettings=$Settings.KeepVersions
    [int]$KeepVersionsDiff=$Settings.KeepVersionsDiff
    [int]$KeepVersionsTrn=$Settings.KeepVersionsTrn
    [int]$KeepWeeklySettings=$Settings.KeepVersionsWeekly
    [string[]]$DayOfWeekSettings=$Settings.DayOfWeek
    [int[]]$DayOfMonthSettings=$Settings.DayOfMonth
    [int]$KeepYarSettingsr=$Settings.KeepVersionsYarly
    [int[]]$DayOfYearSettings=$Settings.DayOfYear
    if ($Settings.CheckArchiveBit){
        [bool]$CheckArchiveBit=$Settings.CheckArchiveBit    
    } else{
        [bool]$CheckArchiveBit=$false
    }

    
    $TaggedDumps=@()
    [int]$DumpIndex=0
    $Dumps | foreach{
        $NextDumpIndex=$ItemIndex+=1
        $Dump=$_
        $NextDump=$Dumps[$NextDumpIndex]
        $Dump | Add-Member -MemberType NoteProperty -Name IsDaily -Value $false -ErrorAction Stop -Force
        $Dump | Add-Member -MemberType NoteProperty -Name IsWeekly -Value $false -ErrorAction Stop -Force
        $Dump | Add-Member -MemberType NoteProperty -Name IsMonthly -Value $false -ErrorAction Stop -Force
        $Dump | Add-Member -MemberType NoteProperty -Name IsYearly -Value $false -ErrorAction Stop -Force
        $Dump | Add-Member -MemberType NoteProperty -Name IsDiff -Value $false -ErrorAction Stop -Force
        $Dump | Add-Member -MemberType NoteProperty -Name IsTrn -Value $false -ErrorAction Stop -Force
        $Dump | Add-Member -MemberType NoteProperty -Name BlockDelete -Value $false -ErrorAction Stop -Force
        $DumpDate=$Dump.CreateDate
        [string]$DumpDayOfWeek=$DumpDate.DayOfWeek
        [int]$DumpDayOfMonth=$DumpDate.Day
        [int]$DumpDayOfYear=$DumpDate.DayOfYear
        
        if(($KeepVersionsDiff -ne 0 -and $Dump.Extension -eq "diff") -or ($KeepVersionsTrn -ne 0 -and $Dump.Extension -eq "trn")){
            if ($Dump.Extension -eq "diff"){
                $Dump.IsDiff=$true 
                
            }
            elseif($Dump.Extension -eq "trn"){
                $Dump.IsTrn=$true
                
            }   
            
        }
        elseif ($DayOfYearSettings -eq $DumpDayOfYear -and $DayOfYearSettings -ne 0){
            $Dump.IsYearly=$true

        }
        elseif($DayOfMonthSettings -eq $DumpDayOfMonth -and $DayOfMonthSettings -ne 0){
            $Dump.IsMonthly=$true
        }
        elseif($DayOfWeekSettings -eq $DumpDayOfWeek -and $DayOfWeekSettings -ne 0){
            $Dump.IsWeekly=$true
        }
        else{
            $Dump.IsDaily=$true
        }
            
        if ($NextDump.IsLogOrDiff){
            if (!$($Dump.IsLogOrDiff)){
                $Dump.BlockDelete=$true
            }
            elseif($Dump.extension -eq "diff"){
                if ($NextDump.extension -eq "trn"){
                    $Dump.BlockDelete=$true
                }
            }
            
        }
        if ($CheckArchiveBit -eq $true -and $Dump.Storage -eq "disk"){
            if($Dump.ArchiveBitIsPresent){
                $Dump.BlockDelete=$true
            }
        }
        $TaggedDumps+=$Dump
        
        $DumpIndex+=1
    }
    
    $TaggedDumps
}
function RemoveOldDumps{
    [cmdletbinding(SupportsShouldProcess=$true)]
    param(
        [parameter(Mandatory=$true)]
        [string[]]$DumpPaths,
        $AllKeepSettings,
        $DefaultKeepSettings,
        [int]$RecurseDeph,
        $Credential
    )
    try{
        $StartRcloneServer=$false
        if($DumpPaths -match ".+:[^\\//]" -or $DumpPaths -match ".+:$"){
            $RcloneUser="rclone"
            $RclonePassword=NewPassword
            $RcloneCredential=CreateCredential -User $RcloneUser -Password $RclonePassword
            StartRcloneApiServer -ErrorAction Stop -RcloneCredential $RcloneCredential
            Write-Verbose "Set StartRcloneServer true"
            $StartRcloneServer=$true
        }
        $AllDumps = New-Object System.Collections.ArrayList
        foreach ($DumpPath in $DumpPaths){
            Write-Verbose "GetDumps -BackupPath $DumpPath -Deph $RecurseDeph"
            GetDumps -BackupPath $DumpPath -Deph $RecurseDeph -Credential $Credential -RcloneCredential $RcloneCredential | foreach {
                $AllDumps.Add($_) | Out-Null
            }
        }
        if ([int]$AllDumps.count -eq 0){
            Write-Error "Backup not found" -ErrorAction Stop
        }

        $MustBeRemoved=@()
        $AllKeepSettings | foreach {
            $BaseName=$_.basename
            $KeepSettings=$_
            $AllDumpsBase=@()
    

                $AllDumps.Clone() | Where-Object {$_.BaseName -eq $BaseName} | foreach {
                    $AllDumpsBase+=$_
                    $AllDumps.Remove($_)
                } 
    
                if ($AllDumpsBase.Count -ne 0){
                    GetOldDumps -Dumps $AllDumpsBase -KeepSettings $KeepSettings | foreach {
                        $MustBeRemoved+=$_
            
                    }
                }
                else{
                    Write-Verbose "Skip Base $BaseName"
                }
        
        }

        if ($AllDumps.count -ne 0 -and !([string]::IsNullOrEmpty($DefaultKeepSettings))){
            $BaseNames=$AllDumps | Select-Object -Property BaseName -Unique | foreach {$_.basename}
            $BaseNames | foreach {
                $BaseName=$_
                $AllDumpsBase=$AllDumps | Where-Object {$_.BaseName -eq $BaseName}
                GetOldDumps -Dumps $AllDumpsBase -KeepSettings $DefaultKeepSettings | foreach {
                    $MustBeRemoved+=$_
                }
            }
    
      
        }
        $Results=@()

        $MustBeRemoved | Sort-Object -property BaseName -Unique |  foreach {
            $BaseName=$_.basename
            $Result=New-Object -TypeName psobject 
            $Result | Add-Member -MemberType NoteProperty -Name BaseName -Value $BaseName
            $Result | Add-Member -MemberType NoteProperty -Name TotalDeleted -Value $([int]0)
            $Result | Add-Member -MemberType NoteProperty -Name OldVersions -Value $([int]0)
            $Result | Add-Member -MemberType NoteProperty -Name OldDiffVersions -Value $([int]0)
            $Result | Add-Member -MemberType NoteProperty -Name OldTrnVersions -Value $([int]0)
            $Result | Add-Member -MemberType NoteProperty -Name OldWeekly -Value $([int]0)
            $Result | Add-Member -MemberType NoteProperty -Name OldMonthly -Value $([int]0)
            $Result | Add-Member -MemberType NoteProperty -Name OldYearly -Value $([int]0)

            $Results+=$Result
        }
        if ($MustBeRemoved.Count -gt 0){
            $MustBeRemoved | foreach {
                $FullName=$_.fullName
                $BaseName=$_.BaseName
                $StorageType=$_.Storage
                $IsDaily=$_.IsDaily
                $IsWeekly=$_.IsWeekly
                $IsMonthly=$_.IsMonthly
                $IsYearly=$_.IsYearly
                $IsDiff=$_.IsDiff
                $IsTrn=$_.IsTrn
                $BlockDelete=$_.BlockDelete
            
                if (!([string]::IsNullOrEmpty($FullName))){
                    if ($BlockDelete -eq $false){
                        RemoveDump -Path $FullName -StorageType $StorageType -Credential $Credential -RcloneCredential $RcloneCredential -WhatIf:$([bool]$WhatIfPreference.IsPresent)
                        if ($?){
                            $Results | Where-Object {$_.Basename -eq $BaseName} | foreach {
                                if ($IsDaily){
                                   $_.OldVersions+=1
                                }
                                elseif($IsDiff){
                                    $_.OldDiffVersions+=1
                                }
                                elseif($IsTrn){
                                    $_.OldTrnVersions+=1
                                }
                                elseif($IsWeekly) {
                                    $_.OldWeekly+=1
                                }
                                elseif($IsMonthly){
                                    $_.OldMonthly+=1
                                }
                                elseif($IsYearly){
                                    $_.OldYearly+=1
                                }   
                                $_.TotalDeleted+=1
                            }    
                        }    
                
                    }
                    else{
                        Write-Verbose "File is locked. There may be log files or differential copies associated with this file, or ArchiveBit is installed. Skip delete $FullName" -Verbose
                    }


                }
        

            }     
            $Results
        }
        else{
            Write-Verbose "Old backup versions not found" -Verbose
        }
        if($StartRcloneServer){
            $RcloneServer.PowerShell.Dispose()
            Remove-Variable -Name RcloneServer -Force -Scope Global -WhatIf:$false
        }
    }catch{
        Write-Error $_
        if($StartRcloneServer){
                $RcloneServer.PowerShell.Dispose()
                Remove-Variable -Name RcloneServer -Force -Scope Global -WhatIf:$false
        }   
    }
    

}
function RenameBackup{
   [CmdletBinding(SupportsShouldProcess=$true)]
    param(
        [parameter(Mandatory=$true)]
        [string]$Path
    )
    
    Get-ChildItem -Path $Path -Recurse | Where-Object {$_.PSIsContainer -eq $false} | foreach {
        $FullPath=$_.FullName
        $FileName=Split-Path -Path $FullPath -Leaf
        if ($FullPath -match ".+\.log$"){
            Remove-Item -Path $FullPath
        }
        if ($FileName -match "(.+)[\s]{1,3}(\d\d\d\d)-(\d\d)-(\d\d)[\s]{1,3}([\d]{1,2})-(\d\d)-(\d\d)\.(.+)$"){
            $BaseName=$Matches[1]
            $Year=$Matches[2]
            $Month=$Matches[3]
            $Day=$Matches[4]
            $Hour=$Matches[5]
            $Minute=$Matches[6]
            $Extension=$Matches[8]
            $Second="00"
            if ($Hour -match "^\d$"){
                $Hour="0"+$Hour
            }
            $NewName="$BaseName"+"_backup_"+"$Year"+"_"+"$Month"+"_"+$Day+"_"+"$Hour"+"$Minute"+$Second+"."+"$Extension"
            if (Test-Path $FullPath){
                Rename-Item -Path $FullPath -NewName $NewName   
            }
            
        }
        elseif($FileName -match "(.+)_([\d]{2})([\d]{2})([\d]{4})_([\d]{2})([\d]{2})(\..+)$"){
            $BaseName=$Matches[1]
            $Day=$Matches[2]
            $Month=$Matches[3]
            $Year=$Matches[4]
            $Hour=$Matches[5]
            $Minute=$Matches[6]
            $Extension=$Matches[7]
            $Second="00"
            $NewName="$BaseName"+"_backup_"+"$Year"+"_"+"$Month"+"_"+$Day+"_"+"$Hour"+"$Minute"+$Second+"$Extension"
            if (Test-Path $FullPath){
                Rename-Item -Path $FullPath -NewName $NewName   
            }
        }
        else{
            Write-Verbose "skip $FullPath" -Verbose
        }
        
        
    }

}
function CreateSettingsObject{
    [cmdletbinding()]
    param(
        [string]$DefaultSettings,
        $Settings
    )
    
    $AllKeepSettings=@()
    if (!([string]::IsNullOrEmpty($DefaultSettings))){
        [psobject]$DefaultKeepSettings=ParseParam -ParamString $DefaultSettings -BaseName "DefaultSettings" -ErrorAction Stop
        $DefaultKeepSettings | Add-Member -MemberType NoteProperty -Name IsDefaultSettings -Value $True -Force
        $AllKeepSettings+=$DefaultKeepSettings    
    }
    
    $Bases=$null
    [string[]]$Bases=$Settings.Keys
    if ($Bases.Count -gt 0){
        foreach ($Base in $Bases){
            if ($Base -match "\s"){
                Write-Error "$Base Incorrect Settings. Database name contains space" -ErrorAction Stop
            }
            $SettingsBase=ParseParam -ParamString $Settings["$Base"] -BaseName $Base -ErrorAction Stop
            $SettingsBase | Add-Member -MemberType NoteProperty -Name IsDefaultSettings -Value $False -Force
            $AllKeepSettings+=$SettingsBase
        }
    }
    else{
        Write-Verbose "Skip Database settings in Settings"
    }
    $AllKeepSettings
}
function ConvertToJson20 {
# Author: Joakim Borger Svendsen, 2017. 
# JSON info: http://www.json.org
# Svendsen Tech. MIT License. Copyright Joakim Borger Svendsen / Svendsen Tech. 2016-present.
# https://github.com/EliteLoser/ConvertTo-Json/blob/master/ConvertTo-STJson.ps1
    [CmdletBinding()]
    #[OutputType([Void], [Bool], [String])]
    Param(
        [AllowNull()]
        [Parameter(Mandatory=$True,
                   ValueFromPipeline=$True,
                   ValueFromPipelineByPropertyName=$True)]
        $InputObject,
        [Switch] $Compress,
        [Switch] $CoerceNumberStrings = $False,
        [Switch] $DateTimeAsISO8601 = $False)
    Begin{
        function EscapeJson {
            param(
                [String] $String)
            # removed: #-replace '/', '\/' `
            # This is returned 
            $String -replace '\\', '\\' -replace '\n', '\n' `
                -replace '\u0008', '\b' -replace '\u000C', '\f' -replace '\r', '\r' `
                -replace '\t', '\t' -replace '"', '\"'
        }
        function GetNumberOrString {
            param(
                $InputObject)
            if ($InputObject -is [System.Byte] -or $InputObject -is [System.Int32] -or `
                ($env:PROCESSOR_ARCHITECTURE -imatch '^(?:amd64|ia64)$' -and $InputObject -is [System.Int64]) -or `
                $InputObject -is [System.Decimal] -or `
                ($InputObject -is [System.Double] -and -not [System.Double]::IsNaN($InputObject) -and -not [System.Double]::IsInfinity($InputObject)) -or `
                $InputObject -is [System.Single] -or $InputObject -is [long] -or `
                ($Script:CoerceNumberStrings -and $InputObject -match $Script:NumberRegex)) {
                Write-Verbose -Message "Got a number as end value."
                "$InputObject"
            }
            else {
                Write-Verbose -Message "Got a string (or 'NaN') as end value."
                """$(EscapeJson -String $InputObject)"""
            }
        }
        function ConvertToJsonInternal {
            param(
                $InputObject, # no type for a reason
                [Int32] $WhiteSpacePad = 0)
    
            [String] $Json = ""
    
            $Keys = @()
    
            Write-Verbose -Message "WhiteSpacePad: $WhiteSpacePad."
    
            if ($null -eq $InputObject) {
                Write-Verbose -Message "Got 'null' in `$InputObject in inner function"
                $null
            }
    
            elseif ($InputObject -is [Bool] -and $InputObject -eq $true) {
                Write-Verbose -Message "Got 'true' in `$InputObject in inner function"
                $true
            }
    
            elseif ($InputObject -is [Bool] -and $InputObject -eq $false) {
                Write-Verbose -Message "Got 'false' in `$InputObject in inner function"
                $false
            }
    
            elseif ($InputObject -is [DateTime] -and $Script:DateTimeAsISO8601) {
                Write-Verbose -Message "Got a DateTime and will format it as ISO 8601."
                """$($InputObject.ToString('yyyy\-MM\-ddTHH\:mm\:ss'))"""
            }
    
            elseif ($InputObject -is [HashTable]) {
                $Keys = @($InputObject.Keys)
                Write-Verbose -Message "Input object is a hash table (keys: $($Keys -join ', '))."
            }
    
            elseif ($InputObject.GetType().FullName -eq "System.Management.Automation.PSCustomObject") {
                $Keys = @(Get-Member -InputObject $InputObject -MemberType NoteProperty |
                    Select-Object -ExpandProperty Name)

                Write-Verbose -Message "Input object is a custom PowerShell object (properties: $($Keys -join ', '))."
            }
    
            elseif ($InputObject.GetType().Name -match '\[\]|Array') {
        
                Write-Verbose -Message "Input object appears to be of a collection/array type. Building JSON for array input object."
        
                $Json += "[`n" + (($InputObject | ForEach-Object {
            
                    if ($null -eq $_) {
                        Write-Verbose -Message "Got null inside array."

                        " " * ((4 * ($WhiteSpacePad / 4)) + 4) + "null"
                    }
            
                    elseif ($_ -is [Bool] -and $_ -eq $true) {
                        Write-Verbose -Message "Got 'true' inside array."

                        " " * ((4 * ($WhiteSpacePad / 4)) + 4) + "true"
                    }
            
                    elseif ($_ -is [Bool] -and $_ -eq $false) {
                        Write-Verbose -Message "Got 'false' inside array."

                        " " * ((4 * ($WhiteSpacePad / 4)) + 4) + "false"
                    }
            
                    elseif ($_ -is [DateTime] -and $Script:DateTimeAsISO8601) {
                        Write-Verbose -Message "Got a DateTime and will format it as ISO 8601."

                        " " * ((4 * ($WhiteSpacePad / 4)) + 4) + """$($_.ToString('yyyy\-MM\-ddTHH\:mm\:ss'))"""
                    }
            
                    elseif ($_ -is [HashTable] -or $_.GetType().FullName -eq "System.Management.Automation.PSCustomObject" -or $_.GetType().Name -match '\[\]|Array') {
                        Write-Verbose -Message "Found array, hash table or custom PowerShell object inside array."

                        " " * ((4 * ($WhiteSpacePad / 4)) + 4) + (ConvertToJsonInternal -InputObject $_ -WhiteSpacePad ($WhiteSpacePad + 4)) -replace '\s*,\s*$'
                    }
            
                    else {
                        Write-Verbose -Message "Got a number or string inside array."

                        $TempJsonString = GetNumberOrString -InputObject $_
                        " " * ((4 * ($WhiteSpacePad / 4)) + 4) + $TempJsonString
                    }

                }) -join ",`n") + "`n$(" " * (4 * ($WhiteSpacePad / 4)))],`n"

            }
            else {
                Write-Verbose -Message "Input object is a single element (treated as string/number)."

                GetNumberOrString -InputObject $InputObject
            }
            if ($Keys.Count) {

                Write-Verbose -Message "Building JSON for hash table or custom PowerShell object."

                $Json += "{`n"

                foreach ($Key in $Keys) {

                    # -is [PSCustomObject]) { # this was buggy with calculated properties, the value was thought to be PSCustomObject

                    if ($null -eq $InputObject.$Key) {
                        Write-Verbose -Message "Got null as `$InputObject.`$Key in inner hash or PS object."
                        $Json += " " * ((4 * ($WhiteSpacePad / 4)) + 4) + """$Key"": null,`n"
                    }

                    elseif ($InputObject.$Key -is [Bool] -and $InputObject.$Key -eq $true) {
                        Write-Verbose -Message "Got 'true' in `$InputObject.`$Key in inner hash or PS object."
                        $Json += " " * ((4 * ($WhiteSpacePad / 4)) + 4) + """$Key"": true,`n"            }

                    elseif ($InputObject.$Key -is [Bool] -and $InputObject.$Key -eq $false) {
                        Write-Verbose -Message "Got 'false' in `$InputObject.`$Key in inner hash or PS object."
                        $Json += " " * ((4 * ($WhiteSpacePad / 4)) + 4) + """$Key"": false,`n"
                    }

                    elseif ($InputObject.$Key -is [DateTime] -and $Script:DateTimeAsISO8601) {
                        Write-Verbose -Message "Got a DateTime and will format it as ISO 8601."
                        $Json += " " * ((4 * ($WhiteSpacePad / 4)) + 4) + """$Key"": ""$($InputObject.$Key.ToString('yyyy\-MM\-ddTHH\:mm\:ss'))"",`n"
                
                    }

                    elseif ($InputObject.$Key -is [HashTable] -or $InputObject.$Key.GetType().FullName -eq "System.Management.Automation.PSCustomObject") {
                        Write-Verbose -Message "Input object's value for key '$Key' is a hash table or custom PowerShell object."
                        $Json += " " * ($WhiteSpacePad + 4) + """$Key"":`n$(" " * ($WhiteSpacePad + 4))"
                        $Json += ConvertToJsonInternal -InputObject $InputObject.$Key -WhiteSpacePad ($WhiteSpacePad + 4)
                    }

                    elseif ($InputObject.$Key.GetType().Name -match '\[\]|Array') {

                        Write-Verbose -Message "Input object's value for key '$Key' has a type that appears to be a collection/array."
                        Write-Verbose -Message "Building JSON for ${Key}'s array value."

                        $Json += " " * ($WhiteSpacePad + 4) + """$Key"":`n$(" " * ((4 * ($WhiteSpacePad / 4)) + 4))[`n" + (($InputObject.$Key | ForEach-Object {

                            if ($null -eq $_) {
                                Write-Verbose -Message "Got null inside array inside inside array."
                                " " * ((4 * ($WhiteSpacePad / 4)) + 8) + "null"
                            }

                            elseif ($_ -is [Bool] -and $_ -eq $true) {
                                Write-Verbose -Message "Got 'true' inside array inside inside array."
                                " " * ((4 * ($WhiteSpacePad / 4)) + 8) + "true"
                            }

                            elseif ($_ -is [Bool] -and $_ -eq $false) {
                                Write-Verbose -Message "Got 'false' inside array inside inside array."
                                " " * ((4 * ($WhiteSpacePad / 4)) + 8) + "false"
                            }

                            elseif ($_ -is [DateTime] -and $Script:DateTimeAsISO8601) {
                                Write-Verbose -Message "Got a DateTime and will format it as ISO 8601."
                                " " * ((4 * ($WhiteSpacePad / 4)) + 8) + """$($_.ToString('yyyy\-MM\-ddTHH\:mm\:ss'))"""
                            }

                            elseif ($_ -is [HashTable] -or $_.GetType().FullName -eq "System.Management.Automation.PSCustomObject" `
                                -or $_.GetType().Name -match '\[\]|Array') {
                                Write-Verbose -Message "Found array, hash table or custom PowerShell object inside inside array."
                                " " * ((4 * ($WhiteSpacePad / 4)) + 8) + (ConvertToJsonInternal -InputObject $_ -WhiteSpacePad ($WhiteSpacePad + 8)) -replace '\s*,\s*$'
                            }

                            else {
                                Write-Verbose -Message "Got a string or number inside inside array."
                                $TempJsonString = GetNumberOrString -InputObject $_
                                " " * ((4 * ($WhiteSpacePad / 4)) + 8) + $TempJsonString
                            }

                        }) -join ",`n") + "`n$(" " * (4 * ($WhiteSpacePad / 4) + 4 ))],`n"

                    }
                    else {

                        Write-Verbose -Message "Got a string inside inside hashtable or PSObject."
                        # '\\(?!["/bfnrt]|u[0-9a-f]{4})'

                        $TempJsonString = GetNumberOrString -InputObject $InputObject.$Key
                        $Json += " " * ((4 * ($WhiteSpacePad / 4)) + 4) + """$Key"": $TempJsonString,`n"

                    }

                }

                $Json = $Json -replace '\s*,$' # remove trailing comma that'll break syntax
                $Json += "`n" + " " * $WhiteSpacePad + "},`n"

            }

            $Json

        }
        $JsonOutput = ""
        $Collection = @()
        # Not optimal, but the easiest now.
        [Bool] $Script:CoerceNumberStrings = $CoerceNumberStrings
        [Bool] $Script:DateTimeAsISO8601 = $DateTimeAsISO8601
        [String] $Script:NumberRegex = '^-?\d+(?:(?:\.\d+)?(?:e[+\-]?\d+)?)?$'
        #$Script:NumberAndValueRegex = '^-?\d+(?:(?:\.\d+)?(?:e[+\-]?\d+)?)?$|^(?:true|false|null)$'

    }

    Process {

        # Hacking on pipeline support ...
        if ($_) {
            Write-Verbose -Message "Adding object to `$Collection. Type of object: $($_.GetType().FullName)."
            $Collection += $_
        }

    }

    End {
        
        if ($Collection.Count) {
            Write-Verbose -Message "Collection count: $($Collection.Count), type of first object: $($Collection[0].GetType().FullName)."
            $JsonOutput = ConvertToJsonInternal -InputObject ($Collection | ForEach-Object { $_ })
        }
        
        else {
            $JsonOutput = ConvertToJsonInternal -InputObject $InputObject
        }
        
        if ($null -eq $JsonOutput) {
            Write-Verbose -Message "Returning `$null."
            return $null # becomes an empty string :/
        }
        
        elseif ($JsonOutput -is [Bool] -and $JsonOutput -eq $true) {
            Write-Verbose -Message "Returning `$true."
            [Bool] $true # doesn't preserve bool type :/ but works for comparisons against $true
        }
        
        elseif ($JsonOutput-is [Bool] -and $JsonOutput -eq $false) {
            Write-Verbose -Message "Returning `$false."
            [Bool] $false # doesn't preserve bool type :/ but works for comparisons against $false
        }
        
        elseif ($Compress) {
            Write-Verbose -Message "Compress specified."
            (
                ($JsonOutput -split "\n" | Where-Object { $_ -match '\S' }) -join "`n" `
                    -replace '^\s*|\s*,\s*$' -replace '\ *\]\ *$', ']'
            ) -replace ( # these next lines compress ...
                '(?m)^\s*("(?:\\"|[^"])+"): ((?:"(?:\\"|[^"])+")|(?:null|true|false|(?:' + `
                    $Script:NumberRegex.Trim('^$') + `
                    ')))\s*(?<Comma>,)?\s*$'), "`${1}:`${2}`${Comma}`n" `
              -replace '(?m)^\s*|\s*\z|[\r\n]+'
        }
        
        else {
            ($JsonOutput -split "\n" | Where-Object { $_ -match '\S' }) -join "`n" `
                -replace '^\s*|\s*,\s*$' -replace '\ *\]\ *$', ']'
        }
    
    }

}
function ConvertFromJson20{ 
    [cmdletbinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string[]]$item
        
    )
    
    function IterateTree {
    [cmdletbinding()]
        param(
            $JsonTree
        )
          $result = @()
         foreach ($node in $jsonTree) {
            $nodeObj = New-Object -TypeName psobject
            foreach ($property in $node.Keys) {
                
                if ($node[$property] -is [System.Collections.Generic.Dictionary[String, Object]] -or $node[$property] -is [Object[]]) {
                    if ($node[$property] -is [Object[]] -and $node[$property][0] -is [string]){
                        $nodeObj  | Add-Member -MemberType NoteProperty -Name $property -Value $node[$property]
                    }
                    else{
                        $inner = @()
                        $inner += IterateTree -jsonTree $node[$property]
                        $nodeObj  | Add-Member -MemberType NoteProperty -Name $property -Value $inner
                    }
                    
                } else {
                    $nodeObj  | Add-Member -MemberType NoteProperty -Name $property -Value $node[$property]
                    #$nodeHash.Add($property, $node[$property])
                }
            }
            $result += $nodeObj
        }
        $result
    }
    try{
        add-type -assembly system.web.extensions
        $ps_js=new-object system.web.script.serialization.javascriptSerializer    
        #The comma operator is the array construction operator in PowerShell
        
        IterateTree -JsonTree $ps_js.DeserializeObject($item)
    } catch{
        Write-Error $_ -ErrorAction Stop
        
    }

        
}
function CreateJsonConfig{
    [cmdletbinding(SupportsShouldProcess=$true)]
    param(
        [parameter(Mandatory=$true)]
        [string]$FilePath,
        [parameter(Mandatory=$true)]
        [string[]]$DumpPaths,
        [psobject[]]$AllKeepSettings,
        [psobject]$DefaultKeepSettings,
        [int]$RecurseDeph,
        $FtpCredential
        
    )


    $JsonHashTable=@{
        "Paths"=$DumpPaths
        "RecurseDeph"=$RecurseDeph
        "Bases"=$AllKeepSettings
        "Default"=$DefaultKeepSettings
        "FtpCredential"=$FtpCredential
    }
    if ([version]$PSVersionTable.PSVersion -ge [version]"3.0"){
        $JsonHashTable | ConvertTo-Json | Out-File -FilePath $FilePath 
    } else{
        $JsonHashTable | ConvertToJson20 | Out-File -FilePath $FilePath 
    }
    
    
    
    
}
function ReadJsonConfig{
    [cmdletbinding()]
    Param(
        [parameter(Mandatory=$true)]
        [string]$FilePath
    )
    $JsonRaw=Get-Content -Path $FilePath -ErrorAction Stop
    if ([version]$PSVersionTable.PSVersion -ge [version]"3.0"){
        $JsonRaw | ConvertFrom-Json -ErrorAction Stop
    } else{
        ConvertFromJson20 -item $JsonRaw -ErrorAction Stop
    }
}
function CreateCredential{
    [cmdletbinding()]
    param(
        [parameter(Mandatory=$true)]
        [string]$User,
        [string]$Password
    )
    
    Write-Verbose "Create Credential User $User, Password $password"
    
    if ($PSBoundParameters["Password"])
    {
        $SecPassword = ConvertTo-SecureString -String $Password -AsPlainText -Force
        $Credential = New-Object System.Management.Automation.PSCredential($User,$SecPassword)  
    }
    else
    {
        $Credential = New-Object System.Management.Automation.PSCredential($User,(new-object System.Security.SecureString))
    }
    
    
    $Credential
}
function GetPlainTextPassword ($SecString){
    $BSTR =[System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecString)
    $PlainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
    $PlainPassword
}
function NewPassword{
    [cmdletbinding()]
    param(
        [int]$Length=18,
        [ValidateSet("Max",'Middle',"Min")]
        [string]$Сomplexity="Middle"
    )
    
    function Scramble-String([string]$inputString)
    {     
        $characterArray = $inputString.ToCharArray()   
        $scrambledStringArray = $characterArray | Get-Random -Count $characterArray.Length     
        $outputString = -join $scrambledStringArray
        return $outputString 
    }
    function Get-RandomCharacters($length, $characters) 
    { 
        $random = 1..$length | ForEach-Object { Get-Random -Maximum $characters.length } 
        $private:ofs="" 
        return [String]$characters[$random]
    }
    $password = Get-RandomCharacters -length $Length -characters 'abcdefghiklmnoprstuvwxyzABCDEFGHKLMNOPRSTUVWXYZ1234567890'
    $password += Get-RandomCharacters -length 1 -characters 'ABCDEFGHKLMNOPRSTUVWXYZ'
    $password += Get-RandomCharacters -length 1 -characters '1234567890'
    #$password += Get-RandomCharacters -length 1 -characters '!"§$%&/()=?}][{@#*+'

    $password = Scramble-String $password
    return $password
}
