if(!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) { Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs; exit }

cd $PSScriptRoot
cd ..
$MountPath=(get-location).Path+"\Mount"
$MountDri=$MountPath+"\Drivers"
$driverXML=$MountPath+"\tools\driver.xml"
$BackupDriver=(get-location).Path+"\BackupDriver"
$Backupwim=(get-location).Path+"\Backupwim"
$UndoXML=(get-location).Path+"\$UndoXML"


Function Deploy-Workpath{
    if( !( (Test-Path $BackupDriver) -and (Test-Path $Backupwim) -and (Test-Path $MountPath)  ) ){ 
        #New-Item -Path (Get-Location) -Name "Tools" -ItemType Directory -Force -ErrorAction SilentlyContinue |Out-Null
        New-Item -Path (Get-Location).Path -Name "BackupDriver" -ItemType Directory -Force -ErrorAction SilentlyContinue |Out-Null 
        New-Item -Path (Get-Location).Path -Name "Mount" -ItemType Directory -Force -ErrorAction SilentlyContinue |Out-Null
        New-Item -Path (Get-Location).Path -Name "Backupwim" -ItemType Directory -Force -ErrorAction SilentlyContinue |Out-Null}      
        
}

function MountWim{
    $GetWim=Dir $PSScriptRoot -Filter "*.wim" | Sort-Object LastWriteTime -Descending|Select-Object -f 1
    if($GetWim -ne $null){
        if(-not(Test-Path $MountDri)){
            Mount-WindowsImage –ImagePath $GetWim.FullName –Index 1 –Path $MountPath -Optimize |out-null }
    }
}

function Un-zipped{
    cd $PSScriptRoot
    [string]$CLI=(get-location).Path+"\CLI"
    $getZIP=dir "$PSScriptRoot\*.zip" |sort Length -Descending |select -First 1
    if($getZIP -ne $null){
        Expand-Archive $getZIP.fullname -DestinationPath $PSScriptRoot -Force -ErrorAction SilentlyContinue
        $chkdir=dir .\* -Include "*.inf","*.sys","*.cat","*.dll","*.exe"
        if($chkdir -ne $null){
            New-Item -Path . -Name "driver" -ItemType Directory -Force -ErrorAction SilentlyContinue
            Move-Item $chkdir -Destination ".\driver" -Force -ErrorAction SilentlyContinue
        }
        Move-Item $getZIP.fullname -Destination $BackupDriver -Force -ErrorAction SilentlyContinue
        $Unzipped=dir $PSScriptRoot -Recurse -Filter "*.inf" |sort Length -Descending |select -f 1
        $getCLI=dir $CLI -Filter "rstcli64.exe" -Recurse -Force -ErrorAction SilentlyContinue 
        if($getCLI -ne $null){
            Move-Item $getCLI.FullName -Destination $Unzipped.DirectoryName -Force -ErrorAction SilentlyContinue
            Remove-Item $getCLI.DirectoryName -Force -ErrorAction SilentlyContinue
        } 
        Write-Host "Extracts files from a specified (zipped) file has been completed  " -ForegroundColor Cyan
        Start-Sleep -s 2
    }else{
        Write-Host `nNot Found any .ZIP file`n  -ForegroundColor Cyan 
        Start-Sleep -s 2
    }
}

function Rename-Folder{
    $getInf=dir $PSScriptRoot -Recurse -Filter "*.inf" |sort Length -Descending |select -f 1
    if($getInf -ne $null){
            $Infcontent=Get-Content $getInf.FullName
            $Global:driverName=$getInf.BaseName
            :searchLoop foreach($Inf in $Infcontent){
                if($Inf -match "driverver"){
                    break searchLoop}
            }
        if($Inf -ne $null){
            $driverVer=$Inf.Substring(9).Replace('/','').Replace('=','').Replace(' ','').Trim()
            Rename-Item $getInf.Directory -NewName $driverVer -Force -ErrorAction SilentlyContinue | Out-Null
            return $driverVer        
        }
    }
}

function Get-Request{
    
    function get-systemid{
        Write-Host `nPlease Enter a SystemID.  "[Undo:U/Exit:Q]" -ForegroundColor Cyan 
        $Global:systemid=Read-Host [XXXX/XXXX]
        switch($systemid){
            U{
                cls
                get-systemid    
            }
            Q{
                EXIt}
        }
    }
    function get-drivername{
        Write-Host `nPlease Enter a Driver name.  "[Undo:U/Exit:Q]" -ForegroundColor Cyan 
        $Global:drivername=Read-Host [Driver name]
        switch($systemid){
            U{
                cls
                get-drivername     
            }
            Q{
                EXIt}
        }
    }
    function get-version{
        $driverFolder=Rename-Folder
        if($driverFolder -eq $null){
            Write-Host `nPlease Enter a Version"," No including date.   [Undo:U/Exit:Q]   -ForegroundColor Cyan
            $Global:version=Read-Host [Version]
            switch($version){
                U{
                    cls
                    get-version
                }
                Q{
                    EXIt
                }
            }
            $DriName=Dir $MountDri -Directory -Recurse
            foreach($Name in $DriName){
                if($Name -match $Global:version){
                    $Global:version=$Name.Name
                    break}
            }
        }else{
            $Global:Version=$driverFolder
        }
    }   
    get-systemid
    get-drivername
    get-version
    cls
}
function MoveToWim{
    $getInf=dir $PSScriptRoot -Recurse -Filter "*.inf" |sort Length -Descending |select -f 1
    $getFolder=($getInf.Directory).name
    if($getInf -ne $null){
        $action=$false
        $WimDriName=Dir $MountDri -Directory
        :searchLoop foreach($Inf in $WimDriName){
            if($Inf.Name -eq $getInf.BaseName){
                $action=$true}
            if($action){
                if(Test-Path "$MountDri\$Inf\$getFolder"){
                    Move-Item $getInf.DirectoryName -Destination $BackupDriver  -Force -ErrorAction SilentlyContinue
                    Remove-Item $getInf.DirectoryName -Force -ErrorAction SilentlyContinue
                }else{
                    Copy-Item $getInf.DirectoryName -Destination $BackupDriver -Recurse -Force -ErrorAction SilentlyContinue
                    Move-Item $getInf.DirectoryName -Destination "$MountDri\$Inf" -Force -ErrorAction SilentlyContinue
                }
                break searchLoop}
        }
        if(-not $action){
            $InfBaseName=$getInf.BaseName
            New-Item  $MountDri -Name $InfBaseName -ItemType Directory -Force |Out-Null
            Copy-Item $getInf.DirectoryName -Destination $BackupDriver -Recurse -Force -ErrorAction SilentlyContinue
            Move-Item $getInf.DirectoryName -Destination "$MountDri\$InfBaseName" -Force -ErrorAction SilentlyContinue
        }
    }
}

function XML-Process($ID,$DriName,$Ver){

  $Global:xml=[xml]$xml=get-content $driverXML
    
  function Add-sysid{
    param(
    [parameter(mandatory=$true)]
    [validateset("103C_53307F","103C_5335KV","103C_53311M")]
    [validateNotNullorEmpty()]
    [string]$Family_Code,
    [parameter(mandatory=$true,HelpMessage="[Project Name], For example:Faroe,Corvo, etc.")]
    [validateNotNullorEmpty()]
    [string]$Product_Description)
    $Global:Product_Description
#----------------Create Element----------------------
    $createElt=$xml.CreateElement("product")
    $createAtriName=$xml.CreateAttribute("description")
    $createAtriSystemid=$xml.CreateAttribute("systemid")
    $createAtriName.Value="$Product_Description"
    $createAtriSystemid.Value="$ID"
    $createElt.Attributes.Append($createAtriName) |Out-Null
    $createElt.Attributes.Append($createAtriSystemid) |Out-Null
#----------------Create Child Element----------
    $createChildElt=$xml.CreateElement("driver")
    $createChildAttriName=$xml.CreateAttribute("name")
    $createChildAttriVer=$xml.CreateAttribute("version")
    $createChildAttriName.Value=$DriName
    $createChildAttriVer.Value=$Ver
    $createChildElt.Attributes.Append($createChildAttriName) |Out-Null
    $createChildElt.Attributes.Append($createChildAttriVer) |Out-Null
    $familygroup=$xml.SelectSingleNode("//family[@code='$Family_Code']")
    $familygroup.AppendChild($createElt)|Out-Null
    $familygroup.LastChild.AppendChild($createChildElt) |Out-Null
    Write-Host `nInsert complete. -ForegroundColor Cyan
    Write-Host "`n<Product description=""$Product_Description"" systemid=""$ID"">`r`n  <driver name=""$DriName"" version=""$Ver"" />`n</Product>`n" 

}
    function second-process {
        Write-Host `n----------------------------------
        Write-Host "Mismatch systemid :$ID" -ForegroundColor Cyan
        Write-Host ------------------------------------
        Write-Host "Do you want to insert the mismath ID ?" -ForegroundColor Cyan
        $Answer=Read-Host "[Y/N]?"
        if($Answer -eq "Y"){
            Write-Host `nFamily code :  -BackgroundColor Black
            Write-Host --------------
            Write-Host 103C_53307F`n103C_5335KV`n103C_53311M  
            Write-Host --------------             
            Add-sysid
        }elseif($Answer -eq "N"){
            Write-host "`nEND Process.`n" -ForegroundColor Cyan
            Start-Sleep -s 2
        }else{
            cls
            second-process  
        }
}

    $MatchID=$xml.SelectSingleNode("//product[@systemid='$ID']")
    if($MatchID -ne $null){
        $MatchDri=$xml.SelectSingleNode("//product[@systemid='$ID']/driver[@name='$DriName']")
        if($MatchDri -ne $null){
            # Update version
            $MatchDri.version="$ver"
            Write-Host `nUpdate complete.  -ForegroundColor Cyan
            Write-Host `n"<driver name=""$DriName"" version=""$Ver"" />"

        }else{
            # Insert Driver 
            $creatElt=$xml.CreateElement("driver")
            $creatAtriName=$xml.CreateAttribute("name")
            $creatAtriVer=$xml.CreateAttribute("version")
            $creatAtriName.Value="$DriName"
            $creatAtriVer.Value="$Ver"
            $creatElt.Attributes.Append($creatAtriName) |Out-Null
            $creatElt.Attributes.Append($creatAtriVer) |Out-Null
            $MatchID.AppendChild($creatElt) |Out-Null
            write-host "`nAdd a new driver complete." -ForegroundColor Cyan
            write-host "`n<driver name=""$DriName"" version=""$Ver"" />"
        }
    }else{
        # Insert Product tag
        second-process
    }
    Copy-Item $driverXML -Destination $UndoXML -Force -ErrorAction SilentlyContinue 
    $xml.Save($driverXML)
    Copy-Item $driverXML -Destination $PSScriptRoot -Force -ErrorAction SilentlyContinue 
}
function Optimize-Wims{
    Cd $PSScriptRoot
    if(Test-Path $MountDri){
        $stringxml="_driver.xml"
        $date=Get-Date -UFormat "%Y%m%d"
        $uberini=Get-Content "$MountPath\UBER.INI" -ErrorAction SilentlyContinue
        foreach($ini in $uberini){
            if($ini -match "maintoolset"){
                $ini=$ini.substring(11).replace("=","").replace("WDT","").replace(".","")
                break
            }
        }
        $String="winpex64_wdt$ini"+"_"+$date+"v1"
        $SSRMstring=Test-Path "$MountPath\tools\automation"
        Dismount-WindowsImage -Path $MountPath -Save -ErrorAction SilentlyContinue |Out-Null
        $getWim=Get-ChildItem -Path "$PSScriptRoot\*.wim" | Sort-Object LastWriteTime -Descending|Select-Object -f 1
        #Move-Item $getWim.fullname -Destination $Backupwim -Force -ErrorAction SilentlyContinue 
        if(Test-Path .\driver.xml){
            Rename-Item -Path ".\driver.xml" -NewName "$String$stringxml" -ErrorAction SilentlyContinue
            Copy-Item ".\*driver.xml" -Destination $Backupwim -Force -ErrorAction SilentlyContinue
        }
        if($SSRMstring){
            $SSRMstring="_SSRM"
            #Export-WindowsImage -SourceImagePath $getWim.fullname -SourceIndex 1 -DestinationImagePath "$PSScriptRoot\$String$SSRMstring.wim" -CompressionType max |Out-Null 
            Rename-Item -Path $getWim.fullname -NewName "$String$SSRMstring.wim" -Force 
            $getWim=dir . -Filter "*.wim" | Sort-Object LastWriteTime -Descending|Select-Object -f 1
            #Copy-Item $getWim.fullname -Destination $Backupwim -Force -ErrorAction SilentlyContinue             
        }else{
            #Export-WindowsImage -SourceImagePath $getWim.fullname -SourceIndex 1 -DestinationImagePath "$PSScriptRoot\$String.wim" -CompressionType max |Out-Null 
            Rename-Item -Path $getWim.fullname -NewName "$String.wim" -Force             
            #Copy-Item $getWim.fullname -Destination $Backupwim -Force -ErrorAction SilentlyContinue             
        }
        $getWim=Get-ChildItem -Path "$PSScriptRoot\*.wim" | Sort-Object LastWriteTime -Descending|Select-Object -f 1
        Copy-Item $getWim.fullname -Destination $Backupwim -Force -ErrorAction SilentlyContinue
     }
     cd ..
}



<# End Sub Function  #>
<# Main Function #>
function Auto-Update{
    MountWim
    Un-zipped
    Get-Request
    MoveToWim
    XML-Process -ID $Global:systemid -DriName $Global:drivername -Ver $Global:version
}



Deploy-Workpath

for($m=0;$m -le 50;$M++){

Write-Host **********************************************
Write-Host "        "Auto WDT Driver Update v4.0         
write-host **********************************************
write-host
write-host "(1)Auto Update"
Write-Host "(2)Commit"
write-host "(3)Unzip"
write-host "(4)Mount"
write-host "(5)Discard"
write-host "(6)Restore driver.xml(TBD.)"
write-host "(Q)Quit"


$act=Read-Host `n`n`nKey in your selection

switch($act){
    1{
        cls
        Auto-Update
        pause
        cls
    }
    2{
        cls
        Optimize-Wims
        cls
    }
    3{
        cls
        Un-zip
        cls
    }
    4{
        MountWim
        cls
    }
    5{
        if(Test-Path $MountPath){
            Dismount-WindowsImage -Path $MountPath -Discard -ErrorAction SilentlyContinue |Out-Null
            Write-Host `nDismount has been completed. -ForegroundColor Cyan
        }else{
            Write-Host """Error""". -ForegroundColor Red

        }
        cls
    }
    Q{
        exit
    }
    default{
        write-host `nEntered a Incorrect selection"," Restarting this process`n  -ForegroundColor red
        pause 
        cls  
    }
}
}