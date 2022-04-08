try{
    $source=Get-Content ".\config.ini" -ErrorAction Stop
}catch{
    [System.Windows.Forms.MessageBox]::Show("Missing the ini file" , "Error")
    exit
}


function get-setting($arg,$source){
   $pattern=".*($arg).*"
   $value= $source | Select-String -Pattern $pattern |Out-String
   $value=$value.Replace($arg,"").Trim()
   return $value
}



Function Get-BlobLocation($Size,$src_path){
    $Hex=[System.Convert]::ToString($Size,16)
    foreach($Cnt in 1..8){
       if($Hex -notmatch "\w{8}"){
           $Hex="0"+$Hex
       }else{
            break
       }
    }
    #$src_path=get-setting -arg "src_path=" -source $source
    $Hex=$Hex.Insert(0,'\').Insert(3,'\').Insert(6,'\').Insert(9,'\').Insert(0,"$src_path")
    Return $Hex
}

function Connect-DB($T_SQL){
    $connection=New-Object System.Data.SqlClient.SqlConnection("Server=$server_ip;Database=$server_db;User ID=sa;Password=$server_pwd;")
    try{
        $connection.Open()
    }catch{
        return "Stop" 
    }
    $SqlCommand=New-Object System.Data.SqlClient.SqlCommand($T_SQL,$connection)
    $adapter = New-Object System.Data.sqlclient.sqlDataAdapter $SqlCommand
    $dataset = New-Object System.Data.DataSet
    $adapter.Fill($dataset) | Out-Null 
    $connection.Close()
    return $dataset
}
$server_ip = get-setting -arg "IP=" -source $source
$server_db = get-setting -arg "DB=" -source $source
$server_pwd = get-setting -arg "PWD=" -source $source

$FileServerName=get-setting -arg "FileServerName=" -source $source
$src_path=get-setting -arg "src_path=" -source $source
$dest_path=get-setting -arg "dest_path=" -source $source
$InsertLength=($src_path+"\00\00\00\00").Length

#-------------------------------------------Check BlolbsToTransfer table---------------------------------

$Command="select Size,md5 from BLObject INNER join BlobsToTransfer on BLObject.DataKey=BlobsToTransfer.DataKey where FileServerName = '$FileServerName';"
$Result=Connect-DB $Command

#------------------------------------------Find Location and FileName hex---------------------------------

IF($Result.Tables[0].Rows.size -ne $null){
            "--------------------------------------------------------------------"
Write-Host  "There are the stuck blobs in the Blobstotransfer table of $FileServerName."
            "--------------------------------------------------------------------`n"
Write-Host  "Transfering" $Result.Tables[0].Rows.size.count "Blobs from $src_path to $dest_path is Starting"

   
[Array]$Result_Hex = $Result.Tables[0].Rows.md5 | Format-Hex  | Select-String "\d*"
[Array]$Result_Size = $Result.Tables[0].Rows.Size

$SrcPath=@()
$DestPath=@()
foreach($size in $Result_Size){
    [Array]$SrcPath=$SrcPath+(Get-BlobLocation $size $src_path)
    [Array]$DestPath=$DestPath+(Get-BlobLocation $size $src_path).Replace("$src_path","$dest_path")
}

$FileName = @()
For($i=0;$i -lt $Result_Hex.Length; $i++){
    [array]$FileName=$FileName+$($Result_Hex[$i].ToString().Substring(11,47).replace(' ','')+".dat" )
}

#---------------------------------------------Find Absolute path------------------------------------------
$AbsSrc_Path = @()
$AbsDest_Path = @()

for($i=0;$i -lt $SrcPath.Count;$i++){
    [array]$AbsSrc_Path=$AbsSrc_Path+$SrcPath[$i].Insert($InsertLength,("\"+$FileName[$i]))
}

for($i=0;$i -lt $DestPath.Count;$i++){
    #[array]$AbsDest_Path=$AbsDest_Path+$DestPath[$i].Insert(32,("\"+$FileName[$i]))
    [array]$AbsDest_Path=$AbsDest_Path+$DestPath[$i]
}


foreach($P in $AbsDest_Path){
    if(!(Test-Path $P)){
        #create a folder if the destination was not existed
        New-Item -Path $P -ItemType Directory -Force 
    }
}



#---------------------------------------------Copy to Dest from Src----------------------------------------


foreach($F in $AbsSrc_Path){
    Try{
        Copy-Item -Path $F -Destination ($F.Replace("$src_path","$dest_path")) -Force -PassThru
    }Catch{
        $Error[0].Exception
    }
}

            "--------------------------------------------------`n"
Write-Host  "Transfer complete."

}else{
               cls
               "************************************************************`n"
    Write-Host "There is no a stuck blob in the Blobstotransfer table of $FileServerName.`n" -ForegroundColor White
               "************************************************************`n"
               "Stop process.`n"
}

Pause

