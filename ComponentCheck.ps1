<#
Author: Duke Yu
Purpose: Component query
Last written time: 2021/05/20
Problem:1.fix SQL syntax happened on Win 8
        2.Sync syntax has something not correct (FIXED)
        3.optimize PML check function
#>
# Hide PowerShell Console
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();
[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'
$consolePtr = [Console.Window]::GetConsoleWindow()
[Console.Window]::ShowWindow($consolePtr, 0)

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()
cd $PSScriptRoot
try{
    $source=Get-Content ".\Configuration\config.ini" -ErrorAction Stop
}catch{
    [System.Windows.Forms.MessageBox]::Show("Missing the Setting directory" , "Error")
    exit
}
function get-setting($serverName,$source){
   $pattern=".*($serverName).*"
   $value= $source | Select-String -Pattern $pattern |Out-String
   $value=$value.Replace($serverName,"").Trim()
   return $value
}
function get-SkuNumber{
    if( ($TextBox1.Text -match "\w{6}\-\w{3}" -or $TextBox1.Text -match ".{10,15}" ) ){
        $Global:Component=""
        $Global:Component=($TextBox1.Text).Trim()
        $TextBox1.Text=""
    }
    else{
        [System.Windows.Forms.MessageBox]::Show("Incorrect SKU or ML. Please check the spell." , "Error")
    }
}
function Ready-check($SkuNumber){ 
    return ("declare @Component varchar(14)='$SkuNumber';`n"+`
            "select Sku.SkuNumber,sku.Revision,skuok.InsertedTime from SKU `n"+`
            "inner join SkuOK on sku.SkuKey=skuok.SkuKey `n"+`
            "where sku.SkuNumber =@Component and `nsku.Revision=(select MAX(Revision) from sku where sku.SkuNumber=@Component);")
}
<#function Ready-check($SkuNumber){ 
    return ("declare @skukeyOK binary(8),@skukeyOKFS binary(8),@ServerCount int,@Component varchar(14)="+"'"+$SkuNumber+"'"+";`n"+` 
    "select @ServerCount=count(FileServerName),@skukeyOKFS=temp.SkuKey from `n"+` 
    " (SELECT top(3) skuokfs.*FROM SkuOKFS INNER JOIN Sku ON SkuOKFS.SkuKey = Sku.SkuKey"+` 
    "`n   WHERE Sku.SkuNumber = @Component order by sku.Revision desc) as Temp group by SkuKey;"+"`n"+`
    "select @skukeyOK=SkuKey from SkuOK where SkuKey=@skukeyOKFS;"+"`n"+`
    "if(@ServerCount >= 2 AND @skukeyOK = @skukeyOKFS)"+"`n"+`
    "SELECT top(3) Sku.SkuNumber,FileServerName,InsertedTime FROM SkuOKFS INNER JOIN Sku ON SkuOKFS.SkuKey = Sku. SkuKey"+"`n"+` 
    "  WHERE Sku.SkuNumber ="+" '"+$SkuNumber+"' "+"order by sku.Revision desc"+"`n"+`
    "else"+"`n"+`
    "select 'Not Ready' as Result;")
}#>
function Execute-spCheckML($ML){
    return "exec spCheckML $ML"
}


function Connect-DB($T_SQL,$server,$DB,$ID,$PWD){
        $dataset=""
        $adapter=""
        $connection=""
        $connection=New-Object System.Data.SqlClient.SqlConnection("Server=$server;Database=$DB;User ID=$ID;Password=$PWD;")
        #$connection=[System.Data.SqlClient.SqlConnection]::new("Server=10.1.1.3;Database=$DB;User ID=$ID;Password=$PWD;")
        try{
            $connection.Open()
        }catch{
            #[System.Windows.Forms.MessageBox]::Show("Database connection Failed. " , "Error")
            return "Stop" 
        }
        #$SqlCommand=[System.Data.SqlClient.SqlCommand]::new($T_SQL,$connection)
        $SqlCommand=new-object System.Data.SqlClient.SqlCommand($T_SQL,$connection)
        $adapter = New-Object System.Data.sqlclient.sqlDataAdapter $SqlCommand
        $dataset = New-Object System.Data.DataSet
        $adapter.Fill($dataset) | Out-Null 
        $connection.Close()
        return $dataset
    }
function Get-Results{
    if($RadioButton1.Checked -eq $true){
        $getSkuNumber=get-SkuNumber
        if($getSkuNumber -eq "OK"){
            return
        }else{
        #P00XKK-B2N
            $Query=Ready-check $Component
            $name=$combobox1.SelectedItem
            $TargetServer=get-setting ($name+"_DB=") $source
            $TargetDB=get-setting ($name+"_DB=") $source
            $TargetID=get-setting ($name+"_ID=") $source
            $TargetPWD=get-setting ($name+"_Password=") $source
            $Global:Results=Connect-DB $Query $TargetServer $TargetDB $TargetID $TargetPWD
            $table=$Results.Tables |Out-String
            if($Results -match "stop"){
                [System.Windows.Forms.MessageBox]::Show("Database connection Failed. " , "Error")
            }else{
                #$table=$Results.Tables | Out-String
                if($table -eq ""){
                    [System.Windows.Forms.MessageBox]::Show("$Component  not ready yet.","Message")
                }else{
                    switch($combobox1.Text){
                        "TAIWISDIV01"{
                                        
                                            $TextBox2.Text=$table
                                            [System.Windows.Forms.MessageBox]::Show("The component is ready to DASH.","Message") 
                                        
                        }
                        "TAIWISDIV02"{ 
                                            $TextBox2.Text=$table
                                            [System.Windows.Forms.MessageBox]::Show("The component is ready to DASH.","Message") 

                         }
                     }
                }

            }
        }
    }elseif($RadioButton2.Checked -eq $true){
        $getSkuNumber=get-SkuNumber
        if($getSkuNumber -eq "OK"){
            return
        }elseif($Component -match "\w{6}\-\w{3}"){
            [System.Windows.Forms.MessageBox]::Show("Incorrect ML. " , "Error")
            return 
        }else{
            $Query=Execute-spCheckML $Component
            $name=$combobox1.SelectedItem
            $TargetServer=get-setting ($name+"_DB=") $source
            $TargetDB=get-setting ($name+"_DB=") $source
            $TargetID=get-setting ($name+"_ID=") $source
            $TargetPWD=get-setting ($name+"_Password=") $source
            $Global:Results=Connect-DB $Query $TargetServer $TargetDB $TargetID $TargetPWD
            if($Results -match "stop"){
                [System.Windows.Forms.MessageBox]::Show("Database connection Failed. " , "Error")
            }else{
                $table=$Results.Tables | Out-String
                if($table -eq ""){
                    [System.Windows.Forms.MessageBox]::Show("Please try again later or check the spell.","Message")
                }else{
                    if($Results.Tables[0].Buildable -eq 1 ){
                        $TextBox2.Text=($Results.Tables[0] | Out-String)#+($Results.Tables[1] | Where-Object {$_.Inbuildplan -ne 1} |Out-String) #+($Results.Tables[2].Message)
                        [System.Windows.Forms.MessageBox]::Show("The ML is buildable.","Message")
                    }else{
                        $TextBox2.Text=($Results.Tables[0] | Out-String)+($Results.Tables[1] | Where-Object {$_.Inbuildplan -ne 1} |Out-String) #+($Results.Tables[2].Message)
                    }
                }

            }
        }
    }elseif($RadioButton3.Checked -eq $true){
        # Radio Button 3  portion
        
        function EOL-check($component){
        return("declare @Component varchar(14)="+"'$component'"+",@ReadyTime datetime,@EOL datetime;`n"+`
                "select @EOL=EndofLife from SKU where SkuNumber=@Component; 
                select @ReadyTime=skuok.InsertedTime from SKU
                inner join SkuOK on sku.SkuKey=skuok.SkuKey  
                where sku.SkuNumber =@Component and 
                sku.Revision =(select MAX(Revision) from sku where sku.SkuNumber=@Component);
                if(@EOL >= GETDATE())
                    begin
	                    if(@ReadyTime is not null)
		                    select 'Ready on '+convert(varchar,@ReadyTime,111) as Result;
	                    else
		                    select 'Not ready yet.' as Result;
                        end
                else
                    if(@EOL <= GETDATE() AND @EOL is not null)   
	                    select 'EOL' as Result;
                    else
                        select 'Not in the server.' as Result;")
        <#
            return ("declare @EOL date;"+"`n"+`
                    "select top 1 @EOL=EndofLife from sku where SkuNumber="+"'"+$component+"'"+ " order by Revision desc;" + "`n"+`
                    "if(@EOL is not null)"+"`n"+`
                    "begin"+"`n"+`
                        "if(@EOL > GETDATE())"+"`n"+`
	                        "select 'OK' as Result ;"+"`n"+`
                        "else"+"`n"+`
	                        "select 'EOL ' as Result;"+"`n"+`
                    "end"+"`n"+`
                    "else"+"`n"+`
                        "select 'Not in the server' as Result;")#>
        }

        function Get-IniContent ($filePath){
            $ini = @{}
            switch -regex -file $FilePath{
                “^\[(.+)\]” # Section
                {
                    $section = $matches[1]
                    $ini[$section] = @{}
                    $CommentCount = 0
                }
                <#“^(;.*)$” # Comment
                {
                    $value = $matches[1]
                    $CommentCount = $CommentCount + 1
                    $name = “Comment” + $CommentCount
                    $ini[$section][$name] = $value
                }#>
                “(Blk_\d+?)=(.*)” # Key
                {
                    $name,$value = $matches[1..2]
                    $ini[$section][$name] = $value
                }
            }
            return $ini
        }

        function get-InIfile{
            $Global:getINIfile=Dir .\*.ini
        }
        get-InIfile
        if($getINIfile -ne $null){           
            $chkResult=""
            foreach($file in $getINIfile){
                $INIcontent=Get-IniContent ($file.fullname)
                $CPNvalue=$INIcontent['FILESETS'].Values
                foreach($val in $CPNvalue){
                    $Query=EOL-check $val
                    $name=$combobox1.SelectedItem
                    $TargetServer=get-setting ($name+"_DB=") $source
                    $TargetDB=get-setting ($name+"_DB=") $source
                    $TargetID=get-setting ($name+"_ID=") $source
                    $TargetPWD=get-setting ($name+"_Password=") $source
                    $Global:Results=Connect-DB $Query $TargetServer $TargetDB $TargetID $TargetPWD
                    if($Results -eq "stop"){
                        [System.Windows.Forms.MessageBox]::Show("Database connection Failed. " , "Error")
                        return
                    }else{
                        if($val.Length -gt 10){
                            $table=$val+ "     ...  " + ($Results.Tables).Result
                            $chkResult=$chkResult+$table
                        }else{
                            $table=$val+ "         ...  " + ($Results.Tables).Result
                            $chkResult=$chkResult+$table
                        }
                        $TextBox2.Text=$TextBox2.Text+$table+[System.Environment]::NewLine
                    }
                }
                Move-Item -Path $file.fullname -Destination ".\PML_out" -Force 
            }
        }else{
            [System.Windows.Forms.MessageBox]::Show("PML was not found.`nPlease put the PML into directory the same as the script.","Error")
        }
        
    }
}

       


$Form                            = New-Object system.Windows.Forms.Form
$Form.Size                       = New-Object System.Drawing.Point(620,510)
$Form.text                       = "Component Check"
$Form.TopMost                    = $false
$Form.BackColor                  = [System.Drawing.ColorTranslator]::FromHtml("#ffffff")
$Form.AutoSize                   = $false
$Form.AcceptButton               = $Button1
$Form.MaximumSize                = New-Object System.Drawing.Point(620,510)
$Form.MinimumSize                = New-Object System.Drawing.Point(620,510)
$Form.StartPosition              ="CenterScreen"

$RadioButton1                    = New-Object system.Windows.Forms.RadioButton
$RadioButton1.text               = "SKU Ready Check"
$RadioButton1.AutoSize           = $true
$RadioButton1.width              = 104
$RadioButton1.height             = 20
$RadioButton1.location           = New-Object System.Drawing.Point(20,10)
$RadioButton1.Font               = New-Object System.Drawing.Font('Calibri',13,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Italic) )
$RadioButton1.Checked            = $true
$RadioButton1.TabStop            = $true

$RadioButton2                    = New-Object system.Windows.Forms.RadioButton
$RadioButton2.text               = "ML Check"
$RadioButton2.AutoSize           = $true
$RadioButton2.width              = 104
$RadioButton2.height             = 20
$RadioButton2.location           = New-Object System.Drawing.Point(20,40)
$RadioButton2.Font               = New-Object System.Drawing.Font('Calibri',13,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Italic))
$RadioButton2.Checked            = $false
$RadioButton2.TabStop            = $true

$RadioButton3                    = New-Object system.Windows.Forms.RadioButton
$RadioButton3.text               = "PML Check"
$RadioButton3.AutoSize           = $true
$RadioButton3.width              = 104
$RadioButton3.height             = 20
$RadioButton3.location           = New-Object System.Drawing.Point(20,70)
$RadioButton3.Font               = New-Object System.Drawing.Font('Calibri',13,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Italic))
$RadioButton3.Checked            = $false
$RadioButton3.TabStop            = $true

  
$TextBox1                        = New-Object system.Windows.Forms.TextBox
$TextBox1.multiline              = $false
$TextBox1.width                  = 160
$TextBox1.height                 = 40
$TextBox1.Anchor                 = 'Top, Left'
$TextBox1.location               = New-Object System.Drawing.Point(260,40)
$TextBox1.Font                   = New-Object System.Drawing.Font('Calibri',13)
$TextBox1.ForeColor              = [System.Drawing.ColorTranslator]::FromHtml("#000000")
$TextBox1.AcceptsTab             = $true
$TextBox1.enabled                = $true
$TextBox1.TabIndex               = 3


$TextBox2                        = New-Object system.Windows.Forms.TextBox
$TextBox2.TextAlign              ='center'
$TextBox2.multiline              = $true
$TextBox2.width                  = 605
$TextBox2.height                 = 312
$TextBox2.enabled                = $true
$TextBox2.location               = New-Object System.Drawing.Point(0,160)
$TextBox2.Font                   = New-Object System.Drawing.Font('Lucida Console',12)
$TextBox2.ScrollBars             = "Vertical"
$TextBox2.ReadOnly               = $true
$textBox2.TextAlign              = "left"
$TextBox2.AcceptsTab             = $false
$TextBox2.TabStop                = $false
$TextBox2.Multiline              = $true

$Label1                          = New-Object System.Windows.Forms.Label
$Label1.Text                     = "SERVER Option"
$Label1.AutoSize                 = $true
$Label1.Font                     = New-Object System.Drawing.Font('Calibri',12)
$Label1.Location                 = New-Object System.Drawing.Point(260,73)
$Label1.ForeColor                = [System.Drawing.ColorTranslator]::FromHtml("#f60606")

$Label2                          = New-Object System.Windows.Forms.Label
$Label2.Text                     = "SKU Number"
$Label2.AutoSize                 = $true
$Label2.Font                     = New-Object System.Drawing.Font('Calibri',12)
$Label2.Location                 = New-Object System.Drawing.Point(260,10)
$Label2.ForeColor                = [System.Drawing.ColorTranslator]::FromHtml("#f60606")

$combobox1                       = New-Object System.Windows.Forms.ComboBox
$combobox1.Location              = New-Object System.Drawing.Point(260,103)
@('TAIWISDIV01','TAIWISDIV02','TAIWISDIV03','CHOWISDIV01') | ForEach-Object {[void] $ComboBox1.Items.Add($_)}
$ComboBox1.Font                  = New-Object System.Drawing.Font('Calibri',14)
$combobox1.AutoSize              = $false
$combobox1.Sorted                = $true
$combobox1.Width                 = 160
$combobox1.Height                = 40
$combobox1.TabIndex              = 3
$combobox1.DropDownStyle         = 'DropDownList'
$combobox1.SelectedIndex         = 2

$Button1                         = New-Object system.Windows.Forms.Button
$Button1.text                    = "RUN"
$Button1.width                   = 100
$Button1.height                  = 100
$Button1.TabIndex                = 4
$Button1.location                = New-Object System.Drawing.Point(460,30)
$Button1.Font                    = New-Object System.Drawing.Font('Calibri',14)
$Button1.Add_Click(
        {   
            $TextBox2.Text=""
            Start-Sleep -s 0.6
            Get-Results
        }
    )
$Form.controls.AddRange(@($RadioButton1,$RadioButton2,$Button1,$TextBox1,$TextBox2,$combobox1,$Label1,$Label2,$RadioButton3))
$Form.ShowDialog()