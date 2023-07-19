Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName system.drawing
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()
#[Windows.Forms.Application]::EnableVisualStyles()
#Import-Module SQLPS -DisableNameChecking
Import-Module sqlserver -DisableNameChecking
$Dir_Path ="C:\Axalta_Assessment_Report\"

#$Dir_Path ="C:\myDB_Assessment_Report\"
$DMA_Path =$Dir_Path


Function Directory_Creation
{
  Add-Type -AssemblyName PresentationFramework
  $Dir_Path ="C:\Axalta_Assessment_Report"+'\'
       
        if(!(Test-Path -Path $Dir_Path)) 
        {
          #New-Item "$InstallDir" -type directory -Force | out-null
            new-item -type directory -path $Dir_Path -Force
            new-item -type directory -path $Dir_Path"DMA" -Force
            new-item -type directory -path $Dir_Path"SKU" -Force
            new-item -type directory -path $Dir_Path"Oracle" -Force
            new-item -type directory -path $Dir_Path"SingleInstance" -Force
            new-item -type directory -path $Dir_Path"SingleInstance_IaaS" -Force
            new-item -type directory -path $Dir_Path"Reports" -Force
            new-item -type directory -path $Dir_Path"Inventory" -Force
            new-item -type directory -path $Dir_Path"Image" -Force
            new-item -type file -path $Dir_Path"Inventory"'\'"MS_Instancelist.txt" -Force
            new-item -type file -path $Dir_Path"Inventory"'\'"server_not_connect.txt" -Force
            new-item -type file -path $Dir_Path"Inventory"'\'"server_connect.txt" -Force
            new-item -type file -path $Dir_Path"Inventory"'\'"serverlist.txt" -Force
            new-item -type file -path $Dir_Path"Inventory"'\'"inventory_list.txt" -Force
            new-item -type file -path $Dir_Path"Inventory"'\'"single_Instance.txt" -Force
            new-item -type file -path $Dir_Path"Inventory"'\'"single_Instance_IaaS.txt" -Force
            new-item -type file -path $Dir_Path"Inventory"'\'"Assess-for-AzureSQLMI.xml" -Force
            
            <#
            Start-Sleep 5
            write-host "Content adding on XML file " -ForegroundColor Yellow
            #Add-Content $Dir_Path"MS_Instancelist.txt" $sname
            start-sleep 3
            Add-Content $Dir_Path"Inventory"'\'"Assess-for-AzureSQLMI.xml" -Value '<?xml version="1.0" encoding="UTF-8"?>
            <AssessmentConfiguration xmlns="http://microsoft.com/schemas/SqlServer/Advisor/AssessmentConfiguration">
            <AssessmentName>Scale-Assessment-for-AzureSQLManagedInstance</AssessmentName>
            <AssessmentSourcePlatform>SqlOnPrem</AssessmentSourcePlatform>
            <AssessmentTargetPlatform>ManagedSqlServer</AssessmentTargetPlatform>
            <AssessmentDatabases>
            <AssessmentDatabase>Server=sqlserver01;Integrated Security=true</AssessmentDatabase>
            </AssessmentDatabases>
            <AssessmentResultDma>C:\DMA_Reports\sqlserver01.dma</AssessmentResultDma>
            <AssessmentResultJson>C:\DMA_Reports\Scale-Assessment-for-AzureSQLManagedInstance1.json</AssessmentResultJson>
            <AssessmentResultCsv>C:\DMA_Reports\Scale-Assessment-for-AzureSQLManagedInstance1.csv</AssessmentResultCsv>
            <AssessmentOverwriteResult>true</AssessmentOverwriteResult>
            <AssessmentEvaluateCompatibilityIssues>true</AssessmentEvaluateCompatibilityIssues>
            <AssessmentEvaluateFeatureParity>true</AssessmentEvaluateFeatureParity>
          </AssessmentConfiguration>'
          write-host "Content added on XML file " -ForegroundColor Green
          #>
          write-host "Directory $Dir_Path and file created sucessfully " -ForegroundColor Yellow
          [void] [System.Windows.MessageBox]::Show( "Directory $Dir_Path Created,Please enter Server name ", "Script completed", "OK", "Information" )
          start-sleep 2
          #exit;
        }
        else 
        {
          write-host "Folder $Dir_Path already exists" -ForegroundColor Green
          #[void] [System.Windows.MessageBox]::Show( "Directory $Dir_Path Already Created,Please check Servers on serverlist.txt before run DMA Report ", "Script completed", "OK", "Information" )
          

        }

}



Function SQL_SERVER_Validation
{
  param (
        [Parameter(Mandatory=$true)][string]$s
    )
  
    $connect=$Dir_Path+"Inventory"+'\'+"server_connect.txt"
    $notconnect=$Dir_Path+"Inventory"+'\'+"server_not_connect.txt"
    $srvoutput=$Dir_Path+"Inventory"+'\'+"MS_Instancelist.txt"

    
  TRY
  {
    $inst_name = (Get-Service -Name MSSQLSERVER,MSSQL$* -ComputerName $s | Where-Object {($_.Status -eq "Running")} -ErrorAction SilentlyContinue)
      $inst = ($inst_name).name -replace 'MSSQL\$'

    
       
      foreach ($i in $inst)
      {
      if($i -eq 'MSSQLSERVER')
        {
         #write-host "Default instance loop" -ForegroundColor Yellow

          $servername = New-Object Microsoft.SqlServer.Management.Smo.Server -ArgumentList $s
        
          $srv=$servername.Name
          #$srv
          $name = $s | Out-File -FilePath $srvoutput -Append

          #Write-Host "$srv,up" -ForegroundColor Green
          "$s,up" | Out-file $connect -Append
          #[void] [System.Windows.MessageBox]::Show( "Server $s is up and running ", "Script completed", "OK", "Information" ) 
          ########################
        } # end if
        else{        
          #$name = $s+'\'+$i | Out-File -FilePath C:\DMA_Reports_Final_named\MS_Instancelist.txt -Append 
          $srv1= $s+'\'+$i 
          #"Named Instance  " +$srv1
          $servername1=$null   
          $servername1 = New-Object Microsoft.SqlServer.Management.Smo.Server -ArgumentList $srv1
           # insert for SKU report
           $name = $srv1 | Out-File -FilePath $srvoutput -Append          
           # Insert for sucess logs
          #Write-Host "$srv1,up" -ForegroundColor Green
          "$srv1,up" | Out-file $connect -Append
          #write-host "Nameed instance loop" -ForegroundColor Yellow
              #$d=$null
              #$d=$srv1.replace('\','_')
              #[void] [System.Windows.MessageBox]::Show( "Named Instance $srv1 is up and running ", "Script completed", "OK", "Information" )
      } # end else
    #} #end TRY
   
     
    } # end foreach
    $true
    #[void] [System.Windows.MessageBox]::Show( "Server $s is up and running ", "Script completed", "OK", "Information" ) 
     }
    catch 
    {
              $err_text = $_.Exception.Message
              #Write-Host "$s Server is not connecting" -ForegroundColorgreen
               #Write-Host "$S,down" -ForegroundColor red
               "$S,down,$err_text" | Out-file $notconnect -Append
                $false
    
    }
      
   #[void] [System.Windows.MessageBox]::Show( "Server $s is down , Unable to connect ", "Script completed", "OK", "Information" )
    
 
  } # end function
    #SQL_SERVER_Validation -s 'sqlserver01\sql_002'

Function DMA-WindowSubForm-PaaS
{
  param (
        [Parameter(Mandatory=$false)][string]$DMA_Opt,
        [Parameter(Mandatory=$false)][string]$DMAReports
    )
  Add-Type -AssemblyName System.Windows.Forms
  Add-Type -AssemblyName System.Drawing

  write-host "This is options $DMA_Opt"
  # Main Form 
  $mainFormSI = New-Object System.Windows.Forms.Form
  $mainFormSI.Font = $header#"Comic Sans MS,8.25"
  $mainFormSI.Text = " DMA PaaS Msgbox"
  $mainFormSI.FormBorderStyle = "FixedDialog"
  $mainFormSI.ForeColor = "white"
  $mainFormSI.BackColor = "Darkblue"
  $mainFormSI.StartPosition = "CenterParent"
  $mainFormSI.width = 600
  $mainFormSI.height = 250
 
  # Title Label
  $titleLabel = New-Object System.Windows.Forms.Label
  $titleLabel.Font = $header#"Comic Sans MS,14"
  $titleLabel.ForeColor = "Yellow"
  $titleLabel.Location = "30,20"
  $titleLabel.Size = "500,50"
  #$titleLabel.Text = "Enter SQL Server Instance name"
  $mainFormSI.Controls.Add($titleLabel);
  #$mainFormSI.Controls.Add($titleLabel)

  

  # Input Box
  $textBoxIn = New-Object System.Windows.Forms.TextBox
  $textBoxIn.Location = "35, 70"
  $textBoxIn.Size = "500, 20"
  $textBoxIn.Text = ""
  $mainFormSI.Controls.Add($textBoxIn)
 
  #Combobox
  $the_combo = New-Object system.Windows.Forms.ComboBox
  $the_combo.location = "35, 100"

  $the_combo.Size = "200, 20"

  $the_combo.DropDownStyle = "Dropdownlist"
  #$ComboList_Items = Get-Content $DMA_Path"DMA_DB_Type.txt"
  If ($DMA_Opt -eq 'DMA_ALL_PaaS')
 {
  $ComboList_Items = @("Azuresqldatabase", "ManagedSqlServer" , "Both")
   $DMA_Opt=$null
   $titleLabel.Text = "Select Target platform for DMA Assessment"
   $textBoxIn.Visible = $false
   $the_combo.location = "35, 75"
 }
  If ($DMA_Opt -eq 'DMA_Single_PaaS')
 {
  $ComboList_Items = @("Azuresqldatabase", "ManagedSqlServer" , "Both")
  $DMA_Opt=$null
  $titleLabel.Text = "Select Target platform for DMA Assessment"
 }

   If ($DMA_Opt -eq 'DMA_All_IaaS')
  {
  $ComboList_Items = @("Sqlserver2012", "Sqlserver2014" ,"Sqlserver2016","SqlServerWindows2017","SqlServerLinux2017","SqlServerWindows2019","SqlServerLinux2019")
  $DMA_Opt=$null
  $titleLabel.Text = "Select Target platform for DMA Assessment"
  $textBoxIn.Visible = $false
  }
   If ($DMA_Opt -eq 'DMA_Single_IaaS')
  {
  $ComboList_Items = @("Sqlserver2012", "Sqlserver2014" ,"Sqlserver2016","SqlServerWindows2017","SqlServerLinux2017","SqlServerWindows2019","SqlServerLinux2019")
  $DMA_Opt=$null
  $titleLabel.Text = "Select Target platform for DMA Assessment"
  }
   If ($DMA_Opt -eq 'SKUGenReport')
  {
  $ComboList_Items = @("Azuresqldatabase", "AzureSqlManagedInstance" ,"AzureSqlVirtualMachine" , "ANY")
  $DMA_Opt=$null
  $textBoxIn.Visible = $false
  $titleLabel.Text = "Select Target platform for SKU Assessment"
  }
 
 
   If ($DMA_Opt -eq 'SKUAssessment')
  {
  #$get_serverlist = $Dir_Path+"Inventory"+'\'+"MS_Instancelist.txt"
  $ComboList_Items = @("ALL","ANY")
  $DMA_Opt=$null
  $titleLabel.Text = "Select SKU Assessment option"
  #$textBoxIn.Visible = $false
  #$mainFormSI.Controls.Add($ResetButtonsku)
  }
 

  #AzureSqlVirtualMachine
  #Loop thru the text file or the array
  #and add the contents to the combobox for selection
  ForEach ($Server in $ComboList_Items) {

    $the_combo.Items.Add($Server)


  }

  $mainFormSI.controls.add($the_combo)

  #action that will capture every time a value is selected on the combobox
  $the_combo_SelectedIndexChanged=
  {

    $targetcheck= $the_combo.text
    write-host "$targetcheck is selected Item from List" -ForegroundColor Yellow

  }

  $the_combo.add_SelectedIndexChanged($the_combo_SelectedIndexChanged)
 
  # Process Button
  $buttonProcess = New-Object System.Windows.Forms.Button
  $buttonProcess.Location = "35,150"
  $buttonProcess.Size = "110, 35"
  $buttonProcess.ForeColor = "Red"
  $buttonProcess.BackColor = "White"
  $buttonProcess.Text = "Process"
  #$buttonProcess.add_Click{processsingleServer}
  $buttonProcess.add_Click{Progresswindowgui -Target $DMAReports }
  $mainFormSI.Controls.Add($buttonProcess)

   $ResetButtonsku = New-Object System.Windows.Forms.Button
  $ResetButtonsku.Location = "250,150"
  $ResetButtonsku.Size = "100,30"
  $ResetButtonsku.ForeColor = "Red"
  $ResetButtonsku.BackColor = "White"
  $ResetButtonsku.Text = "Reset"
  $ResetButtonsku.add_Click{Resetsku}
  #$mainFormSI.Controls.Add($ResetButtonsku)
 
  # Exit Button 
  $exitButton = New-Object System.Windows.Forms.Button
  $exitButton.Location = "450,150"
  $exitButton.Size = "95,30"
  $exitButton.ForeColor = "Red"
  $exitButton.BackColor = "White"
  $exitButton.Text = "Exit"
  $exitButton.add_Click{$mainFormSI.close()}
  $mainFormSI.Controls.Add($exitButton)
  #[void]$mainFormSI.ShowDialog()
  [void]$mainFormSI.ShowDialog()

  
}

Function Progresswindowgui
{

  param (
        [Parameter(Mandatory=$true)][string]$Target
    )

    Add-Type -AssemblyName System.Drawing
  Add-Type -AssemblyName System.Windows.Forms
  $main_form            = New-Object System.Windows.Forms.Form
  $main_form.Text           ='Reports Progressbar'
  $main_form.foreColor      ='white'
  $main_form.BackColor      ='Darkblue'
  $main_form.Font           = $header
  $main_form.Width          = 600
  $main_form.Height         = 250

  $header                   = New-Object System.Drawing.Font("Verdana",13,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
  $procFont                 = New-Object System.Drawing.Font("Verdana",20,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

  $Label                    = New-Object System.Windows.Forms.Label
  $Label.Font               = $header
  $Label.ForeColor          ='yellow'
  $Label.Text               = "Are you sure want to continue......"
  $Label.Location           = New-Object System.Drawing.Point(10,10)
  $Label.Width              = 480
  $Label.Height             = 80

  $StartButton              = New-Object System.Windows.Forms.Button
  $StartButton.Location     = New-Object System.Drawing.Size(350,75)
  $StartButton.Size         = New-Object System.Drawing.Size(120,50)
  $StartButton.Text         = "Start"
  $StartButton.height       = 40
  $StartButton.BackColor    ='white'
  $StartButton.ForeColor    ='red'
  $StartButton.Add_click({Progressbar -Target $Target})

  $EndButton              = New-Object System.Windows.Forms.Button
  $EndButton.Location     = New-Object System.Drawing.Size(350,75)
  $EndButton.Size         = New-Object System.Drawing.Size(120,50)
  $EndButton.Text         = "OK"
  $EndButton.height       = 40
  $EndButton.BackColor    ='white'
  $EndButton.ForeColor    ='blue'
  #$EndButton.add_Click{$main_form.close()}
  
  $EndButton.DialogResult = [System.Windows.Forms.DialogResult]::OK

  $main_form.Controls.AddRange(($Label,$StartButton,$EndButton))

  $main_form.StartPosition = "manual"
  $main_form.Location = New-Object System.Drawing.Size(500, 300)
  $result=$main_form.ShowDialog() 
  $Target=$null


}

Function Progressbar
{

  param (
        [Parameter(Mandatory=$true)][string]$Target
    )
  Add-Type -AssemblyName System.Windows.Forms
  Add-Type -AssemblyName PresentationFramework
  Add-Type -AssemblyName System.Drawing
  
  [System.Windows.Forms.Application]::EnableVisualStyles()
  $ProgressBar              = New-Object System.Windows.Forms.ProgressBar
  $ProgressBar.Location     = New-Object System.Drawing.Point(10,35)
  $ProgressBar.Size         = New-Object System.Drawing.Size(460,40)
  $ProgressBar.Style        = "Marquee"
  $ProgressBar.MarqueeAnimationSpeed = 20
  $main_form.Controls.Add($ProgressBar)

  $Label.Font             = $header
  $Label.ForeColor        ='yellow'
  $Label.Text             ="Processing ..."
  #$ProgressBar.visible
  If ($Target -eq 'DMA_ForAllPaaS')
  {
    $ProgressBar.visible
    #AzureDBDMA 
    DMA_ForAllPaaS
    $Label.Text               = "Process Complete"
    $ProgressBar.Hide()
    $StartButton.Hide()
    $EndButton.Visible
  }
  If ($Target -eq 'SKUReportitem')
  {
    $ProgressBar.visible
    Generate_SKURecommendationReport 
    $Label.Text               = "Process Complete"
    $ProgressBar.Hide()
    $StartButton.Hide()
    $EndButton.Visible
  }

  #SKUAssfunction

  If ($Target -eq 'SKUAssfunction')
  {
    $ProgressBar.visible
    SKUAssessment 
    $Label.Text               = "Process Complete"
    $ProgressBar.Hide()
    $StartButton.Hide()
    $EndButton.Visible
  }
  If ($Target -eq 'DMA-Single-Instance') 
  {
    $ProgressBar.visible
    #DMA-Single-Instance 
    DMA_ForAllPaaS -textboxitem 'Yes'
    $Label.Text               = "Process Complete"
    $ProgressBar.Hide()
    $StartButton.Hide()
    $EndButton.Visible
  }
 
  If ($Target -eq 'DMA_All_IaaS') 
  {
    $ProgressBar.visible
    #DMA-Single-Instance 
    DMA_ForAllPaaS
    $Label.Text               = "Process Complete"
    $ProgressBar.Hide()
    $StartButton.Hide()
    $EndButton.Visible
  }

   
}

 

Function both_azmi
{
    Param(
        [Parameter(Mandatory=$true,Position=0,HelpMessage="Please specify SQL Instance Name")]   [string]$InstanceName,  
        [Parameter(Mandatory=$true ,Position=1,HelpMessage="Please specify list item")] [string]$inputvarListitem
        #[Parameter(Mandatory=$false ,Position=2,HelpMessage="Please specify text box input")] [string]$Textboxitem
        )    
  
    write-host "This both azmi function starting $InstanceName and $inputvarListitem " -ForegroundColor Magenta
    $SourcePlatform = "SqlOnPrem"
     

     
   
    #$Inputvar = $the_combo.text 
 

    if($inputvarListitem -eq "AzureSqlDatabase")
    {
        $Target = "AzureSqlDatabase"
    }
    elseif($inputvarListitem -eq "ManagedSqlServer")
    {
        $Target = "ManagedSqlServer"
    }
    
    elseif($inputvarListitem -eq "Both")
    {
        $Target = "AzureSqlDatabase,ManagedSqlServer"
    }
    else
    {
        #write-host "Wrong Choice"
        #Write-Host "■ [WARNING] " -ForegroundColor Yellow -NoNewline
        #return
        $Target = $inputvarListitem
    }
   
    $Targetlist=$Target.split(',');

        $FolderPath = "C:\Axalta_Assessment_Report"
        foreach($TargetPlatform in $Targetlist)
        {
                $sqlConnectionString = "Initial Catalog=master;Integrated Security=true;Connection Timeout=20";
                                
                $sqlstr = $sqlConnectionString.replace("Connection Timeout=20","");
                $sqlstr = $sqlstr.replace("Initial Catalog=master;","");
                $FileName =$InstanceName.Replace("\","-");
                $FileName =$FileName.Replace(".","")+"_"+$TargetPlatform;
                $sqlstr =  "Server=$InstanceName;"+$sqlstr;
                $app = "C:\Program Files\Microsoft Data Migration Assistant\DmaCmd.exe"

                if($TargetPlatform -eq "AzureSqlDatabase" -or $TargetPlatform -eq "ManagedSqlServer")
                {
                    $arg1 = ' /AssessmentName='+$FileName+' /AssessmentSourcePlatform='+$SourcePlatform+' /AssessmentDatabases="'+$sqlstr+'" /AssessmentTargetPlatform='+$TargetPlatform+' /AssessmentEvaluateCompatibilityIssues /AssessmentEvaluateFeatureParity  /AssessmentOverwriteResult /AssessmentResultDma="'+"$FolderPath\DMA\DMA_Format\"+"$FileName.dma"+'" /AssessmentResultJson="'+"$FolderPath\DMA\Json_Format\"+"$FileName.json"+'" /AssessmentResultCsv="'+"$FolderPath\DMA\CSV_Format\"+"$FileName.csv"+'"'
                }
                else
                {
                    $arg1 = ' /AssessmentName='+$FileName+' /AssessmentDatabases="'+$sqlstr+'" /AssessmentTargetPlatform='+$TargetPlatform+' /AssessmentEvaluateCompatibilityIssues  /AssessmentOverwriteResult /AssessmentResultDma="'+"$FolderPath\DMA\DMA_Format\"+"$FileName.dma"+'" /AssessmentResultJson="'+"$FolderPath\DMA\Json_Format\"+"$FileName.json"+'" /AssessmentResultCsv="'+"$FolderPath\DMA\CSV_Format\"+"$FileName.csv"+'"'
                }

                Write-host "$FileName  -> Assessment is running against source - $SourcePlatform and target- $TargetPlatform."-ForegroundColor Yellow
               #$arg1
                start-process -FilePath $app -ArgumentList $arg1  -Wait
            }
            Write-Host "All captured data will be saved at $FolderPath\DMA"
            }



 Function DMA_ForAllPaaS
    {
      param
          (
            [Parameter(Mandatory=$false ,Position=2,HelpMessage="Please specify text box input for Single Server PaaS")] [string]$Textboxitem
          )
    [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMO')
    $sucess=$Dir_Path+"Inventory"+'\'+"server_connect.txt"
    $failure=$Dir_Path+"Inventory"+'\'+"server_not_connect.txt" #$DMA_Path+"server_not_connect.txt"
    $srvoutput = $Dir_Path+"Inventory"+'\'+"MS_Instancelist.txt"
    $Inputvar = $the_combo.text 
    if ($Textboxitem -eq "")
       {
    $getserverlist = $Dir_Path+"Inventory"+'\'+"serverlist.txt"
    $ServerNames = get-content $getserverlist
    #$DatabaseName="master" $Dir_Path"Inventory"'\'"serverlist.txt"
       #}
    
    $s=$null
    #$Inputvar = $the_combo.text    
    Foreach ($s in $ServerNames)

    {
      TRY
      {
        $s
        write-host "Host Server name is : $s" -ForegroundColor Yellow
       
       
        $inst=$null
        $inst_name = (Get-Service -Name MSSQLSERVER,MSSQL$* -ComputerName $s | Where-Object {($_.Status -eq "Running")} -ErrorAction SilentlyContinue)
        $inst = ($inst_name).name -replace 'MSSQL\$'

        foreach ($i in $inst)
        
        {
        
                 
          if($i -eq 'MSSQLSERVER')
          {
         
            #####################
            write-host "Default instance loop" -ForegroundColor Yellow

            $servername = New-Object Microsoft.SqlServer.Management.Smo.Server -ArgumentList $s
        
            $srv=$servername.Name
            $srv
            $name = $s | Out-File -FilePath $srvoutput -Append
            write-host "Default Instance Server name is : $srv" -ForegroundColor Yellow
            #$servername
  
            #if (Test-Connection -Delay 15 -ComputerName $servername -Count 1 -ErrorAction SilentlyContinue -quiet){
       
            Write-Host "$srv,up" -ForegroundColor Green
            "$srv,up" | Out-file $sucess -Append
            #################Calling function#############
            both_azmi -instancename $srv -inputvarListitem $Inputvar
            ##########################################
    
     }
          else{        
            #$name = $s+'\'+$i | Out-File -FilePath C:\DMA_Reports_Final_named\MS_Instancelist.txt -Append 
            $srv1= $s+'\'+$i 
            "Named Instance  " +$srv1

            ##################
            $servername1=$null   
            $servername1 = New-Object Microsoft.SqlServer.Management.Smo.Server -ArgumentList $srv1
            #$servername
            $snamed=$servername1.Name
            $name = $s+'\'+$i | Out-File -FilePath $srvoutput -Append
            "Named instance 2 " +$snamed
            #$servername1.Databases | select name
            #if (Test-Connection -Delay 15 -ComputerName $servername -Count 1 -ErrorAction SilentlyContinue -quiet)
            #if ($servername.Databases.Contains($DatabaseName))
            #{
       
       
            Write-Host "$srv1,up" -ForegroundColor Green
            "$srv1,up" | Out-file $sucess -Append
            write-host "Nameed instance loop" -ForegroundColor Yellow
              #$d=$null
              #$d=$srv1.replace('\','_')
              #################Calling function#############
            both_azmi -instancename $srv1 -inputvarListitem $Inputvar
            ##########################################
               }                          
                                       
         } 
      }
      catch {
              Write-Host "$s Server is not connecting" -ForegroundColor Red
               Write-Host "$S,down" -ForegroundColor Red
               "$S,down" | Out-file $failure -Append
        }  

        } # end foreach
     } # end of if
        else 
         {
            TRY 
            {
            $ServerNames = $textBoxIn.text
            SQL_SERVER_Validation_singlePaaS -s $ServerNames
                 
                 Write-Host "$ServerNames,up" -ForegroundColor green
                "$ServerNames,up" | Out-file $sucess -Append
                ##############Call Function######################
                both_azmi -instancename $ServerNames -inputvarListitem $Inputvar
                #################################################           
            }
             catch {  
                              

                write-host "Server $ServerNames is not connected using validation function" -ForegroundColor red
                 Write-Host "$ServerNames,down" -ForegroundColor green
              "$ServerNames,down" | Out-file $failure -Append
               [void] [System.Windows.MessageBox]::Show( "$ServerNames unable to connect,Please check  ", "Script completed", "OK", "Information" )
                       ########################################## 
                                  }
         }

        #} 
        }
                 

Function SQL_SERVER_Validation_singlePaaS
{
  param (
        [Parameter(Mandatory=$true)][string]$s
    )
    #Add-Type -AssemblyName Microsoft.SqlServer.Smo
   [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMO')
  Add-Type -AssemblyName PresentationFramework
  Import-Module SQLSERVER -DisableNameChecking
  $connect=$DMA_Path+"server_connect.txt"
  $notconnect=$DMA_Path+"server_not_connect.txt"
  $srvoutput=$DMA_Path+"MS_Instancelist.txt"

    
  #TRY
  #{

          $servername = New-Object ('Microsoft.SqlServer.Management.Smo.Server') -ArgumentList $s
        
          $srv=$servername.Name
          $dbname=$servername.Databases
          $srv
          $dbname
          #$name = $s | Out-File -FilePath $srvoutput -Append

          #Write-Host "$srv,up" -ForegroundColor Green
          #"$srv,up" | Out-file $connect -Append
          #[void] [System.Windows.MessageBox]::Show( "Server $s is up and running ", "Script completed", "OK", "Information" ) 
          ########################
       
    #} #end TRY
        
    #catch
    #{
     #Write-Host "$s Server is not connecting" -ForegroundColor Red
              # Write-Host "$S,down" -ForegroundColor green
              # "$S,down" | Out-file $notconnect -Append
              # [void] [System.Windows.MessageBox]::Show( "Server $s is down ", "Script completed", "OK", "Information" )
    #}
    
 
  } # end function

  Function SKUAssessment
{
   Param(
        [Parameter(Mandatory=$false,Position=0,HelpMessage="Please Ouput Folder Directory")] [string]$InputOutputFolder
        #[Parameter(Mandatory=$True,Position=0,HelpMessage="Please specify Azure ype")] [string]$inputvarsku
        ) 

     $ScriptPath = "C:\Axalta_Assessment_Report"   
         
    if($InputOutputFolder -eq "")
    {
       $InputOutputFolder = "$ScriptPath\DMA\sku_counters"
    } 

    $skuassitem = $the_combo.Text

    if ($skuassitem -eq 'ALL')
    {  
      $textBoxIn.Visible = $false
      $getserverlist = $Dir_Path+"Inventory"+'\'+"MS_Instancelist.txt"
      $getserver = get-content $getserverlist
     

      foreach($InstanceName in $getserver)
             {
                $sqlConnectionString = "Initial Catalog=master;Integrated Security=true;Connection Timeout=20";
                                
                $sqlstr = $sqlConnectionString.replace("Connection Timeout=20","");
                $sqlstr = $sqlstr.replace("Initial Catalog=master;","");
                $FileName =$InstanceName.Replace("\","_");
                $FileName =$FileName.Replace(".","");
                $pathsku = new-item -type directory -path $InputOutputFolder"\"$FileName -Force 
                #$path
                $path = $InputOutputFolder+'\'+$FileName
                #$path
                $sqlstr =  "Server=$InstanceName;"+$sqlstr;
                $app = "C:\Program Files\Microsoft Data Migration Assistant\SqlAssessmentConsole\SqlAssessment.exe"
                $arg1 = ' PerfDataCollection --sqlConnectionStrings "'+$sqlstr+'" --outputFolder "'+$path+'"' 
               #Write-host $app$arg1; 
                start-process -FilePath $app -ArgumentList $arg1 
            }
    }
     elseif ($skuassitem -eq 'ANY')
     {
         #$textBoxIn.Visible = $true
         $the_combo.Enabled = $false
         $InstanceName = $textBoxIn.Text

         $sqlConnectionString = "Initial Catalog=master;Integrated Security=true;Connection Timeout=20";
                                
                $sqlstr = $sqlConnectionString.replace("Connection Timeout=20","");
                $sqlstr = $sqlstr.replace("Initial Catalog=master;","");
                $FileName =$InstanceName.Replace("\","_");
                $FileName =$FileName.Replace(".","");
                $pathsku = new-item -type directory -path $InputOutputFolder"\"$FileName -Force 
                #$path
                $path = $InputOutputFolder+'\'+$FileName
                #$path
                $sqlstr =  "Server=$InstanceName;"+$sqlstr;
                $app = "C:\Program Files\Microsoft Data Migration Assistant\SqlAssessmentConsole\SqlAssessment.exe"
                $arg1 = ' PerfDataCollection --sqlConnectionStrings "'+$sqlstr+'" --outputFolder "'+$path+'"' 
               #Write-host $app$arg1; 
                start-process -FilePath $app -ArgumentList $arg1 
     }

     
     #$TargetPlatform = 'AzureSqlDatabase'
       
             
            Write-Host "During data collection, press Enter to stop data collection.All captured data will be saved at $InputOutputFolder."
           # }
            } # end function


 Function Generate_SKURecommendationReport
{
   Param(
        [Parameter(Mandatory=$false,Position=0,HelpMessage="Please Specify Input/Ouput Folder Directory")] [string]$InputOutputFolder   
        #[Parameter(Mandatory=$True,Position=0,HelpMessage="Please specify Azure ype")] [string]$inputvarsku
        )      
    if($InputOutputFolder -eq "")
    {
        #$InputOutputFolder = $path#"$ScriptPath\DMA\sku_counters"
        $ScriptPath = "C:\Axalta_Assessment_Report" 
        $InputOutputFolder = "$ScriptPath\DMA\sku_counters"
    }

      $date = Get-date -UFormat -%Y%m%d

    $renamejson = "SkuRecommendationReport"+$date+".json"
    $renamehtml = "SkuRecommendationReport"+$date+".html"

      $inputvarsku = $the_combo.text

    if($inputvarsku -eq "AzureSqlDatabase")
    {
        $TargetPlatform = "AzureSqlDatabase"
    }
    elseif($inputvarsku -eq "AzureSqlManagedInstance")
    {
        $TargetPlatform = "AzureSqlManagedInstance"
    }
    
    elseif($inputvarsku -eq "AzureSqlVirtualMachine")
    {
        $TargetPlatform = "AzureSqlVirtualMachine"
    }

    elseif($inputvarsku -eq "ANY")
    {
         $TargetPlatform = "ANY"
    }

        

     $getserverlist = $Dir_Path+"Inventory"+'\'+"MS_Instancelist.txt"
     $getserver = get-content $getserverlist
     #Start-Sleep 4
         
             foreach($InstanceName in $getserver)
             {
    $FileName =$InstanceName.Replace("\","_");
    $FileName =$FileName.Replace(".","");
    $path = $InputOutputFolder+'\'+$FileName
    $app = "C:\Program Files\Microsoft Data Migration Assistant\SqlAssessmentConsole\SqlAssessment.exe"
    $arg1 = ' GetSkuRecommendation --outputFolder "'+$path+'" --targetPlatform "'+$targetPlatform+'"'
    #Write-host $app$arg1; 
    start-process -FilePath $app -ArgumentList $arg1 -wait  
    Write-Host "HTML Recommendation files will be stored at $InputOutputFolder" 
    Start-Sleep 3
    Rename-Item -Path "$path\$renamehtml" -NewName "SkuRecommendationReport_$targetPlatform$date.html"
    Rename-Item -Path "$path\$renamejson" -NewName "SkuRecommendationReport_$targetPlatform$date.json"
    #$mainFormSI.Controls.Add($ResetButtonsku)

             }
      
}

Function DMA_files_view
 {
  param (
        [Parameter(Mandatory=$false,Position=0,HelpMessage="Please specify Report type")] [string]$Filetype
        )
  Add-Type -AssemblyName System.Windows.Forms
  $onclick_buttoncombo =
  {
    #SQL_SERVER_Validation_single -s $textboxInaz.text
    $TargetPlatform2=$null
    $TargetPlatform3=$null
    #if ($the_comboaz.SelectedIndex -eq 0 -or 1)
    #{
      #SQL_SERVER_Validation_single -s $textboxInaz.text
      $TargetPlatform2=$the_comboaz.Text
      #$TargetPlatform2
      $TargetPlatform3 = $the_comboazvm.Text
      $TargetPlatform3
      $i=$TargetPlatform2
      $d=$i.replace('\','-')
      #$comboitem=$TargetPlatform2
      #$comboitem
      If ($Filetype -eq 'DMAfile')
      {
      $new=$DMA_Path+"DMA\DMA_format\"+$d+"_"+$TargetPlatform3+".dma"
      write-host "path is : $new" -ForegroundColor Yellow
      #$url=$new +"\"+ $d+"_"+$TargetPlatform2+".dma"
      #$url=$new +"\"+ $d+"_"+$TargetPlatform3+".dma"
      #write-host "final path is  :--$url"
      #Invoke-Expression $url
      Invoke-Expression $new
      
      }
      

      elseif ($Filetype -eq 'sku_counters')
      #else
      {
       
             #$TargetPlatform4 = 'AzureSqlManagedInstance'
             $new=$DMA_Path+"DMA\"+$Filetype+"\"+$d
             $date = Get-date -UFormat -%Y%m%d
             write-host "path is : $new" -ForegroundColor Yellow
             $s= "SkuRecommendationReport"
      #$url=$new +"\"+ $d+"_"+$TargetPlatform2+".dma"
      #$url=$new +"\"+"*.HTML"
            $url=$new +"\"+ $s+"_"+$TargetPlatform3+$date+".HTML"
            write-host "final path is  :--$url"
            Invoke-Expression $url
          
                
      }
   
    

  }

  
  
  $mainFormHTML = New-Object System.Windows.Forms.Form
  $mainFormHTML.Font = $header#"Comic Sans MS,8.25"
  $mainFormHTML.Text = " DMA Reports Files"
  $mainFormHTML.FormBorderStyle = "FixedDialog"
  $mainFormHTML.ForeColor = "White"
  $mainFormHTML.BackColor = "DarkBlue"
  $mainFormHTML.StartPosition = "CenterParent"
  $mainFormHTML.width = 600
  $mainFormHTML.height = 250
 
  # Title Label
  $titleLabel = New-Object System.Windows.Forms.Label
  $titleLabel.Font = $header#"Comic Sans MS,14"
  $titleLabel.ForeColor = "Yellow"
  $titleLabel.Location = "30,20"
  $titleLabel.Size = "400,30"
  $titleLabel.Text = "Select Instance name"
  $mainFormHTML.Controls.Add($titleLabel);
  #$mainFormHTML.Controls.Add($titleLabel)

  # Input Box
  $textBoxInaz = New-Object System.Windows.Forms.TextBox
  $textBoxInaz.Location = "35, 70"
  $textBoxInaz.Size = "500, 20"
  $textBoxInaz.Text = ""
  #$mainFormHTML.Controls.Add($textBoxInaz)
 
  #Combobox
  $the_comboaz = New-Object system.Windows.Forms.ComboBox
  $the_comboaz.location = "35, 60"
  $the_comboaz.Font = $header
  $the_comboaz.Size = "250, 20"

  $the_comboaz.DropDownStyle = "Dropdownlist"
  #$ComboList_Items = Get-Content $DMA_Path"DMA_DB_Type.txt"
  $get_serverlist = $Dir_Path+"Inventory"+'\'+"MS_Instancelist.txt"
  $ComboList_Items = Get-Content $get_serverlist #@("Azuresqldatabase", "ManagedSqlServer" ,"AzureSqlVirtualMachine")

  #Loop thru the text file or the array
  #and add the contents to the combobox for selection
  ForEach ($Server in $ComboList_Items) {

    $the_comboaz.Items.Add($Server)


  }

  $mainFormHTML.controls.add($the_comboaz)
  # New combobox 
  $the_comboazvm = New-Object system.Windows.Forms.ComboBox
  $the_comboazvm.location = "35, 100"
  $the_comboazvm.Font = $header
  $the_comboazvm.Size = "250, 20"

  $the_comboazvm.DropDownStyle = "Dropdownlist"
  #$ComboList_Items = Get-Content $DMA_Path"DMA_DB_Type.txt"
  $ComboList_Items2 = @("Azuresqldatabase", "AzureSqlManagedInstance" ,"AzureSqlVirtualMachine")# @("Sqlserver2012", "Sqlserver2014" ,"Sqlserver2016","SqlServerWindows2017","SqlServerLinux2017","SqlServerWindows2019","SqlServerLinux2019")

  #Loop thru the text file or the array
  #and add the contents to the combobox for selection
  ForEach ($Server2 in $ComboList_Items2) {

    $the_comboazvm.Items.Add($Server2)
    
    }

    $event_handler = 
  {
    #$the_combo.Items.Clear()
    $targetcheck=$the_comboaz.Text
    if($targetcheck -ne "") #AzureSqlVirtualMachine
    {
      $the_comboaz.Enabled=$false
      $mainFormHTML.controls.add($the_comboazvm)
    }
    else
    {
      $the_comboaz.Enabled=$True
     
    }
  }
   
    $the_comboaz.add_SelectedIndexChanged($event_handler)

    #$the_comboaz.SelectedIndex=0

    # Process Button
  $buttonProcess = New-Object System.Windows.Forms.Button
  $buttonProcess.Location = "35,170"
  $buttonProcess.Size = "100, 25"
  $buttonProcess.ForeColor = "Red"
  $buttonProcess.BackColor = "White"
  $buttonProcess.Text = "View"
  $buttonProcess.add_Click($onclick_buttoncombo)
  $mainFormHTML.Controls.Add($buttonProcess)
 
  # Exit Button 
  $exitButton = New-Object System.Windows.Forms.Button
  $exitButton.Location = "450,170"
  $exitButton.Size = "100,25"
  $exitButton.ForeColor = "Red"
  $exitButton.BackColor = "White"
  $exitButton.Text = "Exit"
  $exitButton.add_Click{$mainFormHTML.close()}
  $mainFormHTML.Controls.Add($exitButton)

  # Reset Button 
  $ResetButton = New-Object System.Windows.Forms.Button
  $ResetButton.Location = "200,170"
  $ResetButton.Size = "100,25"
  $ResetButton.ForeColor = "Red"
  $ResetButton.BackColor = "White"
  $ResetButton.Text = "Reset"
  $ResetButton.add_Click{Reset_button}
  $mainFormHTML.Controls.Add($ResetButton)

  # Audio Help Button 
  $AhelpButton = New-Object System.Windows.Forms.Button
  $AhelpButton.Location = "300,170"
  $AhelpButton.Size = "100,25"
  $AhelpButton.ForeColor = "Red"
  $AhelpButton.BackColor = "White"
  $AhelpButton.Text = "Help"
  $AhelpButton.add_Click{Audio_help}
  $mainFormHTML.Controls.Add($AhelpButton)
  

    [void]$mainFormHTML.ShowDialog()
 
}

Function Audio_help
 {   
  Add-Type -AssemblyName System.Speech
$synthesizer = New-Object -TypeName System.Speech.Synthesis.SpeechSynthesizer
$synthesizer.SelectVoice("Microsoft Zira Desktop")
$synthesizer.Rate = -2
$synthesizer.Speak('Hey, welcome to DMA Report file window')
$synthesizer.Speak('In First Combo box you can select any Instance Name')
$synthesizer.Speak('After that second combo box visible then only you can select Report type and click on view button')
}

Function Reset_button
{
  #$the_comboaz.Enabled=$True
  #$the_comboazvm.Visible=$False
  #$the_comboaz.Text=""

  $mainFormHTML.Controls.Remove($the_comboaz)
  $mainFormHTML.Controls.Remove($the_comboazvm)
  $mainFormHTML.Controls.add($the_comboaz)
  $the_comboaz.Enabled=$True
  #$the_comboazvm.Controls.add($the_comboazvm)
  #$the_comboaz.Text=""
  #$the_comboaz.Dispose()
  #$the_comboaz.Items.Clear()
  #$the_comboaz.Enabled=$True
  #$textBoxInaz.Text =""
  #$the_comboaz.Items.Clear()#$the_combo.Enabled = $false

}

Function Resetsku
{
  #$the_comboaz.Enabled=$True
  #$the_comboazvm.Visible=$False
  #$the_comboaz.Text=""
  $the_combo.Enabled = $true
  $textBoxIn.Text = ""
  #$mainFormHTML.Controls.Remove($the_comboaz)
  #$mainFormHTML.Controls.Remove($the_comboazvm)
  #$mainFormHTML.Controls.add($the_comboaz)
  #$the_comboaz.Enabled=$True
  #$the_comboazvm.Controls.add($the_comboazvm)
  #$the_comboaz.Text=""
  #$the_comboaz.Dispose()
  #$the_comboaz.Items.Clear()
  #$the_comboaz.Enabled=$True
  #$textBoxInaz.Text =""
  #$the_comboaz.Items.Clear()#$the_combo.Enabled = $false

}

Function Inventory
{
  param (
        [Parameter(Mandatory=$false,Position=0,HelpMessage="Please specify Inventory type")] [string]$itype
        )

        if ($itype -eq 'Project')
        {
        #$f = "inventory_list.txt"
        #$new=$DMA_Path+"DMA\"+"Inventory\"+$f
         $documentPath = "C:\Axalta_Assessment_Report\inventory\inventory_list.txt"
        notepad.exe $documentPath
        }
        elseif($itype -eq 'Tool')
        {
        #$f = "serverlist.txt"
        #$new=$DMA_Path+"DMA\"+"Inventory\"+$f
        #$new
         $documentPath = "C:\Axalta_Assessment_Report\inventory\serverlist.txt"
        notepad.exe $documentPath
        }
        elseif($itype -eq 'success')
        {
        
         $documentPath = "C:\Axalta_Assessment_Report\inventory\server_connect.txt"
        notepad.exe $documentPath
        }
        elseif($itype -eq 'Failure')
        {
      
         $documentPath = "C:\Axalta_Assessment_Report\inventory\server_not_connect.txt"
        notepad.exe $documentPath
        }
         elseif($itype -eq 'SKU_Instances')
        {
        
         $documentPath = "C:\Axalta_Assessment_Report\inventory\MS_Instancelist.txt"
        notepad.exe $documentPath
        }
}

function Import_JSON_to_SQLDB
{
    Param(
            [Parameter(Mandatory=$false,Position=0,HelpMessage="Please Input InstanceName")] [string]$InstanceName 
        )      
     
    if($InstanceNamer -eq "")
    {
       Write-Host "InstanceName is missing.Please try again"
       return
    } 

    $InstanceName = "sqlmydb001\SQL2023"#$env:COMPUTERNAME
    $InstanceName
    $Dir_Path ="C:\Axalta_Assessment_Report\"

    $DMA_Path =$Dir_Path

    $InputOutputFolder = "$DMA_Path"+"DMA\Json_Format\"

    if($InstanceNamer -ne "")
    {
        dmaProcessor -serverName $InstanceName -databaseName DMAReportingAxalta -jsonDirectory $InputOutputFolder  -processTo SQLServer
            
            

            [void] [System.Windows.MessageBox]::Show( "JSON file Imported Sucessfully  ", "Script completed", "OK", "Information" )
    }
} 


#Import JSON to SQL on prem or azure
function dmaProcessor 
{
param(
    [parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string] $serverName,

    [parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string] $databaseName,

    [parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string] $jsonDirectory,

    [parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [ValidateSet("SQLServer")] 
    [string] $processTo
)
   #dmaProcessor -serverName 'sqlmydb001' -databaseName 'DMA_Reporting' -jsonDirectory 'C:\myDB_Assessment_Report\DMA'
    #Create database objects
    [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.SMO") | Out-Null
    $srv = New-Object Microsoft.SqlServer.Management.SMO.Server($serverName)
           
    #create reporting database
    $dbCheck = $srv.Databases | Where {$_.Name -eq "$databaseName"} | Select Name
    if(!$dbCheck)
    {            
        $db = New-Object Microsoft.SqlServer.Management.Smo.Database ($srv, $databaseName)

        $db.Create()

        Write-Host("Database $databaseName created successfully") -ForegroundColor Green
    }
    else
    {
            $db=$srv.Databases.Item($databaseName)
            Write-Host ("Database $databaseName already exists") -ForegroundColor Yellow
    }

    #create ReportData table
    $tableCheck = $db.Tables | Where {$_.Name -eq "ReportData"}
    if(!$tableCheck)
    {            
        $ReportDatatbl = New-Object Microsoft.SqlServer.Management.Smo.Table($db, "ReportData")

        $col1 = New-Object Microsoft.SqlServer.Management.Smo.Column($ReportDatatbl, "ImportDate", [Microsoft.SqlServer.Management.Smo.DataType]::DateTime)
        $col2 = New-Object Microsoft.SqlServer.Management.Smo.Column($ReportDatatbl, "InstanceName", [Microsoft.SqlServer.Management.Smo.DataType]::VarChar(50))
        $col3 = New-Object Microsoft.SqlServer.Management.Smo.Column($ReportDatatbl, "Status", [Microsoft.SqlServer.Management.Smo.DataType]::VarChar(30))
        $col4 = New-Object Microsoft.SqlServer.Management.Smo.Column($ReportDatatbl, "Name", [Microsoft.SqlServer.Management.Smo.DataType]::VarChar(255))
        $col5 = New-Object Microsoft.SqlServer.Management.Smo.Column($ReportDatatbl, "SizeMB", [Microsoft.SqlServer.Management.Smo.DataType]::Decimal(6, 38))
        $col6 = New-Object Microsoft.SqlServer.Management.Smo.Column($ReportDatatbl, "SourceCompatibilityLevel", [Microsoft.SqlServer.Management.Smo.DataType]::VarChar(15))
        $col7 = New-Object Microsoft.SqlServer.Management.Smo.Column($ReportDatatbl, "TargetCompatibilityLevel", [Microsoft.SqlServer.Management.Smo.DataType]::VarChar(15))
        $col8 = New-Object Microsoft.SqlServer.Management.Smo.Column($ReportDatatbl, "Category", [Microsoft.SqlServer.Management.Smo.DataType]::VarChar(30))
        $col9 = New-Object Microsoft.SqlServer.Management.Smo.Column($ReportDatatbl, "Severity", [Microsoft.SqlServer.Management.Smo.DataType]::VarChar(15))
        $col10 = New-Object Microsoft.SqlServer.Management.Smo.Column($ReportDatatbl, "ChangeCategory", [Microsoft.SqlServer.Management.Smo.DataType]::VarChar(20))
        $col11 = New-Object Microsoft.SqlServer.Management.Smo.Column($ReportDatatbl, "RuleId", [Microsoft.SqlServer.Management.Smo.DataType]::VarChar(100))
        $col12 = New-Object Microsoft.SqlServer.Management.Smo.Column($ReportDatatbl, "Title", [Microsoft.SqlServer.Management.Smo.DataType]::VarCharMax)
        $col13 = New-Object Microsoft.SqlServer.Management.Smo.Column($ReportDatatbl, "Impact", [Microsoft.SqlServer.Management.Smo.DataType]::VarCharMax)
        $col14 = New-Object Microsoft.SqlServer.Management.Smo.Column($ReportDatatbl, "Recommendation", [Microsoft.SqlServer.Management.Smo.DataType]::VarCharMax)
        $col15 = New-Object Microsoft.SqlServer.Management.Smo.Column($ReportDatatbl, "MoreInfo", [Microsoft.SqlServer.Management.Smo.DataType]::VarCharMax)
        $col16 = New-Object Microsoft.SqlServer.Management.Smo.Column($ReportDatatbl, "ImpactedObjectName", [Microsoft.SqlServer.Management.Smo.DataType]::VarChar(255))
        $col17 = New-Object Microsoft.SqlServer.Management.Smo.Column($ReportDatatbl, "ImpactedObjectType", [Microsoft.SqlServer.Management.Smo.DataType]::VarChar(50))
        $col18 = New-Object Microsoft.SqlServer.Management.Smo.Column($ReportDatatbl, "ImpactDetail", [Microsoft.SqlServer.Management.Smo.DataType]::VarCharMax)
        $col19 = New-Object Microsoft.SqlServer.Management.Smo.Column($ReportDatatbl, "DBOwner", [Microsoft.SqlServer.Management.Smo.DataType]::VarChar(128))
        $col20 = New-Object Microsoft.SqlServer.Management.Smo.Column($ReportDatatbl, "AssessmentTarget", [Microsoft.SqlServer.Management.Smo.DataType]::VarChar(50))
        $col21 = New-Object Microsoft.SqlServer.Management.Smo.Column($ReportDatatbl, "AssessmentName", [Microsoft.SqlServer.Management.Smo.DataType]::VarChar(128))
              
        $ReportDatatbl.Columns.Add($col1)
        $ReportDatatbl.Columns.Add($col2)
        $ReportDatatbl.Columns.Add($col3)
        $ReportDatatbl.Columns.Add($col4)
        $ReportDatatbl.Columns.Add($col5)
        $ReportDatatbl.Columns.Add($col6)
        $ReportDatatbl.Columns.Add($col7)
        $ReportDatatbl.Columns.Add($col8)
        $ReportDatatbl.Columns.Add($col9)
        $ReportDatatbl.Columns.Add($col10)
        $ReportDatatbl.Columns.Add($col11)
        $ReportDatatbl.Columns.Add($col12)
        $ReportDatatbl.Columns.Add($col13)
        $ReportDatatbl.Columns.Add($col14)
        $ReportDatatbl.Columns.Add($col15)
        $ReportDatatbl.Columns.Add($col16)
        $ReportDatatbl.Columns.Add($col17)
        $ReportDatatbl.Columns.Add($col18) 
        $ReportDatatbl.Columns.Add($col19)
        $ReportDatatbl.Columns.Add($col20)
        $ReportDatatbl.Columns.Add($col21)    
            
        $ReportDatatbl.Create()
        Write-Host ("Table ReportData created successfully") -ForegroundColor Green
    }
    else
    {
        Write-Host ("Table ReportData already exists") -ForegroundColor Yellow
    }

    #create AzureFeatureParity table
    $tableCheck2 = $db.Tables | Where {$_.Name -eq "AzureFeatureParity"}
    if(!$tableCheck2)
    {            
        $AzureReportDatatbl = New-Object Microsoft.SqlServer.Management.Smo.Table($db, "AzureFeatureParity")

        $col1 = New-Object Microsoft.SqlServer.Management.Smo.Column($AzureReportDatatbl, "ImportDate", [Microsoft.SqlServer.Management.Smo.DataType]::DateTime)
        $col2 = New-Object Microsoft.SqlServer.Management.Smo.Column($AzureReportDatatbl, "ServerName", [Microsoft.SqlServer.Management.Smo.DataType]::VarChar(128))
        $col3 = New-Object Microsoft.SqlServer.Management.Smo.Column($AzureReportDatatbl, "Version", [Microsoft.SqlServer.Management.Smo.DataType]::VarChar(15))
        $col4 = New-Object Microsoft.SqlServer.Management.Smo.Column($AzureReportDatatbl, "Status", [Microsoft.SqlServer.Management.Smo.DataType]::VarChar(10))
        $col5 = New-Object Microsoft.SqlServer.Management.Smo.Column($AzureReportDatatbl, "Category", [Microsoft.SqlServer.Management.Smo.DataType]::VarChar(50))
        $col6 = New-Object Microsoft.SqlServer.Management.Smo.Column($AzureReportDatatbl, "Severity", [Microsoft.SqlServer.Management.Smo.DataType]::VarChar(50))
        $col7 = New-Object Microsoft.SqlServer.Management.Smo.Column($AzureReportDatatbl, "FeatureParityCategory", [Microsoft.SqlServer.Management.Smo.DataType]::VarChar(50))
        $col8 = New-Object Microsoft.SqlServer.Management.Smo.Column($AzureReportDatatbl, "RuleID", [Microsoft.SqlServer.Management.Smo.DataType]::VarChar(100))
        $col9 = New-Object Microsoft.SqlServer.Management.Smo.Column($AzureReportDatatbl, "Title", [Microsoft.SqlServer.Management.Smo.DataType]::VarChar(1000))
        $col10 = New-Object Microsoft.SqlServer.Management.Smo.Column($AzureReportDatatbl, "Impact", [Microsoft.SqlServer.Management.Smo.DataType]::VarChar(1000))
        $col11 = New-Object Microsoft.SqlServer.Management.Smo.Column($AzureReportDatatbl, "Recommendation", [Microsoft.SqlServer.Management.Smo.DataType]::VarCharMax)
        $col12 = New-Object Microsoft.SqlServer.Management.Smo.Column($AzureReportDatatbl, "MoreInfo", [Microsoft.SqlServer.Management.Smo.DataType]::VarCharMax)
        $col13 = New-Object Microsoft.SqlServer.Management.Smo.Column($AzureReportDatatbl, "ImpactedDatabasename", [Microsoft.SqlServer.Management.Smo.DataType]::VarChar(128))
        $col14 = New-Object Microsoft.SqlServer.Management.Smo.Column($AzureReportDatatbl, "ImpactedObjectType", [Microsoft.SqlServer.Management.Smo.DataType]::VarChar(30))
        $col15 = New-Object Microsoft.SqlServer.Management.Smo.Column($AzureReportDatatbl, "ImpactDetail", [Microsoft.SqlServer.Management.Smo.DataType]::VarCharMax)

        $AzureReportDatatbl.Columns.Add($col1)
        $AzureReportDatatbl.Columns.Add($col2)
        $AzureReportDatatbl.Columns.Add($col3)
        $AzureReportDatatbl.Columns.Add($col4)
        $AzureReportDatatbl.Columns.Add($col5)
        $AzureReportDatatbl.Columns.Add($col6)
        $AzureReportDatatbl.Columns.Add($col7)
        $AzureReportDatatbl.Columns.Add($col8)
        $AzureReportDatatbl.Columns.Add($col9)
        $AzureReportDatatbl.Columns.Add($col10)
        $AzureReportDatatbl.Columns.Add($col11)
        $AzureReportDatatbl.Columns.Add($col12)
        $AzureReportDatatbl.Columns.Add($col13)
        $AzureReportDatatbl.Columns.Add($col14)
        $AzureReportDatatbl.Columns.Add($col15)
            
        $AzureReportDatatbl.Create()
        Write-Host ("Table AzureFeatureParity created successfully") -ForegroundColor Green
    }
    else
    {
        Write-Host ("Table AzureFeatureParity already exists") -ForegroundColor Yellow
    }

    #create BreakingChangeWeighting table
    $tableCheck3 = $db.Tables | Where {$_.Name -eq "BreakingChangeWeighting"}
    if(!$tableCheck3)
    {            
        $BreakingChangetbl = New-Object Microsoft.SqlServer.Management.Smo.Table($db, "BreakingChangeWeighting")

        $col1 = New-Object Microsoft.SqlServer.Management.Smo.Column($BreakingChangetbl, "RuleId", [Microsoft.SqlServer.Management.Smo.DataType]::VarChar(36))
        $col1.Nullable = $false
        $col2 = New-Object Microsoft.SqlServer.Management.Smo.Column($BreakingChangetbl, "Title", [Microsoft.SqlServer.Management.Smo.DataType]::VarChar(150))
        $col3 = New-Object Microsoft.SqlServer.Management.Smo.Column($BreakingChangetbl, "Effort", [Microsoft.SqlServer.Management.Smo.DataType]::TinyInt)
        $col4 = New-Object Microsoft.SqlServer.Management.Smo.Column($BreakingChangetbl, "FixTime", [Microsoft.SqlServer.Management.Smo.DataType]::TinyInt)
        $col5 = New-Object Microsoft.SqlServer.Management.Smo.Column($BreakingChangetbl, "Cost", [Microsoft.SqlServer.Management.Smo.DataType]::TinyInt)
        $col6 = New-Object Microsoft.SqlServer.Management.Smo.Column($BreakingChangetbl, "ChangeRank", [Microsoft.SqlServer.Management.Smo.DataType]::TinyInt)
        $Col6.Computed = $True
        $Col6.ComputedText = "(Effort + FixTime + Cost) / 3"
       
        $BreakingChangetbl.Columns.Add($col1)
        $BreakingChangetbl.Columns.Add($col2)
        $BreakingChangetbl.Columns.Add($col3)
        $BreakingChangetbl.Columns.Add($col4)
        $BreakingChangetbl.Columns.Add($col5)
        $BreakingChangetbl.Columns.Add($col6)
        
        $BreakingChangetbl.Create()

        $PK = New-Object Microsoft.SqlServer.Management.Smo.Index($BreakingChangetbl,"PK_BreakingChangeWeighting_RuleId")
        $PK.IndexKeyType = "DriPrimaryKey"

        $IdxCol = New-Object Microsoft.SqlServer.Management.Smo.IndexedColumn($PK, $col1.Name)
        $PK.IndexedColumns.Add($IdxCol) 
        $PK.Create()

        Write-Host ("Table BreakingChangeWeighting created successfully") -ForegroundColor Green
    }
    else
    {
        Write-Host ("Table BreakingChangeWeighting already exists") -ForegroundColor Yellow
    }

    #Create views
    $vwCheck1 = $db.Views | Where {$_.Name -eq "DatabaseCategoryRanking"}
    if(!$vwCheck1)
    {
        $vwDatabaseCategoryRanking = New-Object -TypeName Microsoft.SqlServer.Management.SMO.View -argumentlist $db, "DatabaseCategoryRanking", "dbo"  
  
        $vwDatabaseCategoryRanking.TextHeader = "CREATE VIEW [dbo].[DatabaseCategoryRanking] AS"  
        $vwDatabaseCategoryRanking.TextBody=@"
WITH DatabaseRanking
AS
(
SELECT	[Name]
		,ChangeCategory
		,COUNT(*) AS "NumberOfIssues"
		,(CONVERT(NUMERIC(5,2),COUNT(*))/(SELECT CONVERT(NUMERIC(5,2),COUNT(*)) FROM reportdata r2 Where r1.[name] = r2.[name])) * 100 AS "ChangeCategoryPercentage"
FROM	reportdata r1
GROUP BY [Name], ChangeCategory
)
SELECT	[Name] AS "DatabaseName"
	,ChangeCategory
	,ChangeCategoryPercentage
FROM DatabaseRanking;
"@
  
        $vwDatabaseCategoryRanking.Create()  
        Write-Host ("View DatabaseCategoryRanking created successfully") -ForegroundColor Green
    }
    else
    {
        Write-Host ("View DatabaseCategoryRanking already exists") -ForegroundColor Yellow
    }
        
    $vwCheck2 = $db.Views | Where {$_.Name -eq "UpgradeSuccessRanking"}
    if(!$vwCheck2)
    {
        $vwUpgradeSuccessRanking = New-Object -TypeName Microsoft.SqlServer.Management.SMO.View -argumentlist $db, "UpgradeSuccessRanking", "dbo"  
  
        $vwUpgradeSuccessRanking.TextHeader = "CREATE VIEW [dbo].[UpgradeSuccessRanking] AS"  
        $vwUpgradeSuccessRanking.TextBody=@"
WITH issuecount
AS
(
-- currently doesn't take into account diminishing returns for repeating issues
-- removed NotDefined as these are for feature parity, not migration blockers and should therefore be excluded in calculations
SELECT	InstanceName
		,NAME
		,TargetCompatibilityLevel
		,COALESCE(CASE changecategory WHEN 'BehaviorChange' THEN COUNT(*) END,0) AS 'BehaviorChange'
		,COALESCE(CASE changecategory WHEN 'Deprecated' THEN COUNT(*) END,0) AS 'DeprecatedCount'
		,COALESCE(CASE changecategory WHEN 'BreakingChange' THEN SUM(ChangeRank) END ,0) AS 'BreakingChange'
		--,COALESCE(CASE changecategory WHEN 'NotDefined' THEN COUNT(*) END,0) AS 'NotDefined'
		,COALESCE(CASE changecategory WHEN 'MigrationBlocker' THEN COUNT(*) END,0) AS 'MigrationBlocker'
FROM reportdata rd
LEFT JOIN BreakingChangeWeighting bcw
ON rd.RuleId = bcw.ruleid
WHERE changecategory != 'NotDefined'
and TargetCompatibilityLevel != 'NA'
GROUP BY InstanceName,name, changecategory, TargetCompatibilityLevel
),
distinctissues
AS
(
SELECT	InstanceName
		,NAME
		,TargetCompatibilityLevel
		,MAX(BehaviorChange) AS 'BehaviorChange'
		,MAX(DeprecatedCount) AS 'DeprecatedCount'
		,MAX(BreakingChange) AS 'BreakingChange'
		--,MAX(NotDefined) AS 'NotDefined'
		,MAX(MigrationBlocker) AS 'MigrationBlocker'
FROM	issuecount
GROUP BY InstanceName,name, TargetCompatibilityLevel
),
IssueTotaled
AS
(
SELECT	*, behaviorchange + deprecatedcount + breakingchange + MigrationBlocker AS 'Total'
FROM	distinctissues 
),
RankedDatabases
AS
(
SELECT	InstanceName
		,Name
		,TargetCompatibilityLevel
		,CAST(100-((BehaviorChange + 0.00) / (total + 0.00)) * 100 AS DECIMAL(5,2)) AS 'BehaviorChange'
		,CAST(100-((DeprecatedCount + 0.00) / (total + 0.00)) * 100 AS DECIMAL(5,2)) AS 'DeprecatedCount'
		,CAST(100-((BreakingChange + 0.00) / (total + 0.00)) * 100 AS DECIMAL(5,2)) AS 'BreakingChange'
		--,CAST(100-((NotDefined + 0.00) / (total + 0.00)) * 100 AS DECIMAL(5,2)) AS 'NotDefined'
		,CAST(100-((MigrationBlocker + 0.00) / (total + 0.00)) * 100 AS DECIMAL(5,2)) AS 'MigrationBlocker'
FROM	IssueTotaled
)
-- This section will ensure that if there are 0 issues in a category we return 1.  This ensures the reports show data
SELECT	 InstanceName
		,[Name]
		,TargetCompatibilityLevel
		,CASE  WHEN BehaviorChange > 0 THEN BehaviorChange ELSE 1 END AS "BehaviorChange"
		,CASE  WHEN DeprecatedCount > 0 THEN DeprecatedCount ELSE 1 END AS "DeprecatedCount"
		,CASE  WHEN BreakingChange > 0 THEN BreakingChange ELSE 1 END AS "BreakingChange"
		--,CASE  WHEN NotDefined > 0 THEN NotDefined ELSE 1 END AS "NotDefined" 
		,CASE  WHEN MigrationBlocker > 0 THEN MigrationBlocker ELSE 1 END AS "MigrationBlocker" 
FROM	RankedDatabases
"@
  
        $vwUpgradeSuccessRanking.Create() 
        Write-Host ("View UpgradeSuccessRanking created successfully") -ForegroundColor Green 
    }
    else
    {
        Write-Host ("View UpgradeSuccessRanking already exists") -ForegroundColor Yellow
    }

    $vwCheck3 = $db.Views | Where {$_.Name -eq "UpgradeSuccessRanking_OnPrem"}
    if(!$vwCheck3)
    {
        $vwUpgradeSuccessRankingop = New-Object -TypeName Microsoft.SqlServer.Management.SMO.View -argumentlist $db, "UpgradeSuccessRanking_OnPrem", "dbo"  
  
        $vwUpgradeSuccessRankingop.TextHeader = "CREATE VIEW [dbo].[UpgradeSuccessRanking_OnPrem] AS"  
        $vwUpgradeSuccessRankingop.TextBody=@"
WITH issuecount
AS
(
-- currently doesn't take into account diminishing returns for repeating issues
-- removed NotDefined as these are for feature parity, not migration blockers and should therefore be excluded in calculations
SELECT	InstanceName
		,NAME
		,TargetCompatibilityLevel
		,COALESCE(CASE changecategory WHEN 'BehaviorChange' THEN COUNT(*) END,0) AS 'BehaviorChange'
		,COALESCE(CASE changecategory WHEN 'Deprecated' THEN COUNT(*) END,0) AS 'DeprecatedCount'
		,COALESCE(CASE changecategory WHEN 'BreakingChange' THEN SUM(ChangeRank) END ,0) AS 'BreakingChange'
FROM	ReportData rd
LEFT JOIN BreakingChangeWeighting bcw
	ON rd.RuleId = bcw.ruleid
WHERE	ChangeCategory != 'NotDefined'
	AND TargetCompatibilityLevel != 'NA'
	AND AssessmentTarget IN ('SqlServer2012', 'SqlServer2014', 'SqlServer2016' ,'SqlServer2017','SqlServer2019', 'SqlServer2022')
GROUP BY InstanceName,name, changecategory, TargetCompatibilityLevel
),
distinctissues
AS
(
SELECT	InstanceName
		,NAME
		,TargetCompatibilityLevel
		,MAX(BehaviorChange) AS 'BehaviorChange'
		,MAX(DeprecatedCount) AS 'DeprecatedCount'
		,MAX(BreakingChange) AS 'BreakingChange'
FROM	issuecount
GROUP BY InstanceName,name, TargetCompatibilityLevel
),
IssueTotaled
AS
(
SELECT	*, behaviorchange + deprecatedcount + breakingchange AS 'Total'
FROM	distinctissues 
),
RankedDatabases
AS
(
SELECT	InstanceName
		,Name
		,TargetCompatibilityLevel
		,CAST(100-((BehaviorChange + 0.00) / (total + 0.00)) * 100 AS DECIMAL(5,2)) AS 'BehaviorChange'
		,CAST(100-((DeprecatedCount + 0.00) / (total + 0.00)) * 100 AS DECIMAL(5,2)) AS 'DeprecatedCount'
		,CAST(100-((BreakingChange + 0.00) / (total + 0.00)) * 100 AS DECIMAL(5,2)) AS 'BreakingChange'
FROM	IssueTotaled
)
-- This section will ensure that if there are 0 issues in a category we return 1.  This ensures the reports show data
SELECT	 InstanceName
		,[Name]
		,TargetCompatibilityLevel
		,CASE  WHEN BehaviorChange > 0 THEN BehaviorChange ELSE 1 END AS "BehaviorChange"
		,CASE  WHEN DeprecatedCount > 0 THEN DeprecatedCount ELSE 1 END AS "DeprecatedCount"
		,CASE  WHEN BreakingChange > 0 THEN BreakingChange ELSE 1 END AS "BreakingChange"
FROM	RankedDatabases

"@
  
        $vwUpgradeSuccessRankingop.Create() 
        Write-Host ("View UpgradeSuccessRanking_OnPrem created successfully") -ForegroundColor Green 
    }
    else
    {
        Write-Host ("View UpgradeSuccessRanking_OnPrem already exists") -ForegroundColor Yellow
    }


    $vwCheck4 = $db.Views | Where {$_.Name -eq "UpgradeSuccessRanking_Azure"}
    if(!$vwCheck4)
    {
        $vwUpgradeSuccessRankingaz = New-Object -TypeName Microsoft.SqlServer.Management.SMO.View -argumentlist $db, "UpgradeSuccessRanking_Azure", "dbo"  
  
        $vwUpgradeSuccessRankingaz.TextHeader = "CREATE VIEW [dbo].[UpgradeSuccessRanking_Azure] AS"  
        $vwUpgradeSuccessRankingaz.TextBody=@"
WITH issuecount
AS
(
-- currently doesn't take into account diminishing returns for repeating issues
-- removed NotDefined as these are for feature parity, not migration blockers and should therefore be excluded in calculations
SELECT	InstanceName
		,NAME
		,TargetCompatibilityLevel
		,COALESCE(CASE changecategory WHEN 'BehaviorChange' THEN COUNT(*) END,0) AS 'BehaviorChange'
		,COALESCE(CASE changecategory WHEN 'Deprecated' THEN COUNT(*) END,0) AS 'DeprecatedCount'
		,COALESCE(CASE changecategory WHEN 'BreakingChange' THEN SUM(ChangeRank) END ,0) AS 'BreakingChange'
		,COALESCE(CASE changecategory WHEN 'MigrationBlocker' THEN COUNT(*) END,0) AS 'MigrationBlocker'
FROM	ReportData rd
LEFT JOIN BreakingChangeWeighting bcw
	ON	rd.RuleId = bcw.ruleid
WHERE	changecategory != 'NotDefined'
	AND TargetCompatibilityLevel != 'NA'
	AND AssessmentTarget = 'AzureSQLDatabaseV12'
GROUP BY InstanceName, [Name], changecategory, TargetCompatibilityLevel
),
distinctissues
AS
(
SELECT	InstanceName
		,[Name]
		,TargetCompatibilityLevel
		,MAX(BehaviorChange) AS 'BehaviorChange'
		,MAX(DeprecatedCount) AS 'DeprecatedCount'
		,MAX(BreakingChange) AS 'BreakingChange'
		,MAX(MigrationBlocker) AS 'MigrationBlocker'
FROM	issuecount
GROUP BY InstanceName, [Name], TargetCompatibilityLevel
),
IssueTotaled
AS
(
SELECT	*, behaviorchange + deprecatedcount + breakingchange + MigrationBlocker AS 'Total'
FROM	distinctissues 
),
RankedDatabases
AS
(
SELECT	InstanceName
		,[Name]
		,TargetCompatibilityLevel
		,CAST(100-((BehaviorChange + 0.00) / (total + 0.00)) * 100 AS DECIMAL(5,2)) AS 'BehaviorChange'
		,CAST(100-((DeprecatedCount + 0.00) / (total + 0.00)) * 100 AS DECIMAL(5,2)) AS 'DeprecatedCount'
		,CAST(100-((BreakingChange + 0.00) / (total + 0.00)) * 100 AS DECIMAL(5,2)) AS 'BreakingChange'
		,CAST(100-((MigrationBlocker + 0.00) / (total + 0.00)) * 100 AS DECIMAL(5,2)) AS 'MigrationBlocker'
FROM	IssueTotaled
)
-- This section will ensure that if there are 0 issues in a category we return 1.  This ensures the reports show data
SELECT	 InstanceName
		,[Name]
		,TargetCompatibilityLevel
		,CASE  WHEN BehaviorChange > 0 THEN BehaviorChange ELSE 1 END AS "BehaviorChange"
		,CASE  WHEN DeprecatedCount > 0 THEN DeprecatedCount ELSE 1 END AS "DeprecatedCount"
		,CASE  WHEN BreakingChange > 0 THEN BreakingChange ELSE 1 END AS "BreakingChange"
		,CASE  WHEN MigrationBlocker > 0 THEN MigrationBlocker ELSE 1 END AS "MigrationBlocker" 
FROM	RankedDatabases
"@
  
        $vwUpgradeSuccessRankingaz.Create() 
        Write-Host ("View UpgradeSuccessRanking_Azure created successfully") -ForegroundColor Green 
    }
    else
    {
        Write-Host ("View UpgradeSuccessRanking_Azure already exists") -ForegroundColor Yellow
    }

    #Create Table Types
    $ttCheck = $db.UserDefinedTableTypes | Where {$_.Name -eq "JSONResults"}
    if(!$ttCheck)
    {
        $JSONResultstt = New-Object -TypeName Microsoft.SqlServer.Management.Smo.UserDefinedTableType -ArgumentList $db, "JSONResults"
        
        $col1 = New-Object Microsoft.SqlServer.Management.Smo.Column($JSONResultstt, "ImportDate", [Microsoft.SqlServer.Management.Smo.DataType]::DateTime)
        $col2 = New-Object Microsoft.SqlServer.Management.Smo.Column($JSONResultstt, "InstanceName", [Microsoft.SqlServer.Management.Smo.DataType]::VarChar(50))
        $col3 = New-Object Microsoft.SqlServer.Management.Smo.Column($JSONResultstt, "Status", [Microsoft.SqlServer.Management.Smo.DataType]::VarChar(30))
        $col4 = New-Object Microsoft.SqlServer.Management.Smo.Column($JSONResultstt, "Name", [Microsoft.SqlServer.Management.Smo.DataType]::VarChar(255))
        $col5 = New-Object Microsoft.SqlServer.Management.Smo.Column($JSONResultstt, "SizeMB", [Microsoft.SqlServer.Management.Smo.DataType]::VarChar(30))
        $col6 = New-Object Microsoft.SqlServer.Management.Smo.Column($JSONResultstt, "SourceCompatibilityLevel", [Microsoft.SqlServer.Management.Smo.DataType]::VarChar(15))
        $col7 = New-Object Microsoft.SqlServer.Management.Smo.Column($JSONResultstt, "TargetCompatibilityLevel", [Microsoft.SqlServer.Management.Smo.DataType]::VarChar(15))
        $col8 = New-Object Microsoft.SqlServer.Management.Smo.Column($JSONResultstt, "Category", [Microsoft.SqlServer.Management.Smo.DataType]::VarChar(30))
        $col9 = New-Object Microsoft.SqlServer.Management.Smo.Column($JSONResultstt, "Severity", [Microsoft.SqlServer.Management.Smo.DataType]::VarChar(15))
        $col10 = New-Object Microsoft.SqlServer.Management.Smo.Column($JSONResultstt, "ChangeCategory", [Microsoft.SqlServer.Management.Smo.DataType]::VarChar(20))
        $col11 = New-Object Microsoft.SqlServer.Management.Smo.Column($JSONResultstt, "RuleId", [Microsoft.SqlServer.Management.Smo.DataType]::VarChar(100))
        $col12 = New-Object Microsoft.SqlServer.Management.Smo.Column($JSONResultstt, "Title", [Microsoft.SqlServer.Management.Smo.DataType]::VarCharMax)
        $col13 = New-Object Microsoft.SqlServer.Management.Smo.Column($JSONResultstt, "Impact", [Microsoft.SqlServer.Management.Smo.DataType]::VarCharMax)
        $col14 = New-Object Microsoft.SqlServer.Management.Smo.Column($JSONResultstt, "Recommendation", [Microsoft.SqlServer.Management.Smo.DataType]::VarCharMax)
        $col15 = New-Object Microsoft.SqlServer.Management.Smo.Column($JSONResultstt, "MoreInfo", [Microsoft.SqlServer.Management.Smo.DataType]::VarCharMax)
        $col16 = New-Object Microsoft.SqlServer.Management.Smo.Column($JSONResultstt, "ImpactedObjectName", [Microsoft.SqlServer.Management.Smo.DataType]::VarChar(255))
        $col17 = New-Object Microsoft.SqlServer.Management.Smo.Column($JSONResultstt, "ImpactedObjectType", [Microsoft.SqlServer.Management.Smo.DataType]::VarChar(50))
        $col18 = New-Object Microsoft.SqlServer.Management.Smo.Column($JSONResultstt, "ImpactDetail", [Microsoft.SqlServer.Management.Smo.DataType]::VarCharMax)
        $col19 = New-Object Microsoft.SqlServer.Management.Smo.Column($JSONResultstt, "DBOwner", [Microsoft.SqlServer.Management.Smo.DataType]::VarChar(128))
        $col20 = New-Object Microsoft.SqlServer.Management.Smo.Column($JSONResultstt, "AssessmentTarget", [Microsoft.SqlServer.Management.Smo.DataType]::VarChar(50))
        $col21 = New-Object Microsoft.SqlServer.Management.Smo.Column($JSONResultstt, "AssessmentName", [Microsoft.SqlServer.Management.Smo.DataType]::VarChar(128))
      
        $JSONResultstt.Columns.Add($col1)
        $JSONResultstt.Columns.Add($col2)
        $JSONResultstt.Columns.Add($col3)
        $JSONResultstt.Columns.Add($col4)
        $JSONResultstt.Columns.Add($col5)
        $JSONResultstt.Columns.Add($col6)
        $JSONResultstt.Columns.Add($col7)
        $JSONResultstt.Columns.Add($col8)
        $JSONResultstt.Columns.Add($col9)
        $JSONResultstt.Columns.Add($col10)
        $JSONResultstt.Columns.Add($col11)
        $JSONResultstt.Columns.Add($col12)
        $JSONResultstt.Columns.Add($col13)
        $JSONResultstt.Columns.Add($col14)
        $JSONResultstt.Columns.Add($col15)
        $JSONResultstt.Columns.Add($col16)
        $JSONResultstt.Columns.Add($col17)
        $JSONResultstt.Columns.Add($col18)  
        $JSONResultstt.Columns.Add($col19)
        $JSONResultstt.Columns.Add($col20)   
        $JSONResultstt.Columns.Add($col21) 

        $JSONResultstt.Create()
        Write-Host ("Table Type JSONResults created successfully") -ForegroundColor Green 
    }
    else
    {
        Write-Host ("Table Type JSONResults already exists") -ForegroundColor Yellow
    }
      
    $ttCheck2 = $db.UserDefinedTableTypes | Where {$_.Name -eq "AzureFeatureParityResults"}
    if(!$ttCheck2)
    {
        $AzureParityResultstt = New-Object -TypeName Microsoft.SqlServer.Management.Smo.UserDefinedTableType -ArgumentList $db, "AzureFeatureParityResults"
        
        $col1 = New-Object Microsoft.SqlServer.Management.Smo.Column($AzureParityResultstt, "ImportDate", [Microsoft.SqlServer.Management.Smo.DataType]::DateTime)
        $col2 = New-Object Microsoft.SqlServer.Management.Smo.Column($AzureParityResultstt, "ServerName", [Microsoft.SqlServer.Management.Smo.DataType]::VarChar(128))
        $col3 = New-Object Microsoft.SqlServer.Management.Smo.Column($AzureParityResultstt, "Version", [Microsoft.SqlServer.Management.Smo.DataType]::VarChar(15))
        $col4 = New-Object Microsoft.SqlServer.Management.Smo.Column($AzureParityResultstt, "Status", [Microsoft.SqlServer.Management.Smo.DataType]::VarChar(10))
        $col5 = New-Object Microsoft.SqlServer.Management.Smo.Column($AzureParityResultstt, "Category", [Microsoft.SqlServer.Management.Smo.DataType]::VarChar(50))
        $col6 = New-Object Microsoft.SqlServer.Management.Smo.Column($AzureParityResultstt, "Severity", [Microsoft.SqlServer.Management.Smo.DataType]::VarChar(50))
        $col7 = New-Object Microsoft.SqlServer.Management.Smo.Column($AzureParityResultstt, "FeatureParityCategory", [Microsoft.SqlServer.Management.Smo.DataType]::VarChar(50))
        $col8 = New-Object Microsoft.SqlServer.Management.Smo.Column($AzureParityResultstt, "RuleID", [Microsoft.SqlServer.Management.Smo.DataType]::VarChar(100))
        $col9 = New-Object Microsoft.SqlServer.Management.Smo.Column($AzureParityResultstt, "Title", [Microsoft.SqlServer.Management.Smo.DataType]::VarChar(1000))
        $col10 = New-Object Microsoft.SqlServer.Management.Smo.Column($AzureParityResultstt, "Impact", [Microsoft.SqlServer.Management.Smo.DataType]::VarChar(1000))
        $col11 = New-Object Microsoft.SqlServer.Management.Smo.Column($AzureParityResultstt, "Recommendation", [Microsoft.SqlServer.Management.Smo.DataType]::VarCharMax)
        $col12 = New-Object Microsoft.SqlServer.Management.Smo.Column($AzureParityResultstt, "MoreInfo", [Microsoft.SqlServer.Management.Smo.DataType]::VarCharMax)
        $col13 = New-Object Microsoft.SqlServer.Management.Smo.Column($AzureParityResultstt, "ImpactedDatabasename", [Microsoft.SqlServer.Management.Smo.DataType]::VarChar(128))
        $col14 = New-Object Microsoft.SqlServer.Management.Smo.Column($AzureParityResultstt, "ImpactedObjectType", [Microsoft.SqlServer.Management.Smo.DataType]::VarChar(30))
        $col15 = New-Object Microsoft.SqlServer.Management.Smo.Column($AzureParityResultstt, "ImpactDetail", [Microsoft.SqlServer.Management.Smo.DataType]::VarCharMax)

      
        $AzureParityResultstt.Columns.Add($col1)
        $AzureParityResultstt.Columns.Add($col2)
        $AzureParityResultstt.Columns.Add($col3)
        $AzureParityResultstt.Columns.Add($col4)
        $AzureParityResultstt.Columns.Add($col5)
        $AzureParityResultstt.Columns.Add($col6)
        $AzureParityResultstt.Columns.Add($col7)
        $AzureParityResultstt.Columns.Add($col8)
        $AzureParityResultstt.Columns.Add($col9)
        $AzureParityResultstt.Columns.Add($col10)
        $AzureParityResultstt.Columns.Add($col11)
        $AzureParityResultstt.Columns.Add($col12)
        $AzureParityResultstt.Columns.Add($col13)
        $AzureParityResultstt.Columns.Add($col14)
        $AzureParityResultstt.Columns.Add($col15)

        $AzureParityResultstt.Create()
        Write-Host ("Table Type AzureFeatureParityResults created successfully") -ForegroundColor Green 
    }
    else
    {
        Write-Host ("Table Type AzureFeatureParityResults already exists") -ForegroundColor Yellow
    }  
      
    #Create Stored Procedures
    $procCheck = $db.StoredProcedures | Where {$_.Name -eq "JSONResults_Insert"}
    if(!$procCheck)
    {
        $JSONResults_Insert = New-Object -TypeName Microsoft.SqlServer.Management.Smo.StoredProcedure -ArgumentList $db, "JSONResults_Insert", "dbo"
        
        $JSONResults_Insert.TextHeader = "CREATE PROCEDURE dbo.JSONResults_Insert @JSONResults JSONResults READONLY AS"
        $JSONResults_Insert.TextBody = @"
BEGIN

INSERT INTO dbo.ReportData (ImportDate, InstanceName, [Status], [Name], SizeMB, SourceCompatibilityLevel, TargetCompatibilityLevel, Category, Severity, ChangeCategory, RuleId, Title, Impact, Recommendation, MoreInfo, ImpactedObjectName, ImpactedObjectType, ImpactDetail, DBOwner, AssessmentTarget, AssessmentName)
SELECT ImportDate, InstanceName, [Status], [Name], SizeMB, SourceCompatibilityLevel, TargetCompatibilityLevel, Category, Severity, ChangeCategory, RuleId, Title, Impact, Recommendation, MoreInfo, ImpactedObjectName, ImpactedObjectType, ImpactDetail, DBOwner, AssessmentTarget, AssessmentName
FROM @JSONResults

END
"@

        $JSONResults_Insert.Create()
        Write-Host ("Stored Procedure JSONNResults_Insert created successfully") -ForegroundColor Green 
    }
    else
    {
        Write-Host ("Stored Procedure JSONNResults_Insert already exists") -ForegroundColor Yellow
    }

    $procCheck2 = $db.StoredProcedures | Where {$_.Name -eq "AzureFeatureParityResults_Insert"}
    if(!$procCheck2)
    {
        $AzureFeatureParityResults_Insert = New-Object -TypeName Microsoft.SqlServer.Management.Smo.StoredProcedure -ArgumentList $db, "AzureFeatureParityResults_Insert", "dbo"
        
        $AzureFeatureParityResults_Insert.TextHeader = "CREATE PROCEDURE dbo.AzureFeatureParityResults_Insert @AzureFeatureParityResults AzureFeatureParityResults READONLY AS"
        $AzureFeatureParityResults_Insert.TextBody = @"
BEGIN

INSERT INTO dbo.AzureFeatureParity (ImportDate, ServerName, Version, Status, Category, Severity, FeatureParityCategory, RuleID, Title, Impact, Recommendation, MoreInfo, ImpactedDatabasename, ImpactedObjectType, ImpactDetail)
SELECT ImportDate, ServerName, Version, Status, Category, Severity, FeatureParityCategory, RuleID, Title, Impact, Recommendation, MoreInfo, ImpactedDatabasename, ImpactedObjectType, ImpactDetail
FROM @AzureFeatureParityResults

END
"@

        $AzureFeatureParityResults_Insert.Create()
        Write-Host ("Stored Procedure AzureFeatureParityResults_Insert created successfully") -ForegroundColor Green 
    }
    else
    {
        Write-Host ("Stored Procedure AzureFeatureParityResults_Insert already exists") -ForegroundColor Yellow
    }

    # END CREATE DATABASE OBJECTS #


    #Make processed directory inside the folder that contains the json files
    if(!$jsonDirectory.EndsWith("\"))
    {
        $jsonDirectory = "$jsonDirectory\"
    }
    $processedDir = "$jsonDirectory`Processed"

    if((Test-Path $processedDir) -eq $false)
    {
        new-item $processedDir -ItemType directory 
        Write-Host ("Processed directory created successfully at [$processDir]") -ForegroundColor Green
    }
    else
    {
        Write-Host ("Processed directory already exists") -ForegroundColor Yellow
    }
       
    # if there are no files to process stop importer
    $FileCheck = Get-ChildItem $jsonDirectory -Filter *.JSON
    if($FileCheck.Count -eq 0)
    {
        Write-Host ("There are no JSON assessment files to process") -ForegroundColor Yellow 
        Break
    }
    
    
    $connectionString = "Server=$serverName;Database=$databaseName;Trusted_Connection=True;"

    #Populate the breaking change reference data
    $RefDataCheck = $db.Tables | Where {$_.Name -eq "BreakingChangeWeighting"} | Select RowCount
    if($RefDataCheck.RowCount -eq 0)
    {

        #populate static data into BreakingChangeWeighting
                
        $CommandText = @'
INSERT INTO BreakingChangeWeighting VALUES ('Microsoft.Rules.Data.Upgrade.UR00001','Syntax issue on the source server',1,1,1),
('Microsoft.Rules.Data.Upgrade.UR00006','BACKUP LOG WITH NO_LOG|TRUNCATE_ONLY statements are not supported',1,1,1),
('Microsoft.Rules.Data.Upgrade.UR00007','BACKUP/RESTORE TRANSACTION statements are deprecated or discontinued',1,1,1),
('Microsoft.Rules.Data.Upgrade.UR00013','COMPUTE clause is not allowed in database compatibility 110',1,1,1),
('Microsoft.Rules.Data.Upgrade.UR00020','Read-only databases cannot be upgraded',1,1,1),
('Microsoft.Rules.Data.Upgrade.UR00021','Verify all filegroups are writeable during the upgrade process',1,1,1),
('Microsoft.Rules.Data.Upgrade.UR00023','SQL Server native SOAP support is discontinued in SQL Server 2014 and above',1,1,1),
('Microsoft.Rules.Data.Upgrade.UR00044','Remove user-defined type (UDT)s named after the reserved GEOMETRY and GEOGRAPHY data types',1,1,1),
('Microsoft.Rules.Data.Upgrade.UR00050','Table hints in indexed view definitions are ignored in compatibility mode 80 and are not allowed in compatibility mode 90 or above',1,1,1),
('Microsoft.Rules.Data.Upgrade.UR00058','After upgrade, new reserved keywords cannot be used as identifiers',1,1,1),
('Microsoft.Rules.Data.Upgrade.UR00062','Tables and Columns named NEXT may lead to an error using compatibility Level 110 and above',1,1,1),
('Microsoft.Rules.Data.Upgrade.UR00086','XML is a reserved system type name',1,1,1),
('Microsoft.Rules.Data.Upgrade.UR00110','New column in output of sp_helptrigger may impact applications',1,1,1),
('Microsoft.Rules.Data.Upgrade.UR00113','SQL Mail has been discontinued',1,1,1),
('Microsoft.Rules.Data.Upgrade.UR00300','Remove the use of PASSWORD in BACKUP command',1,1,1),
('Microsoft.Rules.Data.Upgrade.UR00301','WITH CHECK OPTION is not supported in views that contain TOP in compatibility mode 90 and above',1,1,1),
('Microsoft.Rules.Data.Upgrade.UR00302','Discontinued DBCC commands referenced in your T-SQL objects',1,1,1),
('Microsoft.Rules.Data.Upgrade.UR00308','Legacy style RAISERROR calls should be replaced with modern equivalents',1,1,1),
('Microsoft.Rules.Data.Upgrade.UR00311','Detected statements that reference removed system stored procedures that are not available in database compatibility level 100 and higher',1,1,1),
('Microsoft.Rules.Data.Upgrade.UR00318','FOR BROWSE is not allowed in views in 90 or later compatibility modes',1,1,1),
('Microsoft.Rules.Data.Upgrade.UR00321','Non ANSI style left outer join usage',1,1,1),
('Microsoft.Rules.Data.Upgrade.UR00322','Non ANSI style right outer join usage',1,1,1),
('Microsoft.Rules.Data.Upgrade.UR00326','Constant expressions are not allowed in the ORDER BY clause in 90 or later compatibility modes',1,1,1),
('Microsoft.Rules.Data.Upgrade.UR00332','FASTFIRSTROW table hint usage',1,1,1),
('Microsoft.Rules.Data.Upgrade.UR00336','Certain XPath functions are not allowed in OPENXML queries',1,1,1)
'@

        $conn = New-Object System.Data.SqlClient.SqlConnection $connectionString 
        $conn.Open() | Out-Null

        $cmd = New-Object System.Data.SqlClient.SqlCommand 
        $cmd.Connection = $conn
        $cmd.CommandType = [System.Data.CommandType]"Text"
        $cmd.CommandText= $CommandText
              
        $ds=New-Object system.Data.DataSet
        $da=New-Object system.Data.SqlClient.SqlDataAdapter($cmd)
        $da.fill($ds)
        $conn.Close()
    }
   
    # importer for SQL2014 and previous versions. Done via PowerShell
    Get-ChildItem $jsonDirectory -Filter *.JSON | 
    Foreach-Object {
        
        $filename = $_.FullName

        #ReportData datatable                                                                                                                                                                                                                                                                                {                   
        $datatable = New-Object -type system.data.datatable
        $datatable.columns.add("ImportDate",[DateTime]) | Out-Null
        $datatable.columns.add("InstanceName",[String]) | Out-Null
        $datatable.columns.add("Status",[String]) | Out-Null
        $datatable.columns.add("Name",[String]) | Out-Null
        $datatable.columns.add("SizeMB",[Decimal]) | Out-Null
        $datatable.columns.add("SourceCompatibilityLevel",[String]) | Out-Null
        $datatable.columns.add("TargetCompatibilityLevel",[String]) | Out-Null
        $datatable.columns.add("Category",[String]) | Out-Null
        $datatable.columns.add("Severity",[String]) | Out-Null
        $datatable.columns.add("ChangeCategory",[String]) | Out-Null
        $datatable.columns.add("RuleId",[String]) | Out-Null
        $datatable.columns.add("Title",[String]) | Out-Null
        $datatable.columns.add("Impact",[String]) | Out-Null
        $datatable.columns.add("Recommendation",[String]) | Out-Null
        $datatable.columns.add("MoreInfo",[String]) | Out-Null
        $datatable.columns.add("ImpactedObjectName",[String]) | Out-Null
        $datatable.columns.add("ImpactedObjectType",[String]) | Out-Null
        $datatable.columns.add("ImpactDetail",[string]) | Out-Null
        $datatable.columns.add("DBOwner",[string]) | Out-Null
        $datatable.columns.add("AssessmentTarget",[string]) | Out-Null
        $datatable.columns.add("AssessmentName",[string]) | Out-Null

        #AzureFeatureParity datatable
        $azuredatatable = New-Object -type system.data.datatable
        $azuredatatable.columns.add("ImportDate",[DateTime]) | Out-Null
        $azuredatatable.columns.add("ServerName",[String]) | Out-Null
        $azuredatatable.columns.add("Version",[String]) | Out-Null
        $azuredatatable.columns.add("Status",[String]) | Out-Null
        $azuredatatable.columns.add("Category",[String]) | Out-Null
        $azuredatatable.columns.add("Severity",[String]) | Out-Null
        $azuredatatable.columns.add("FeatureParityCategory",[String]) | Out-Null
        $azuredatatable.columns.add("RuleID",[String]) | Out-Null
        $azuredatatable.columns.add("Title",[String]) | Out-Null
        $azuredatatable.columns.add("Impact",[String]) | Out-Null
        $azuredatatable.columns.add("Recommendation",[String]) | Out-Null
        $azuredatatable.columns.add("MoreInfo",[String]) | Out-Null
        $azuredatatable.columns.add("ImpactedDatabasename",[String]) | Out-Null
        $azuredatatable.columns.add("ImpactedObjectType",[String]) | Out-Null
        $azuredatatable.columns.add("ImpactDetail",[String]) | Out-Null


        $processStartTime = Get-Date
        $datetime = Get-Date                    
        $content = Get-Content $_.FullName -Raw
        
        # when a database assessment fails the assessment recommendations and impacted objects arrays
        # will be blank.  Setting them to default values allows for the errors to be captured
        $blankAssessmentRecommendations =   (New-Object PSObject |
                                           Add-Member -PassThru NoteProperty CompatibilityLevel NA |
                                           Add-Member -PassThru NoteProperty Category NA          |
                                           Add-Member -PassThru NoteProperty Severity NA     |
                                           Add-Member -PassThru NoteProperty ChangeCategory NA |
                                           Add-Member -PassThru NoteProperty RuleId NA |
                                           Add-Member -PassThru NoteProperty Title NA |
                                           Add-Member -PassThru NoteProperty Impact NA |
                                           Add-Member -PassThru NoteProperty Recommendation NA |
                                           Add-Member -PassThru NoteProperty MoreInfo NA |
                                           Add-Member -PassThru NoteProperty ImpactedObjects NA
                                        ) 
        
        $blankImpactedObjects = (New-Object PSObject |
                                           Add-Member -PassThru NoteProperty Name NA |
                                           Add-Member -PassThru NoteProperty ObjectType NA          |
                                           Add-Member -PassThru NoteProperty ImpactDetail NA     
                                        )

        $blankImpactedDatabases = (New-Object PSObject |
                                           Add-Member -PassThru NoteProperty Name NA |
                                           Add-Member -PassThru NoteProperty ObjectType NA          |
                                           Add-Member -PassThru NoteProperty ImpactDetail NA     
                                        ) 


        # Start looping through each JSON array
        
        #fill dataset for ReportData table
        foreach($obj in (ConvertFrom-Json $content)) #level 1, the actual file
        {          
            foreach($database in $obj.Databases) #level 2, the sources
            {
                $database.AssessmentRecommendations = if($database.AssessmentRecommendations.Length -eq 0) {$blankAssessmentRecommendations } else {$database.AssessmentRecommendations}
                
                foreach($assessment in $database.AssessmentRecommendations) #level 3, the assessment
                {
                    
                    $assessment.ImpactedObjects = if ($assessment.ImpactedObjects.Length -eq 0) {$blankImpactedObjects} else {$assessment.ImpactedObjects}

                    foreach($impactedobj in $assessment.ImpactedObjects) #level 4, the impacted objects
                    {
                                                
                        #TODO Get date here will eventually be replace with timestamp from JSON file
                        $datatable.rows.add((Get-Date).toString(), $database.ServerName, $database.Status, $database.Name, $database.SizeMB, $database.CompatibilityLevel, $assessment.CompatibilityLevel, $assessment.Category, $assessment.severity, $assessment.ChangeCategory, $assessment.RuleId, $assessment.Title, $assessment.Impact, $assessment.Recommendation, $assessment.MoreInfo, $impactedobj.Name, $impactedobj.ObjectType, $impactedobj.ImpactDetail, $null, $obj.TargetPlatform, $obj.Name) | Out-Null
                    }
                }
            }
        }           

        #fill data set for AzureFeatureParity table
        foreach($obj in (ConvertFrom-Json $content)) #level 1, the actual file
        {          
            foreach($serverInstances in $obj.ServerInstances) #level 2, the ServerInstances
            {
                foreach($assessment in $serverInstances.AssessmentRecommendations) #level 3, the assessment
                {
                    $impactedDatabases = if (($assessment.ImpactedDatabases -eq $null) -or ($assessment.ImpactedDatabases.Length -eq 0)) {$blankImpactedDatabases} else {$assessment.ImpactedDatabases}
                        
                    foreach($impacteddbs in $impactedDatabases) #level 4, the impacted objects
                    {                       
                        #TODO Get date here will eventually be replace with timestamp from JSON file
                        $azuredatatable.rows.add((Get-Date).toString(), $serverInstances.ServerName, $serverInstances.Version, $serverInstances.Status, $assessment.Category, $assessment.Severity, $assessment.FeatureParityCategory, $assessment.RuleId, $assessment.Title, $assessment.Impact, $assessment.Recommendation, $assessment.MoreInfo, $impacteddbs.Name, $impacteddbs.ObjectType, $impacteddbs.ImpactDetail) | Out-Null
                    }
                        
                }
            }
        }

        $rowcount_rd = $datatable.rows.Count
        $rowcount_afp = $azuredatatable.rows.Count

        $query1='dbo.JSONResults_Insert' 
        $query2='dbo.AzureFeatureParityResults_Insert'  

        #Connect
        $conn = New-Object System.Data.SqlClient.SqlConnection $connectionString 
        $conn.Open() | Out-Null

        $cmd1 = New-Object System.Data.SqlClient.SqlCommand
        $cmd1.Connection = $conn
        $cmd1.CommandType = [System.Data.CommandType]"StoredProcedure"
        $cmd1.CommandText= $query1
        $cmd1.Parameters.Add("@JSONResults" , [System.Data.SqlDbType]::Structured) | Out-Null
        $cmd1.Parameters["@JSONResults"].Value =$datatable

        $cmd2 = New-Object System.Data.SqlClient.SqlCommand
        $cmd2.Connection = $conn
        $cmd2.CommandType = [System.Data.CommandType]"StoredProcedure"
        $cmd2.CommandText= $query2
        $cmd2.Parameters.Add("@AzureFeatureParityResults" , [System.Data.SqlDbType]::Structured) | Out-Null
        $cmd2.Parameters["@AzureFeatureParityResults"].Value = $azuredatatable
                     
        $ds1=New-Object system.Data.DataSet
        $da1=New-Object system.Data.SqlClient.SqlDataAdapter($cmd1)
          
        $ds2=New-Object system.Data.DataSet
        $da2=New-Object system.Data.SqlClient.SqlDataAdapter($cmd2)
      
        # ensure that the dataset can write to the database, if not the dont move the file to processed directory
        try
        {
            $da1.fill($ds1) | Out-Null
            $da2.fill($ds2) | out-null
   
            try
            {
                Move-Item $filename $processedDir -Force
            }
            catch
            {
                write-host("Error moving file $filename to directory") -ForegroundColor Red
                $error[0]|format-list -force
            }

        }
        catch
        {
            $rowcount_rd = 0
            $rowcount_afp = 0
            write-host("Error writing results for file $filename to database") -ForegroundColor Red
            $error[0]|format-list -force
        }

        $conn.Close()

        $processEndTime = Get-Date
        $processTime = NEW-TIMESPAN -Start $processStartTime -End $processEndTime
        Write-Host("Rows Processed for ReportData Table = $rowcount_rd  Rows processed for AzureFeatureParityTable = $rowcount_afp for file $filename Total Processing Time = $processTime")

        $datatable.Clear()
        $azuredatatable.Clear()
        
    }
}

#dmaProcessor -serverName 'sqlmydb001' -databaseName 'DMA_PBI_Reporting' -jsonDirectory 'C:\myDB_Assessment_Report\DMA\json_format'
Function DMA_Report_All
{
  #Add-Type -AssemblyName Microsoft.SqlServer.Smo
  Import-Module SQlserver -DisableNameChecking
  $D=$DMA_path+"Reports"+"\"+"DMA_Report_$(get-date -f dd-MM-yyyy)"+".csv"
  Remove-Item $D
  #$host=$env:COMPUTERNAME
  $TargetServerInstance = "sqlmydb001\SQL2023"#$env:COMPUTERNAME#"sqlvs001"
  $Servername = $TargetServerInstance
  $SQLServer = New-Object -TypeName Microsoft.SqlServer.Management.Smo.Server -ArgumentList $TargetServerInstance
  $dbname = "DMAReportingAxalta"
  $query1=";WITH cte AS 
    (
SELECT [ImportDate]
      ,[InstanceName]
      ,[Status]
      ,[Name]
      ,[SizeMB]
      ,[SourceCompatibilityLevel]
      ,[TargetCompatibilityLevel]
      ,[Category]
      ,[Severity]
      ,[ChangeCategory]
      ,[RuleId]
      ,[Title]
      ,[Impact]
      ,[Recommendation]
      ,[MoreInfo]
      ,[ImpactedObjectName]
      ,[ImpactedObjectType]
      ,[ImpactDetail]
      , rn = ROW_NUMBER() OVER (PARTITION BY [InstanceName], [Name], [ImpactedObjectName] 
      ORDER BY [InstanceName], [Name], [ImpactedObjectName], [TargetCompatibilityLevel] DESC)
    FROM [DMAReportingAxalta].[dbo].[ReportData] where ChangeCategory in ('BreakingChange','MigrationBlocker','Information')
    )
    SELECT * 
    FROM cte
    WHERE rn = 2 --and ServerName in ('sqlserver01','SQLSERVER01\SQL_002')
    ORDER BY [InstanceName], [Name], [ImpactedObjectName], [TargetCompatibilityLevel]; 
    GO
  "
  #$eqecutequery1= Invoke-SQLCMD $query1  -ServerInstance $SQLServer
  $execq1=Invoke-Sqlcmd -ServerInstance $Servername -Database $dbname -Query  $query1 -EA "silentlycontinue" |Export-csv -path $D -NoTypeInformation
  #[void] [System.Windows.MessageBox]::Show( "Report Generated Sucessfully on server $SQLServer", "Script completed", "OK", "Information" )

  $css = @"
<style>
h1, h5, th { text-align: center; font-family: Segoe UI; }
table { margin: auto; font-family: Segoe UI; box-shadow: 10px 10px 5px #888; border: thin ridge grey; }
th { background: #0046c3; color: #fff; max-width: 400px; padding: 5px 10px; }
td { font-size: 11px; padding: 5px 20px; color: #000; }
tr { background: #b8d1f3; }
tr:nth-child(even) { background: #dae5f4; }
tr:nth-child(odd) { background: #b8d1f3; }
</style>
"@
  #$DMA_path="C:\myDB_Assessment_Report\"
  $DMAReportfilepath=$DMA_path +"Reports"+"\"+"DMA_Report_$(get-date -f dd-MM-yyyy)"+".csv"
  $DMAReportfilepath
  $HTMLfilepath=$DMA_path +"Reports"+"\"+"DMA_Report_$(get-date -f dd-MM-yyyy)"+".HTML"
  $HTMLfilepath
  Import-CSV $DMAReportfilepath | ConvertTo-Html -Head $css -Body "<h1>DMA Report For All SQL Servers</h1>`n<h5>Generated on $(Get-Date)</h5>" | Out-File $HTMLfilepath
  Start-Sleep -s 2
  Invoke-Expression $HTMLfilepath
}
Function All_Summary_Report
{
  #Add-Type -AssemblyName Microsoft.SqlServer.Smo
  Import-Module SQlserver -DisableNameChecking
  $Dir_Path ="C:\Axalta_Assessment_Report\"

#$Dir_Path ="C:\myDB_Assessment_Report\"
$DMA_Path =$Dir_Path
 
  $D=$DMA_path+"Reports"+"\"+"DMA_Summery_Report_$(get-date -f dd-MM-yyyy)"+".csv"
  #Remove-Item $D
  #$hostname= $env:COMPUTERNAME
  $TargetServerInstance = "sqlmydb001\SQL2023"#$env:COMPUTERNAME#"sqlvs001"
  $Servername = $TargetServerInstance
  $SQLServer = New-Object -TypeName Microsoft.SqlServer.Management.Smo.Server -ArgumentList $TargetServerInstance

  $query1="select  InstanceName,Name,TotalBreakingChanges= TotalAzureSQLDBBreakingChanges+TotalMIBreakingChanges,
						   TotalMigrationBlocker =(TotalAzureSQLDBBlocker+TotatlMIBlocker),
						   TotalAzureSQLDBInformationChanges,TotalMIInformationChanges,
 case when (TotalAzureSQLDBBreakingChanges+TotalMIBreakingChanges > 0) or (TotalAzureSQLDBBreakingChanges+TotalAzureSQLDBBlocker) > 0 then 'No' else 'Yes' end as Migrate_AzureSQL_DB,
case when (TotalMIBreakingChanges > 0) or (TotalMIBreakingChanges+TotatlMIBlocker) > 0 then 'No' else 'Yes' end as Migrate_Managed_Instance
 from
 (
select InstanceName,Name,
	(select  sum(case when r.ChangeCategory='BreakingChange' and AssessmentTarget='AzureSQLDatabase' and TargetCompatibilityLevel = 'CompatLevel150' then 1 else 0 end)  as ReportBreakinghanges from [DMAReportingAxalta].[dbo].[ReportData] r where dbinfo.InstanceName=r.InstanceName and r.Name=dbinfo.Name) TotalAzureSQLDBBreakingChanges,
	(select  sum(case when r.ChangeCategory='BreakingChange' and AssessmentTarget='ManagedSQLServer' and TargetCompatibilityLevel = 'CompatLevel150' then 1 else 0 end)  as ReportBreakinghanges from [DMAReportingAxalta].[dbo].[ReportData] r where dbinfo.InstanceName=r.InstanceName and r.Name=dbinfo.Name) TotalMIBreakingChanges,
	(select  sum(case when r.ChangeCategory='MigrationBlocker' and AssessmentTarget='AzureSQLDatabase' and TargetCompatibilityLevel = 'CompatLevel150' then 1 else 0 end)  as ReportBreakinghanges from [DMAReportingAxalta].[dbo].[ReportData] r where dbinfo.InstanceName=r.InstanceName and r.Name=dbinfo.Name) TotalAzureSQLDBBlocker,
	(select  sum(case when r.ChangeCategory='MigrationBlocker' and AssessmentTarget='ManagedSQLServer' and TargetCompatibilityLevel = 'CompatLevel150' then 1 else 0 end)  as ReportBreakinghanges from [DMAReportingAxalta].[dbo].[ReportData] r where dbinfo.InstanceName=r.InstanceName and r.Name=dbinfo.Name) TotatlMIBlocker,
	(select  sum(case when r.ChangeCategory='Information' and AssessmentTarget='AzureSQLDatabase' and TargetCompatibilityLevel = 'CompatLevel150' then 1 else 0 end)  as ReportBreakinghanges from [DMAReportingAxalta].[dbo].[ReportData] r where dbinfo.InstanceName=r.InstanceName and r.Name=dbinfo.Name) TotalAzureSQLDBInformationChanges,
	(select  sum(case when r.ChangeCategory='Information' and AssessmentTarget='ManagedSQLServer' and TargetCompatibilityLevel = 'CompatLevel150' then 1 else 0 end)  as ReportBreakinghanges from [DMAReportingAxalta].[dbo].[ReportData] r where dbinfo.InstanceName=r.InstanceName and r.Name=dbinfo.Name) TotalMIInformationChanges
	 --(select count(*) from DMA_Assessment dm where dm.Instance_Name = dbinfo.InstanceName and dm.DBNAME=dbinfo.name and dm.TargetServer='AzureSqlDatabase')  as Migrate_AzureSQL_DB,
	 --(select count(*) from DMA_Assessment dm where dm.Instance_Name = dbinfo.InstanceName and dm.DBNAME=dbinfo.name and dm.TargetServer='ManagedSqlServer')  as Migrate_Managed_Instance
	from 
	(
	select  InstanceName,Name from [DMAReportingAxalta].[dbo].[ReportData] group by InstanceName,Name)dbinfo
	)FinalOutput
  "
  #$eqecutequery1= Invoke-SQLCMD $query1  -ServerInstance $SQLServer
  $execq1=Invoke-Sqlcmd -ServerInstance $Servername -Database master -Query  $query1 -EA "silentlycontinue" | Export-csv -path $D -NoTypeInformation
  #[void] [System.Windows.MessageBox]::Show( "Details and Summery Reports Generated Sucessfully on server $SQLServer", "Script completed", "OK", "Information" )
  #HTML Report generation

  $css = @"
<style>
h1, h5, th { text-align: center; font-family: Segoe UI; }
table { margin: auto; font-family: Segoe UI; box-shadow: 10px 10px 5px #888; border: thin ridge grey; }
th { background: #0046c3; color: #fff; max-width: 400px; padding: 5px 10px; }
td { font-size: 11px; padding: 5px 20px; color: #000; }
tr { background: #b8d1f3; }
tr:nth-child(even) { background: #dae5f4; }
tr:nth-child(odd) { background: #b8d1f3; }
</style>
"@

  $DMAReportfilepath=$DMA_path+"Reports"+"\"+"DMA_Summery_Report_$(get-date -f dd-MM-yyyy)"+".csv"
  $DMAReportfilepath
  $HTMLfilepath=$DMA_path +"Reports"+"\"+"DMA_Summery_Report_$(get-date -f dd-MM-yyyy)"+".HTML"
  #Remove-Item $HTMLfilepath
  Import-CSV $DMAReportfilepath | ConvertTo-Html -Head $css -Body "<h1>DMA Summary Report For All SQL Servers</h1>`n<h5>Generated on $(Get-Date)</h5>" | Out-File $HTMLfilepath
  Start-Sleep -s 2
  Invoke-Expression $HTMLfilepath
}
#All_Summary_Report
Function Main_Summary_Report
{

Function Summary_SingleServer_Report
{
  #Add-Type -AssemblyName Microsoft.SqlServer.Smo
  Import-Module SQlserver -DisableNameChecking
  $instancename = $textBoxIn.Text
  $D=$DMA_path+"Reports"+"\"+$instancename+"_DMA_Summery_Report_$(get-date -f dd-MM-yyyy)"+".csv"
  $D2=$DMA_path+"Reports"+"\"+$instancename+"_DMA_Detail_Report_$(get-date -f dd-MM-yyyy)"+".csv"
  #Remove-Item $D
  #Remove-Item $D2
  #$hostname= $env:COMPUTERNAME
  $TargetServerInstance = "sqlmydb001\SQL2023"#$env:COMPUTERNAME#"sqlvs001"
  $Servername = $TargetServerInstance

  $SQLServer = New-Object -TypeName Microsoft.SqlServer.Management.Smo.Server -ArgumentList $TargetServerInstance

  $query1="SELECT	rd.ServerName
    ,rd.DatabaseName
    ,rd.TargetCompatibilityLevel
    ,COALESCE(CASE changecategory WHEN 'BehaviorChange' THEN COUNT(*) END,0) AS 'BehaviorChange'
    ,COALESCE(CASE changecategory WHEN 'Deprecated' THEN COUNT(*) END,0) AS 'DeprecatedCount'
    --,COALESCE(CASE changecategory WHEN 'BreakingChange' THEN SUM(ChangeRank) END ,0) AS 'BreakingChange'
    ,COALESCE(CASE changecategory WHEN 'BreakingChange' THEN COUNT(*) END ,0) AS 'BreakingChange'
    ,COALESCE(CASE changecategory WHEN 'MigrationBlocker' THEN COUNT(*) END,0) AS 'MigrationBlocker'
    FROM	[DMA].dbo.ReportData rd
    --LEFT JOIN [DMA_reporting].dbo.BreakingChangeWeighting bcw
    --ON	rd.Title = bcw.Title
    WHERE	changecategory != 'NotDefined'
    AND TargetCompatibilityLevel != 'NA'
    AND TargetCompatibilityLevel = 'CompatLevel150'
    AND ServerName = '$instancename'
    --AND AssessmentTarget = 'AzureSQLDatabaseV12'
    GROUP BY ServerName, [DatabaseName], changecategory, TargetCompatibilityLevel
  "
  #$eqecutequery1= Invoke-SQLCMD $query1  -ServerInstance $SQLServer
  $execq1=Invoke-Sqlcmd -ServerInstance $Servername -Database master -Query  $query1 -EA "silentlycontinue" |Export-csv -path $D -NoTypeInformation
  #[void] [System.Windows.MessageBox]::Show( "Details and Summery Reports Generated Sucessfully on server $SQLServer", "Script completed", "OK", "Information" )
  #HTML Report generation

  $css = @"
<style>
h1, h5, th { text-align: center; font-family: Segoe UI; }
table { margin: auto; font-family: Segoe UI; box-shadow: 10px 10px 5px #888; border: thin ridge grey; }
th { background: #0046c3; color: #fff; max-width: 400px; padding: 5px 10px; }
td { font-size: 11px; padding: 5px 20px; color: #000; }
tr { background: #b8d1f3; }
tr:nth-child(even) { background: #dae5f4; }
tr:nth-child(odd) { background: #b8d1f3; }
</style>
"@

  #$DMAReportfilepath=$DMA_path+"Reports"+"\"+"DMA_Summery_Report_$(get-date -f dd-MM-yyyy)"+".csv"
  $DMAReportfilepath = $D
  $DMAReportfilepath
  #$HTMLfilepath=$DMA_path +"Reports"+"\"+"DMA_Summery_Report_$(get-date -f dd-MM-yyyy)"+".HTML"
     $HTMLfilepath  = $DMA_path+"Reports"+"\"+$instancename+"_DMA_Summery_Report_$(get-date -f dd-MM-yyyy)"+".HTML"
  #Remove-Item $HTMLfilepath
  Import-CSV $DMAReportfilepath | ConvertTo-Html -Head $css -Body "<h1>DMA Summary Report For $instancename Servers</h1>`n<h5>Generated on $(Get-Date)</h5>" | Out-File $HTMLfilepath
  Start-Sleep -s 2
  Invoke-Expression $HTMLfilepath

  ####################
  $query2=";WITH cte AS 
    (
    SELECT [ImportDate]
    ,[ServerName]
    ,[DatabaseName]
    ,[CurrentCompatibilityLevel]
    ,[ServerVersion]
    ,[TargetCompatibilityLevel]
    ,[Severity]
    ,[ChangeCategory]
    ,[Title]
    ,[Impact]
    ,[Recommendation]
    ,[MoreInfo]
    ,[ImpactedObjectName]
    ,[ImpactedObjectType]
    ,[ImpactDetail]
    , rn = ROW_NUMBER() OVER (PARTITION BY [ServerName], [DatabaseName], [ImpactedObjectName] 
      ORDER BY [ServerName], [DatabaseName], [ImpactedObjectName], [TargetCompatibilityLevel] DESC)
    FROM [DMA].[dbo].[ReportData]
    where [ServerName] = '$instancename'
    )
    SELECT * 
    FROM cte
    WHERE rn = 1 and ServerName ='$instancename'
    ORDER BY [ServerName], [DatabaseName], [ImpactedObjectName], [TargetCompatibilityLevel]; 
    GO
  "
  #$eqecutequery1= Invoke-SQLCMD $query1  -ServerInstance $SQLServer
  $execq2=Invoke-Sqlcmd -ServerInstance $Servername -Database DMA -Query  $query2 -EA "silentlycontinue" |Export-csv -path $D2 -NoTypeInformation

  $DMAReportfilepath2 = $D2
  $DMAReportfilepath2
  #$HTMLfilepath=$DMA_path +"Reports"+"\"+"DMA_Summery_Report_$(get-date -f dd-MM-yyyy)"+".HTML"
     $HTMLfilepath2  = $DMA_path+"Reports"+"\"+$instancename+"_DMA_Detail_Report_$(get-date -f dd-MM-yyyy)"+".HTML"
  #Remove-Item $HTMLfilepath2
  Import-CSV $DMAReportfilepath2 | ConvertTo-Html -Head $css -Body "<h1>DMA Detail Report For $instancename Servers</h1>`n<h5>Generated on $(Get-Date)</h5>" | Out-File $HTMLfilepath2
  Start-Sleep -s 2
  Invoke-Expression $HTMLfilepath2
  }
  ####################

  Add-Type -AssemblyName System.Windows.Forms
  Add-Type -AssemblyName System.Drawing
 
   
  # Main Form 
  $mainFormSI = New-Object System.Windows.Forms.Form
  $mainFormSI.Font = $header#"Comic Sans MS,8.25"
  $mainFormSI.Text = " DMA Single Instance Msgbox"
  $mainFormSI.FormBorderStyle = "FixedDialog"
  $mainFormSI.ForeColor = "white"
  $mainFormSI.BackColor = "Darkblue"
  $mainFormSI.StartPosition = "CenterParent"
  $mainFormSI.width = 600
  $mainFormSI.height = 250
 
  # Title Label
  $titleLabel = New-Object System.Windows.Forms.Label
  $titleLabel.Font = "Comic Sans MS,14"
  $titleLabel.ForeColor = "Yellow"
  $titleLabel.Location = "30,20"
  $titleLabel.Size = "400,30"
  $titleLabel.Text = "Enter SQL Server Instance name"
  $mainFormSI.Controls.Add($titleLabel);
  #$mainFormSI.Controls.Add($titleLabel)

  # Input Box
  $textBoxIn = New-Object System.Windows.Forms.TextBox
  $textBoxIn.Location = "35, 70"
  $textBoxIn.Size = "500, 20"
  $textBoxIn.Text = ""
  $mainFormSI.Controls.Add($textBoxIn)
 
  # Process Button
  $buttonProcess = New-Object System.Windows.Forms.Button
  $buttonProcess.Location = "35,150"
  $buttonProcess.Size = "75, 23"
  $buttonProcess.ForeColor = "Red"
  $buttonProcess.BackColor = "White"
  $buttonProcess.Text = "Process"
  #$buttonProcess = [System.Windows.MessageBox]::Show('Would  you like to play a game?','Game input','YesNoCancel','Error')
  #$buttonProcess.add_Click{processsingleServer}
  $buttonProcess.add_Click{Summary_SingleServer_Report}
  $mainFormSI.Controls.Add($buttonProcess)
 
  # Exit Button 
  $exitButton = New-Object System.Windows.Forms.Button
  $exitButton.Location = "450,150"
  $exitButton.Size = "75,23"
  $exitButton.ForeColor = "Red"
  $exitButton.BackColor = "White"
  $exitButton.Text = "Exit"
  $exitButton.add_Click{$mainFormSI.close()}
  $mainFormSI.Controls.Add($exitButton)
  #[void]$mainFormSI.ShowDialog()
  [void]$mainFormSI.ShowDialog()

  }

  Function Message
{
  Add-Type -AssemblyName PresentationFramework
  [void] [System.Windows.MessageBox]::Show( "Coming Soon.... ", "Script completed", "OK", "Information" )
}

Function introduction_myDB
{
Add-Type -AssemblyName System.Speech
$synthesizer = New-Object -TypeName System.Speech.Synthesis.SpeechSynthesizer
$synthesizer.Rate = -2
$synthesizer.Speak('Hey, Welcome to myDB Assessment Tool')
$synthesizer.Speak('Today I will present you some introduction about myDB Assessment Tool for sql server database assessment')
 $synthesizer.Speak('myDB Assessment Tool enables you to upgrade a modern data platform by detecting compatibility issues, 
 that can impact database functionality on your new version of SQL Server. 
 It recommends performance and reliability improvements for your target environment. 
 It allows you to uncontained objects from your source server to your target server.
 Basically, it gives you a list of issues to tackle before you can go ahead with your migration. 
 Not all of these issues are blocking, most might be warnings and would not prevent you to upgrade to SQL Server two zero one nine. 
 After myDB has finished the assessment you will get an overview of all the issues found per compatibility. 
 You also have the option to export this report,
 which is good to know, myDB does save the results for later itself. 
 Once finished we have three type of file saved in directory.
 These three file format are JSON,csv and last is dma file.
 Thank you ')

}

Function DTU_calculate
 {
 Add-Type -AssemblyName System.Windows.Forms
  $onclick_buttoncombo =
  {
    Function Get-AzureSQLDBDTU
{
  [CmdletBinding()]
  PARAM (
    [Parameter(Mandatory = $true,HelpMessage = 'Please enter the number of processor cores')][int]$Core,
    [Parameter(Mandatory = $true,HelpMessage = 'Please enter the apiPerformanceItems input parameters in JSON format')][String]$apiPerformanceItems
  )
  $DUTCalculatorUri = "http://dtucalculator.azurewebsites.net/api/calculate?cores=$cores"
  try
  {
    $Response = Invoke-WebRequest -UseBasicParsing -Headers @{
      'Content-Type' = 'application/json'
    } -Uri $DUTCalculatorUri -Body $apiPerformanceItems -Method Post
  }
  catch 
  {
    throw 'failed to invoke the REST API.'
    Exit -1
  }


  If ($Response.StatusCode -eq 200)
  {
    Write-Verbose 'Response received.'
    $ResponseContent = ConvertFrom-Json -InputObject $Response.Content
  }
  $ResponseContent
}
#region variables
$SampleInterval = 1
$MaxSamples = 300 #3600
$DatabaseName = $the_comboazvm.Text
$LogicalDriveLetter = 'C'
$ComputerName = $the_comboaz.Text
#endregion

#Get number of processors cores
$processors = get-wmiobject -query 'select * from win32_processor' -ComputerName $ComputerName
$Cores = 0
Foreach ($processor in $Processors)
{
  $Cores = $numberOfCores + $processor.NumberOfCores
}

#region collect perf counters
$counters = @('\Processor(_Total)\% Processor Time', "\LogicalDisk($LogicalDriveLetter`:)\Disk Reads/sec", "\LogicalDisk($LogicalDriveLetter`:)\Disk Writes/sec", "\SQLServer:Databases($DatabaseName)\Log Bytes Flushed/sec")

$arrRawPerfValues = Get-Counter -Counter $counters -SampleInterval $SampleInterval -MaxSamples $MaxSamples -ComputerName $ComputerName
$arrPerfValues = @()
Foreach ($item in $arrRawPerfValues)
{
  $processorTime =$item.CounterSamples[0].CookedValue
  $diskReads = $item.CounterSamples[1].CookedValue
  $diskWrites = $item.CounterSamples[2].CookedValue
  $logBytesFlushed = $item.CounterSamples[3].CookedValue
  $properties = @{
  diskReads       = $diskReads
  diskWrites      = $diskWrites
  logBytesFlushed = $logBytesFlushed
  processorTime   = $processorTime
  }
  $objPerf = New-Object -TypeName psobject -Property $properties
  $arrPerfValues += $objPerf
}
#endregion
$Dir_Path ="C:\myDB_Assessment_Report\"
$DMA_Path =$Dir_Path
#region Calculate DTU
#construct web API JSON parameter
$apiPerformanceItems = ConvertTo-Json -InputObject $arrPerfValues

#Invoke the web API to calculate DUT
$DTUCalculationResult = Get-AzureSQLDBDTU -Core $Cores -apiPerformanceItems $apiPerformanceItems
$DatabaseName
$D=$DMA_path+"Reports"+"\"+"SKU_Recommendation_in_DTU_$(get-date -f dd-MM-yyyy)"+".csv"
Remove-Item $D
$DTUCalculationResult.Recommendations |Export-csv -path $D -NoTypeInformation
#$DTUCalculationResult.SelectedServiceTiers
$css = @"
<style>
h1, h5, th { text-align: center; font-family: Segoe UI; }
table { margin: auto; font-family: Segoe UI; box-shadow: 10px 10px 5px #888; border: thin ridge grey; }
th { background: #0046c3; color: #fff; max-width: 400px; padding: 5px 10px; }
td { font-size: 11px; padding: 5px 20px; color: #000; }
tr { background: #b8d1f3; }
tr:nth-child(even) { background: #dae5f4; }
tr:nth-child(odd) { background: #b8d1f3; }
</style>
"@
#$DMA_path="C:\myDB_Assessment_Report\"
#$DMAReportfilepath=$DMA_path+"Reports"+"\"+"SKU_Recommendation_in_DTU_$(get-date -f dd-MM-yyyy)"+".csv"
#$DMAReportfilepath
$HTMLfilepath=$DMA_path +"Reports"+"\"+"SKU_Recommendation_in_DTU_$(get-date -f dd-MM-yyyy)"+".HTML"
$HTMLfilepath
Import-CSV $D | ConvertTo-Html -Head $css -Body "<h1>SKU Recommendation in DTU Report of $DatabaseName</h1>`n<h5>Generated on $(Get-Date)</h5>" | Out-File $HTMLfilepath
Start-Sleep -s 2
Invoke-Expression $HTMLfilepath
#endregion
[void] [System.Windows.MessageBox]::Show( "DTU Report has been generated Sucessfully  ", "Script completed", "OK", "Information" )   

  }

  
  
  $mainFormHTML = New-Object System.Windows.Forms.Form
  $mainFormHTML.Font = $header#"Comic Sans MS,8.25"
  $mainFormHTML.Text = " DTU Reports msgbox"
  $mainFormHTML.FormBorderStyle = "FixedDialog"
  $mainFormHTML.ForeColor = "White"
  $mainFormHTML.BackColor = "DarkBlue"
  $mainFormHTML.StartPosition = "CenterParent"
  $mainFormHTML.width = 600
  $mainFormHTML.height = 250
 
  # Title Label
  $titleLabel = New-Object System.Windows.Forms.Label
  $titleLabel.Font = $header#"Comic Sans MS,14"
  $titleLabel.ForeColor = "DarkBlue"
  $titleLabel.Location = "30,20"
  $titleLabel.Size = "400,30"
  $titleLabel.Text = "Select Instance name"
  $mainFormHTML.Controls.Add($titleLabel);
  #$mainFormHTML.Controls.Add($titleLabel)

  # Input Box
  $textBoxInaz = New-Object System.Windows.Forms.TextBox
  $textBoxInaz.Location = "35, 70"
  $textBoxInaz.Size = "500, 20"
  $textBoxInaz.Text = ""
  #$mainFormHTML.Controls.Add($textBoxInaz)
 
  #Combobox
  $the_comboaz = New-Object system.Windows.Forms.ComboBox
  $the_comboaz.location = "35, 60"

  $the_comboaz.Size = "200, 20"

  $the_comboaz.DropDownStyle = "Dropdownlist"
  #$ComboList_Items = Get-Content $DMA_Path"DMA_DB_Type.txt"
  $get_serverlist = $Dir_Path+"Inventory"+'\'+"MS_Instancelist.txt"
  $ComboList_Items = Get-Content $get_serverlist #@("Azuresqldatabase", "ManagedSqlServer" ,"AzureSqlVirtualMachine")

  #Loop thru the text file or the array
  #and add the contents to the combobox for selection
  ForEach ($Server in $ComboList_Items) {

    $the_comboaz.Items.Add($Server)


  }

  $mainFormHTML.controls.add($the_comboaz)
  # New combobox 
  $the_comboazvm = New-Object system.Windows.Forms.ComboBox
  $the_comboazvm.location = "35, 100"

  $the_comboazvm.Size = "200, 20"

  $the_comboazvm.DropDownStyle = "Dropdownlist"
  #$ComboList_Items = Get-Content $DMA_Path"DMA_DB_Type.txt"
  $srvname = $the_comboaz.Text
  $servername = New-Object Microsoft.SqlServer.Management.Smo.Server -ArgumentList $srvname

          $S=$servername.Name
          $ComboList_Items2 = $servername.Databases | Select-Object name

  #$ComboList_Items2 = @("Azuresqldatabase", "ManagedSqlServer" ,"AzureSqlVirtualMachine")# @("Sqlserver2012", "Sqlserver2014" ,"Sqlserver2016","SqlServerWindows2017","SqlServerLinux2017","SqlServerWindows2019","SqlServerLinux2019")

  #Loop thru the text file or the array
  #and add the contents to the combobox for selection
  ForEach ($dbname in $ComboList_Items2.name) {

    $the_comboazvm.Items.Add($dbname)
    
    }

    $event_handler = 
  {
    #$the_combo.Items.Clear()
    $targetcheck=$the_comboaz.Text
    if($targetcheck -ne "") #AzureSqlVirtualMachine
    {
      $the_comboaz.Enabled=$false
      $mainFormHTML.controls.add($the_comboazvm)
       
    }
    else
    {
      $the_comboaz.Enabled=$True
     
    }
  }
   
    $the_comboaz.add_SelectedIndexChanged($event_handler)

    #$the_comboaz.SelectedIndex=0

    # Process Button
  $buttonProcess = New-Object System.Windows.Forms.Button
  $buttonProcess.Location = "35,170"
  $buttonProcess.Size = "75, 25"
  $buttonProcess.ForeColor = "Red"
  $buttonProcess.BackColor = "White"
  $buttonProcess.Font = $header
  $buttonProcess.Text = "View"
  $buttonProcess.add_Click($onclick_buttoncombo)#Progresswindowgui_DTU
  #$buttonProcess.add_Click({Progresswindowgui_DTU -Target 'DTU'})
  $mainFormHTML.Controls.Add($buttonProcess)
 
  # Exit Button 
  $exitButton = New-Object System.Windows.Forms.Button
  $exitButton.Location = "450,170"
  $exitButton.Size = "75,25"
  $exitButton.ForeColor = "Red"
  $exitButton.BackColor = "White"
  $exitButton.Font = $header
  $exitButton.Text = "Exit"
  $exitButton.add_Click{$mainFormHTML.close()}
  $mainFormHTML.Controls.Add($exitButton)

  # Reset Button 
  $ResetButton = New-Object System.Windows.Forms.Button
  $ResetButton.Location = "200,170"
  $ResetButton.Size = "85,25"
  $ResetButton.ForeColor = "Red"
  $ResetButton.BackColor = "White"
  $ResetButton.Font = $header
  $ResetButton.Text = "Reset"
  $ResetButton.add_Click{Reset_button}
  $mainFormHTML.Controls.Add($ResetButton)

  # Audio Help Button 
  $AhelpButton = New-Object System.Windows.Forms.Button
  $AhelpButton.Location = "300,170"
  $AhelpButton.Size = "75,23"
  $AhelpButton.ForeColor = "Red"
  $AhelpButton.BackColor = "White"
  $AhelpButton.Text = "Help"
  $AhelpButton.add_Click{Audio_help}
  $mainFormHTML.Controls.Add($AhelpButton)
  

    [void]$mainFormHTML.ShowDialog()
 
}

Function Speech_Exp
{
Add-Type -AssemblyName System.Speech
$synthesizer = New-Object -TypeName System.Speech.Synthesis.SpeechSynthesizer
$synthesizer.Speak('Hey, Welcome to myDB Assessment Tool')
$synthesizer.Speak('Good to see you')
}

<#
Function Progresswindowgui_DTU
{

  param (
        [Parameter(Mandatory=$true)][string]$Target
    )

    Add-Type -AssemblyName System.Drawing
  Add-Type -AssemblyName System.Windows.Forms
  $main_form            = New-Object System.Windows.Forms.Form
  $main_form.Text           ='Reports Progressbar'
  $main_form.foreColor      ='white'
  $main_form.BackColor      ='Darkblue'
  $main_form.Font           = $header
  $main_form.Width          = 600
  $main_form.Height         = 250

  $header                   = New-Object System.Drawing.Font("Verdana",13,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
  $procFont                 = New-Object System.Drawing.Font("Verdana",20,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

  $Label                    = New-Object System.Windows.Forms.Label
  $Label.Font               = $header
  $Label.ForeColor          ='yellow'
  $Label.Text               = "Are you sure want to continue......"
  $Label.Location           = New-Object System.Drawing.Point(10,10)
  $Label.Width              = 480
  $Label.Height             = 50

  $StartButton              = New-Object System.Windows.Forms.Button
  $StartButton.Location     = New-Object System.Drawing.Size(350,75)
  $StartButton.Size         = New-Object System.Drawing.Size(120,50)
  $StartButton.Text         = "Start"
  $StartButton.height       = 40
  $StartButton.BackColor    ='white'
  $StartButton.ForeColor    ='red'
  $StartButton.Add_click({Progressbar_DTU -Target $Target})

  $EndButton              = New-Object System.Windows.Forms.Button
  $EndButton.Location     = New-Object System.Drawing.Size(350,75)
  $EndButton.Size         = New-Object System.Drawing.Size(120,50)
  $EndButton.Text         = "OK"
  $EndButton.height       = 40
  $EndButton.BackColor    ='white'
  $EndButton.ForeColor    ='blue'
  #$EndButton.add_Click{$main_form.close()}
  
  $EndButton.DialogResult = [System.Windows.Forms.DialogResult]::OK

  $main_form.Controls.AddRange(($Label,$StartButton,$EndButton))

  $main_form.StartPosition = "manual"
  $main_form.Location = New-Object System.Drawing.Size(500, 300)
  $result=$main_form.ShowDialog() 
  $Target=$null


}

Function Progressbar_DTU
{

  param (
        [Parameter(Mandatory=$true)][string]$Target
    )
  Add-Type -AssemblyName System.Windows.Forms
  Add-Type -AssemblyName PresentationFramework
  Add-Type -AssemblyName System.Drawing
  
  [System.Windows.Forms.Application]::EnableVisualStyles()
  $ProgressBar              = New-Object System.Windows.Forms.ProgressBar
  $ProgressBar.Location     = New-Object System.Drawing.Point(10,35)
  $ProgressBar.Size         = New-Object System.Drawing.Size(460,40)
  $ProgressBar.Style        = "Marquee"
  $ProgressBar.MarqueeAnimationSpeed = 20
  $main_form.Controls.Add($ProgressBar)

  $Label.Font             = $header
  $Label.ForeColor        ='yellow'
  $Label.Text             ="Processing ..."
  #$ProgressBar.visible
  
  If ($Target -eq 'DTU') 
  {
    $ProgressBar.visible
    #DMA-Single-Instance 
    DTU_Calculate
    $Label.Text               = "Process Complete"
    $ProgressBar.Hide()
    $StartButton.Hide()
    $EndButton.Visible
  } 
}



#>
# Main Window Form Setup SkuRecommendationReport-20221219
################################################################## Objects
# Main Form .Net Objects
$mainForm         = New-Object System.Windows.Forms.Form
$menuMain         = New-Object System.Windows.Forms.MenuStrip
$menuDMA          = New-Object System.Windows.Forms.ToolStripMenuItem
$menuSKU         = New-Object System.Windows.Forms.ToolStripMenuItem
#$menuTools        = New-Object System.Windows.Forms.ToolStripMenuItem
#$menuOpen         = New-Object System.Windows.Forms.ToolStripMenuItem
$menuazuredbDMA   = New-Object System.Windows.Forms.ToolStripMenuItem
$menuMIDMA        = New-Object System.Windows.Forms.ToolStripMenuItem
$menuSave         = New-Object System.Windows.Forms.ToolStripMenuItem
$menuSaveAs       = New-Object System.Windows.Forms.ToolStripMenuItem
$menuSqlserver    = New-Object System.Windows.Forms.ToolStripMenuItem
$menuAzureVM      = New-Object System.Windows.Forms.ToolStripMenuItem
$menusql2012vm    = New-Object System.Windows.Forms.ToolStripMenuItem
$menusql2014vm    = New-Object System.Windows.Forms.ToolStripMenuItem
$menuFullScr      = New-Object System.Windows.Forms.ToolStripMenuItem
$menuazDBSKU      = New-Object System.Windows.Forms.ToolStripMenuItem
$menuazmiSKU      = New-Object System.Windows.Forms.ToolStripMenuItem
$menuazvmSKU      = New-Object System.Windows.Forms.ToolStripMenuItem
#$menuOptions      = New-Object System.Windows.Forms.ToolStripMenuItem
#$menuOptions1     = New-Object System.Windows.Forms.ToolStripMenuItem
#$menuOptions2     = New-Object System.Windows.Forms.ToolStripMenuItem
#$menusql2012      = New-Object System.Windows.Forms.ToolStripMenuItem
#$menusql2014      = New-Object System.Windows.Forms.ToolStripMenuItem
#$menusql2019LIN   = New-Object System.Windows.Forms.ToolStripMenuItem
#$menusql2019      = New-Object System.Windows.Forms.ToolStripMenuItem
#$menusql2017LIN   = New-Object System.Windows.Forms.ToolStripMenuItem
#$menusql2017     = New-Object System.Windows.Forms.ToolStripMenuItem
#$menusql2016     = New-Object System.Windows.Forms.ToolStripMenuItem
#$menusql2019LINvm   = New-Object System.Windows.Forms.ToolStripMenuItem
#$menusql2019vm     = New-Object System.Windows.Forms.ToolStripMenuItem
#$menusql2017LINvm   = New-Object System.Windows.Forms.ToolStripMenuItem
#$menusql2017vm     = New-Object System.Windows.Forms.ToolStripMenuItem
#$menusql2016vm     = New-Object System.Windows.Forms.ToolStripMenuItem
$menuExit         = New-Object System.Windows.Forms.ToolStripMenuItem
$menusins         = New-Object System.Windows.Forms.ToolStripMenuItem
$menuExitDM       = New-Object System.Windows.Forms.ToolStripMenuItem

$menuHelp         = New-Object System.Windows.Forms.ToolStripMenuItem
$menuAbout        = New-Object System.Windows.Forms.ToolStripMenuItem
$mainToolStrip    = New-Object System.Windows.Forms.ToolStrip
#$toolStripOpen    = New-Object System.Windows.Forms.ToolStripButton
#$toolStripSave    = New-Object System.Windows.Forms.ToolStripButton
#$toolStripSaveAs  = New-Object System.Windows.Forms.ToolStripButton
#$toolStripFullScr = New-Object System.Windows.Forms.ToolStripButton
#$toolStripAbout   = New-Object System.Windows.Forms.ToolStripButton
#$toolStripExit    = New-Object System.Windows.Forms.ToolStripButton
$statusStrip      = New-Object System.Windows.Forms.StatusStrip
$statusLabel      = New-Object System.Windows.Forms.ToolStripStatusLabel
$menuSQLServer    = New-Object System.Windows.Forms.ToolStripMenuItem
$menuPaaSAss      = New-Object System.Windows.Forms.ToolStripMenuItem
$menuIaaSAss      = New-Object System.Windows.Forms.ToolStripMenuItem
$menuPaaSSingleIns = New-Object System.Windows.Forms.ToolStripMenuItem
$menuPaaSAllIns  = New-Object System.Windows.Forms.ToolStripMenuItem
$menuIaaSSigIns = New-Object System.Windows.Forms.ToolStripMenuItem
$menuIaaSAllIns = New-Object System.Windows.Forms.ToolStripMenuItem
$menuOracle = New-Object System.Windows.Forms.ToolStripMenuItem
$menuIaaSSg = New-Object System.Windows.Forms.ToolStripMenuItem
$menuIaaSAll = New-Object System.Windows.Forms.ToolStripMenuItem
$menureferdoc = New-Object System.Windows.Forms.ToolStripMenuItem
$menurazdbdoc = New-Object System.Windows.Forms.ToolStripMenuItem

$menuIaaSAllIaas = New-Object System.Windows.Forms.ToolStripMenuItem
$menuIaaSSgIaas = New-Object System.Windows.Forms.ToolStripMenuItem
$menurazmidoc = New-Object System.Windows.Forms.ToolStripMenuItem
$menurazmidoc2 = New-Object System.Windows.Forms.ToolStripMenuItem
$menuver = New-Object System.Windows.Forms.ToolStripMenuItem
$menuAWS = New-Object System.Windows.Forms.ToolStripMenuItem
#$menuGCP = New-Object System.Windows.Forms.ToolStripMenuItem

$menurazrinvm=New-Object System.Windows.Forms.ToolStripMenuItem

$menurazlog=New-Object System.Windows.Forms.ToolStripMenuItem

$menurazlogsucess =New-Object System.Windows.Forms.ToolStripMenuItem
$menurazlogfail =New-Object System.Windows.Forms.ToolStripMenuItem

$menurazGR=New-Object System.Windows.Forms.ToolStripMenuItem

$menurazFM=New-Object System.Windows.Forms.ToolStripMenuItem
$menurazMath=New-Object System.Windows.Forms.ToolStripMenuItem

$menurazrinskuins = New-Object System.Windows.Forms.ToolStripMenuItem

$menurjsontosql = New-Object System.Windows.Forms.ToolStripMenuItem

$menurazGRSall = New-Object System.Windows.Forms.ToolStripMenuItem
$menurazGRsingle = New-Object System.Windows.Forms.ToolStripMenuItem

################################################################## Icons
#Add-Type -TypeDefinition $code -ReferencedAssemblies System.Drawing

# Extract PowerShell Icon from PowerShell Exe
$iconPS   = [Drawing.Icon]::ExtractAssociatedIcon((Get-Command powershell).Path)

# background. This is where I need help. 
 $Image = [system.drawing.image]::FromFile("C:\Axalta_Assessment_Report\image\myDB_Assessment23.jpg") 
 #$Image = [system.drawing.image]::FromFile("C:\myDB_Assessment_Report\image\once22.png")

#this location is where my question arises. It won't work on another user's machine. 


 $mainForm.BackgroundImage = $Image
 
 $header = New-Object System.Drawing.Font("Verdana", 10, [System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
 $procFont = New-Object System.Drawing.Font("Verdana", 20, [System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

################################################################## Main Form Setup
# Main Form
$mainForm.Height          = 500
$mainForm.Icon            = $iconPS
$mainForm.MainMenuStrip   = $menuMain
$mainForm.Width           = 1000
$mainForm.StartPosition   = "CenterScreen"
#$mainForm.ForeColor = "White"
$mainForm.BackColor = "DarkBlue"
$mainForm.Text            = "myDBAssessment"
#$mainForm.Font = "Comic Sans MS,14"
$mainForm.Font = $header
$mainForm.ForeColor = "DarkBlue"
$mainForm.BackgroundImage = $Image
$mainForm.BackgroundImageLayout = "stretch"
#$mainForm.BackgroundImageLayout = "center"
$mainForm.AutoScale = $true
$mainForm.Controls.Add($menuMain)


################################################################## Main Menu

# Main ToolStrip
[void]$mainForm.Controls.Add($mainToolStrip)

# Main Menu Bar
[void]$mainForm.Controls.Add($menuMain)

#$menuSQLServer   =  New-Object System.Windows.Forms.ToolStripMenuItem
#$menuPaaSAss   =  New-Object System.Windows.Forms.ToolStripMenuItem
#$menuIaaSAss = New-Object System.Windows.Forms.ToolStripMenuItem

# Menu Options - SQL Server
$menuSQLServer.Text = "&Azure"
$menuSQLServer.Font = $header#"Comic Sans MS,14"
$menuSQLServer.ForeColor = "DarkBlue"
#$menuDMA.Font = Arial
$menuSQLServer.Add_Click{$menuOracle.Enabled = $False}
[void]$menuMain.Items.Add($menuSQLServer)

# Menu Options - File
$menuDMA.Text = "&DMA Report"
$menuDMA.Font = $header
$menuDMA.ForeColor = "DarkBlue"
#$menuDMA.Font = Arial
[void]$menuSQLServer.DropDownItems.Add($menuDMA)



# Menu Options - SQL Server / PaaS
#$menuazuredbDMA.Image        = [System.IconExtractor]::Extract("shell32.dll", 4, $true)
#$menuazuredbDMA.ShortcutKeys = "Control, O"
$menuPaaSAss.Text         = "&Target PaaS Assessment"
$menuPaaSAss.ForeColor = "DarkBlue"
#$menuPaaSAss.Add_Click{AzureDBDMA -Target 'Azuresqldatabase'}
[void]$menuDMA.DropDownItems.Add($menuPaaSAss)

# Menu Options - SQL Server / PaaS/ All
#$menusql2012vm.Image     = [System.IconExtractor]::Extract("shell32.dll", 33, $true)
$menuPaaSSingleIns.Text      = "&All SQL Server"
$menuPaaSSingleIns.ForeColor = "DarkBlue"
#$menuPaaSSingleIns.Add_Click{AllSQLServer}
#$menuPaaSSingleIns.Add_Click{AzureDBDMA}
$menuPaaSSingleIns.Add_Click{DMA-WindowSubForm-PaaS -DMA_Opt 'DMA_All_PaaS' -DMAReports 'DMA_ForAllPaaS'}
#$menuPaaSSingleIns.Add_Click{Speech_Exp}
[void]$menuPaaSAss.DropDownItems.Add($menuPaaSSingleIns)

# Menu Options - SQL Server / PaaS/ Single Instance
#$menusql2012vm.Image     = [System.IconExtractor]::Extract("shell32.dll", 33, $true)Speech_Exp
$menuPaaSAllIns.Text      = "&Single SQL Instance"
$menuPaaSAllIns.ForeColor = "DarkBlue"
#$menuPaaSAllIns.Add_Click{DMA-WindowSubForm-PaaS -DMA_Opt 'DMA_Single_PaaS' -DMAReports 'DMA-Single-Instance'}
$menuPaaSAllIns.Add_Click{DMA-WindowSubForm-PaaS -DMA_Opt 'DMA_Single_PaaS' -DMAReports 'DMA-Single-Instance'}
#$menuPaaSAllIns.Add_Click{Progresswindowgui -Target 'DMA-Single-Instance'}
[void]$menuPaaSAss.DropDownItems.Add($menuPaaSAllIns)

# Menu Options - Sql Server / IaaS Assessment
#$menuMIDMA.Image        = [System.IconExtractor]::Extract("shell32.dll", 36, $true)
#$menuMIDMA.ShortcutKeys = "F2"
$menuIaaSAss.Text         = "&Target IaaS Assessment"
$menuIaaSAss.ForeColor = "DarkBlue"
#$menuIaaSAss.Add_Click{AzureDBDMA -Target 'ManagedSqlServer'}
[void]$menuDMA.DropDownItems.Add($menuIaaSAss)

# Menu Options - SQL Server / PaaS/ Single Instance
#$menusql2012vm.Image     = [System.IconExtractor]::Extract("shell32.dll", 33, $true)
$menuIaaSAllIns.Text      = "&Target SQL Server (IaaS)"
$menuIaaSAllIns.ForeColor = "DarkBlue"
#$menuIaaSAllIns.Add_Click{AzureDBDMA -Target 'Sqlserver2012'}
[void]$menuIaaSAss.DropDownItems.Add($menuIaaSAllIns)

# Menu Options - SQL Server / PaaS/ Single Instance
#$menusql2012vm.Image     = [System.IconExtractor]::Extract("shell32.dll", 33, $true)
$menuIaaSAllIaas.Text      = "&All SQL Server"
$menuIaaSAllIaas.ForeColor = "DarkBlue"
#$menuIaaSAllIaas.Add_Click{AllSQLServerIaaS}
$menuIaaSAllIaas.Add_Click{DMA-WindowSubForm-PaaS -DMA_Opt 'DMA_All_IaaS' -DMAReports 'DMA_All_IaaS'}
[void]$menuIaaSAllIns.DropDownItems.Add($menuIaaSAllIaas)


# Menu Options - SQL Server / PaaS/ Single Instance
#$menusql2012vm.Image     = [System.IconExtractor]::Extract("shell32.dll", 33, $true)
$menuIaaSSgIaas.Text      = "&Single SQL Instance"
$menuIaaSSgIaas.ForeColor = "DarkBlue"
#$menuIaaSSgIaas.Add_Click{Single-InstanceIaaS}
$menuIaaSSgIaas.Add_Click{DMA-WindowSubForm-PaaS -DMA_Opt 'DMA_Single_IaaS' -DMAReports 'DMA-Single-Instance'}
[void]$menuIaaSAllIns.DropDownItems.Add($menuIaaSSgIaas)


# Menu Options - SQL Server / PaaS/ Single Instance
#$menusql2012vm.Image     = [System.IconExtractor]::Extract("shell32.dll", 33, $true)
$menuIaaSSigIns.Text      = "&Target SQL Azure VM"
$menuIaaSSigIns.ForeColor = "DarkBlue"
#$menuIaaSSigIns.Add_Click{Single-InstanceIaaS}
[void]$menuIaaSAss.DropDownItems.Add($menuIaaSSigIns)

# Menu Options - SQL Server / PaaS/ Single Instance
#$menusql2012vm.Image     = [System.IconExtractor]::Extract("shell32.dll", 33, $true)
$menuIaaSAll.Text      = "&All SQL Server"
$menuIaaSAll.ForeColor = "DarkBlue"
$menuIaaSAll.Add_Click{DMA-WindowSubForm-PaaS -DMA_Opt 'DMA_All_IaaS' -DMAReports 'DMA_All_IaaS'}
[void]$menuIaaSSigIns.DropDownItems.Add($menuIaaSAll)



# Menu Options - SQL Server / PaaS/ Single Instance
#$menusql2012vm.Image     = [System.IconExtractor]::Extract("shell32.dll", 33, $true)
$menuIaaSSg.Text      = "&Single SQL Instance"
$menuIaaSSg.ForeColor = "DarkBlue"
$menuIaaSSg.Add_Click{DMA-WindowSubForm-PaaS -DMA_Opt 'DMA_Single_IaaS' -DMAReports 'DMA-Single-Instance'}
[void]$menuIaaSSigIns.DropDownItems.Add($menuIaaSSg)


# Menu Options - SKU
$menuSKU.Text      = "&SKU Report"
$menuSKU.Font = $header#"Comic Sans MS,14"
$menuSKU.ForeColor = "DarkBlue"
[void]$menuSQLServer.DropDownItems.Add($menuSKU)

# Menu Options - View / Full Screen
#$menuazDBSKU.Image        = [System.IconExtractor]::Extract("shell32.dll",34, $true)
#$menuazDBSKUr.ShortcutKeys = "Control, F"
$menuazDBSKU.Text         = "&SKU_Assessment"
$menuazDBSKU.ForeColor = "DarkBlue"
#$menuazDBSKU.Add_Click{SKUAssessment}
$menuazDBSKU.Add_Click{DMA-WindowSubForm-PaaS -DMA_Opt 'SKUAssessment' -DMAReports 'SKUAssfunction'}
[void]$menuSKU.DropDownItems.Add($menuazDBSKU)

# Menu Options - SKU / ManagedInstance
#$menuazmiSKU.Image        = [System.IconExtractor]::Extract("shell32.dll",34, $true)
#$menuazmiSKU.ShortcutKeys = "Control, F"
$menuazmiSKU.Text         = "&SKU_Report_Generation"
$menuazmiSKU.ForeColor = "DarkBlue"
$menuazmiSKU.Add_Click{DMA-WindowSubForm-PaaS -DMA_Opt 'SKUGenReport' -DMAReports 'SKUReportitem'}
[void]$menuSKU.DropDownItems.Add($menuazmiSKU)


<#
# Menu Options - View / Full Screen
#$menuazDBSKU.Image        = [System.IconExtractor]::Extract("shell32.dll",34, $true)
#$menuazDBSKUr.ShortcutKeys = "Control, F"
$menuazDBSKU.Text         = "&Target Azure_SQLDB  (PaaS)"
$menuazDBSKU.ForeColor = "DarkBlue"
$menuazDBSKU.Add_Click{SKU -Target 'AzureSqlDatabase'}
[void]$menuSKU.DropDownItems.Add($menuazDBSKU)

# Menu Options - SKU / ManagedInstance
#$menuazmiSKU.Image        = [System.IconExtractor]::Extract("shell32.dll",34, $true)
#$menuazmiSKU.ShortcutKeys = "Control, F"
$menuazmiSKU.Text         = "&Target Azure_SQLManagedInstance  (PaaS)"
$menuazmiSKU.ForeColor = "DarkBlue"
$menuazmiSKU.Add_Click{SKU -Target 'AzureSqlManagedInstance'}
[void]$menuSKU.DropDownItems.Add($menuazmiSKU)

# Menu Options - SKU / AzureVM
#$menuazvmSKU.Image        = [System.IconExtractor]::Extract("shell32.dll",34, $true)
#$menuazvmSKU.ShortcutKeys = "Control, F"
$menuazvmSKU.Text         = "&Target Azure_SQLVM  (IaaS)"
$menuazvmSKU.ForeColor = "DarkBlue"
$menuazvmSKU.Add_Click{SKU -Target 'AzureSqlVirtualMachine'}
[void]$menuSKU.DropDownItems.Add($menuazvmSKU)
#>

# Menu Options - File / Exit
#$menuExit.Image        = [System.IconExtractor]::Extract("shell32.dll", 10, $true)
#$menuExit.ShortcutKeys = "Control, X"
$menuExitDM.Text         = "&Exit"
$menuExitDM.ForeColor = "DarkBlue"
$menuExitDM.Add_Click{$mainForm.Close()}
[void]$menuSQLServer.DropDownItems.Add($menuExitDM)


# Menu Options - AWS
$menuAWS.Text = "&AWS"
$menuAWS.Font = $header#"Comic Sans MS,14"
$menuAWS.ForeColor = "DarkBlue"
#$menuDMA.Font = Arial
$menuAWS.Add_Click{message}
[void]$menuMain.Items.Add($menuAWS)

<#
    # Menu Options - GCP
    $menuGCP.Text = "&GCP"
    $menuGCP.Font = $header#"Comic Sans MS,14"
    $menuGCP.ForeColor = "Blue"
    #$menuDMA.Font = Arial
    $menuGCP.Add_Click{Message}
    [void]$menuMain.Items.Add($menuGCP)
#>

# Menu Options - Oracle
$menuOracle.Text = "&Oracle/postgresql"
$menuOracle.Font = $header#"Comic Sans MS,14"
$menuOracle.ForeColor = "DarkBlue"
#$menuOracle.Enabled = $True
#$menuOracle.On_Click{$menureferdoc.Enable = $false}
#$menuDMA.Font = Arial
$menuOracle.Add_Click{Start-Process C:\myDB_Assessment_Report\Oracle\putty.exe}
$menuOracle.Add_Click{$mainForm.Close()}
[void]$menuMain.Items.Add($menuOracle)



# Menu Options - Oracle
$menureferdoc.Text = "&Files"
$menureferdoc.Font = $header#"Comic Sans MS,14"
$menureferdoc.ForeColor = "DarkBlue"
#$menuDMA.Font = Arial
[void]$menuMain.Items.Add($menureferdoc)


# Menu Options - Oracle
$menurazdbdoc.Text = "&DMA Reports"
$menurazdbdoc.Font = $header#"Comic Sans MS,14"
$menurazdbdoc.ForeColor = "DarkBlue"
$menurazdbdoc.Add_Click{DMA_files_view -Filetype 'DMAfile'}
#$menurazdbdoc.Add_Click{DMA-WindowSubForm-PaaS -DMA_Opt 'DMA_Fileserver'}
#$menuDMA.Font = Arial
[void]$menureferdoc.DropDownItems.Add($menurazdbdoc)

# Menu Options - reference
$menurazmidoc.Text = "&SKU Reports"
$menurazmidoc.Font = $header#"Comic Sans MS,14"
$menurazmidoc.ForeColor = "DarkBlue"
$menurazmidoc.Add_Click{DMA_files_view -Filetype 'sku_counters'}
#$menuDMA.Font = Arial
[void]$menureferdoc.DropDownItems.Add($menurazmidoc)

# Menu Options - reference
$menurazmidoc2.Text = "&SQLServer_Inventory"
$menurazmidoc2.Font = $header#"Comic Sans MS,14"
$menurazmidoc2.ForeColor = "DarkBlue"
$menurazmidoc2.Add_Click{Inventory -itype 'Tool'}
#$menuDMA.Font = Arial
[void]$menureferdoc.DropDownItems.Add($menurazmidoc2)

# Menu Options - Inventory Managment
$menurazrinvm.Text = "&Project_Inventory"
$menurazrinvm.Font = $header#"Comic Sans MS,14"
$menurazrinvm.ForeColor = "DarkBlue"
$menurazrinvm.Add_Click{Inventory -itype 'Project'}
#$menuDMA.Font = Arial
[void]$menureferdoc.DropDownItems.Add($menurazrinvm)

# Menu Options - Inventory Managment
$menurazrinskuins.Text = "&SKU_Instances_list"
$menurazrinskuins.Font = $header#"Comic Sans MS,14"
$menurazrinskuins.ForeColor = "DarkBlue"
$menurazrinskuins.Add_Click{Inventory -itype 'SKU_Instances'}
#$menuDMA.Font = Arial
[void]$menureferdoc.DropDownItems.Add($menurazrinskuins)

# Menu Options - JSON to SQL
$menurjsontosql.Text = "&Import JSON to SQL"
$menurjsontosql.Font = $header#"Comic Sans MS,14"
$menurjsontosql.ForeColor = "DarkBlue"
$menurjsontosql.Add_Click{Import_JSON_to_SQLDB}
#$menuDMA.Font = Arial
[void]$menureferdoc.DropDownItems.Add($menurjsontosql)

# Menu Options - Logs
$menurazlog.Text = "&Logs"
$menurazlog.Font = $header#"Comic Sans MS,14"
$menurazlog.ForeColor = "DarkBlue"
#$menurazlog.Add_Click{Logs}
#$menuDMA.Font = Arial
[void]$menureferdoc.DropDownItems.Add($menurazlog)

# Menu Options - Logs/Sucess
$menurazlogsucess.Text = "&Success_list"
$menurazlogsucess.Font = $header#"Comic Sans MS,14"
$menurazlogsucess.ForeColor = "DarkBlue"
$menurazlogsucess.Add_Click{Inventory -itype 'success'}
#$menuDMA.Font = Arial
[void]$menurazlog.DropDownItems.Add($menurazlogsucess)

# Menu Options - Logs/Failure
$menurazlogfail.Text = "&Failure_list"
$menurazlogfail.Font = $header#"Comic Sans MS,14"
$menurazlogfail.ForeColor = "DarkBlue"
$menurazlogfail.Add_Click{Inventory -itype 'failure'}
#$menuDMA.Font = Arial
[void]$menurazlog.DropDownItems.Add($menurazlogfail)

# Menu Options - Generate DMA report from SQL table
$menurazGR.Text = "&Generate Report(SQL Data)"
$menurazGR.Font = $header#"Comic Sans MS,14"
$menurazGR.ForeColor = "DarkBlue"
#$menurazGR.Add_Click{DMA_Report}
#$menurazGR.Add_Click{HTML_File_DMA_Report}
#$menurazGR.Add_Click{Summary_Report}
#$menuDMA.Font = Arial
[void]$menureferdoc.DropDownItems.Add($menurazGR)

# Menu Options - Generate DMA report from SQL table
$menurazGRSall.Text = "&Summary for All"
$menurazGRSall.Font = $header#"Comic Sans MS,14"
$menurazGRSall.ForeColor = "DarkBlue"
$menurazGRSall.Add_Click{DMA_Report_All}
#$menurazGRSall.Add_Click{HTML_File_DMA_Report_All}
$menurazGRSall.Add_Click{All_Summary_Report}
#$menuDMA.Font = Arial
[void]$menurazGR.DropDownItems.Add($menurazGRSall)


# Menu Options - Generate DMA report from SQL table
$menurazGRsingle.Text = "&Summary_Single_Server"
$menurazGRsingle.Font = $header#"Comic Sans MS,14"
$menurazGRsingle.ForeColor = "DarkBlue"
#$menurazGRsingle.Add_Click{DMA_Report}
#$menurazGRsingle.Add_Click{HTML_File_DMA_Report}
$menurazGRsingle.Add_Click{Main_Summary_Report}
#$menuDMA.Font = Arial
[void]$menurazGR.DropDownItems.Add($menurazGRsingle)


# Menu Options - Framework for Migration
$menurazFM.Text = "&DTU"
$menurazFM.Font = $header#"Comic Sans MS,14"
$menurazFM.ForeColor = "DarkBlue"
$menurazFM.Add_Click{DTU_Calculate}#DMA-WindowSubForm-PaaS -DMAReports 'DTU'
#$menurazFM.Add_Click{Progresswindowgui_DTU -Target 'DTU'}
#$menurazGR.Add_Click{HTML_File_DMA_Report}
#$menurazFM.Add_Click{Start-Process ((Resolve-Path "C:\myDB_Assessment_Report\Reports\Database Migrations Capabilities  SQL Server and MySQL.pdf").Path)}
#$menuDMA.Font = Arial
[void]$menureferdoc.DropDownItems.Add($menurazFM)
<#
# Menu Options - Mathology
$menurazMath.Text = "&mathology"
$menurazMath.Font = $header#"Comic Sans MS,14"
$menurazMath.ForeColor = "DarkBlue"
#$menurazGR.Add_Click{DMA_Report}
#$menurazGR.Add_Click{HTML_File_DMA_Report}
$menurazMath.Add_Click{message}
#$menuDMA.Font = Arial
[void]$menureferdoc.DropDownItems.Add($menurazMath)
#>
# Menu Options - Version
$menuver.Text      = "&Version"
$menuver.Font = $header#"Comic Sans MS,14"
$menuver.ForeColor = "DarkBlue"
[void]$menuMain.Items.Add($menuver)

# Menu Options - Help / About
$menuAbout.Image     = [System.Drawing.SystemIcons]::Information
$menuAbout.Text      = "About myDB Assessment"
$menuAbout.ForeColor = "DarkBlue"
$menuAbout.Add_Click{introduction_myDB}
[void]$menuver.DropDownItems.Add($menuAbout)

# Menu Options - Help
$menuHelp.Text      = "&Help"
$menuHelp.Font = $header#"Comic Sans MS,14"
$menuHelp.ForeColor = "DarkBlue"
$menuHelp.Add_Click{message}
#$menuHelp.Add_Click{Start-Process ((Resolve-Path "C:\myDB_Assessment_Report\Reports\myDB_Assessment.pdf").Path)}
[void]$menuMain.Items.Add($menuHelp)



################################################################## ToolBar Buttons
################################################################## Functions

#####################################

    #[void]$mainForm.Close()

 # End About

# Show Main Form
$mainForm.add_Shown({Directory_Creation} )
[void] $mainForm.ShowDialog()