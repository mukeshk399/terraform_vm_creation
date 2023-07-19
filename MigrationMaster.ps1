<#
    ############################################
    #  Created by-----Mukesh Kumar             #
    #                                          #
    # Emailid--------kum-mukesh@hcl.com        #
    #                                          #
    # Program Name---Migration Master           #
    #                                          #
    # Version--------V1.0                      #
    #                                          #
    # Date-----------22-Sept-2022              #
    ############################################
#>
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName system.drawing
#Import-Module SQLPS -DisableNameChecking
Import-Module -name sqlserver -DisableNameChecking
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMO') | out-null
#[System.Windows.Form.Application]::EnableVisualStyles()
$Dir_Path ="U:\Solution_Master\"
$DMA_Path =$Dir_Path

#$test = $menupaassubazdb.Text.Replace('&','')
#$test

Function MigrationMaster-WindowSubForm
{
  param (
        [Parameter(Mandatory=$true)][string]$Sol_Opt,
        [Parameter(Mandatory=$true)][string]$SolutionMasterReports
    )
  Add-Type -AssemblyName System.Windows.Forms
  Add-Type -AssemblyName System.Drawing

  write-host "This is options of $Sol_Opt"
  write-host "This is selection solution of $SolutionMasterReports"
  # Main Form 
  $mainFormSI = New-Object System.Windows.Forms.Form
  $mainFormSI.Font = $header#"Comic Sans MS,8.25"
  $mainFormSI.Text = " Solution Master Msgbox"
  $mainFormSI.FormBorderStyle = "FixedDialog"
  $mainFormSI.ForeColor = "white"
  $mainFormSI.BackColor = "Darkblue"
  $mainFormSI.StartPosition = "CenterParent"
  $mainFormSI.width = 500
  $mainFormSI.height = 300
 
  # Title Label
  $titleLabel = New-Object System.Windows.Forms.Label
  $titleLabel.Font = "Comic Sans MS,14"
  $titleLabel.ForeColor = "Yellow"
  $titleLabel.Location = "30,20"
  $titleLabel.Size = "400,30"
  $titleLabel.Text = "Enter Source Server Name "
  #$mainFormSI.Controls.Add($titleLabel);
  #$mainFormSI.Controls.Add($titleLabel)

   #$password = New-Object Windows.Forms.MaskedTextBox
   #$password.PasswordChar = '*'
   #$password.Top  = 100
   #$password.Left = 80

    # Title Label
  $titleLabeIM = New-Object System.Windows.Forms.Label
  $titleLabeIM.Font = "Comic Sans MS,14"
  $titleLabeIM.ForeColor = "Yellow"
  $titleLabeIM.Location = "30,20"
  $titleLabeIM.Size = "400,30"
  $titleLabeIM.Text = "Enter Target Server Name "
  #$mainFormSI.Controls.Add($titleLabeIM);

  
  # Input Box
  $textBoxIn1 = New-Object System.Windows.Forms.TextBox
  $textBoxIn1.Location = "35, 55"
  $textBoxIn1.Size = "300, 20"
  $textBoxIn1.Text = ""
  #$textBoxIn1.PasswordChar = "*"
  $mainFormSI.Controls.Add($textBoxIn1)

  # Title Label
  $titleLabe2 = New-Object System.Windows.Forms.Label
  $titleLabe2.Font = "Comic Sans MS,14"
  $titleLabe2.ForeColor = "Yellow"
  $titleLabe2.Location = "35,80"
  $titleLabe2.Size = "300,30"
  $titleLabe2.Text = "Enter Source Database Name "
  #$mainFormSI.Controls.Add($titleLabe2);
  #$mainFormSI.Controls.Add($titleLabel)

  # Title Label
  $titleLabeIM2 = New-Object System.Windows.Forms.Label
  $titleLabeIM2.Font = "Comic Sans MS,14"
  $titleLabeIM2.ForeColor = "Yellow"
  $titleLabeIM2.Location = "35,80"
  $titleLabeIM2.Size = "300,30"
  $titleLabeIM2.Text = "Enter Target Database Name "
  #$mainFormSI.Controls.Add($titleLabeIM2);

  # Input Box
  $textBoxIn2 = New-Object System.Windows.Forms.TextBox
  $textBoxIn2.Location = "35, 120"
  $textBoxIn2.Size = "300, 20"
  $textBoxIn2.Text = ""
  $mainFormSI.Controls.Add($textBoxIn2)

  # Title Label
  $titleLabe3 = New-Object System.Windows.Forms.Label
  $titleLabe3.Font = "Comic Sans MS,14"
  $titleLabe3.ForeColor = "Yellow"
  $titleLabe3.Location = "35,150"
  $titleLabe3.Size = "300,30"
  $titleLabe3.Text = "Select Action "
  $mainFormSI.Controls.Add($titleLabe3);
  #$mainFormSI.Controls.Add($titleLabel)

  #Combobox
  $the_combo = New-Object system.Windows.Forms.ComboBox
  $the_combo.location = "35, 180"

  $the_combo.Size = "300, 20"

  $the_combo.DropDownStyle = "Dropdownlist"
  #$ComboList_Items = Get-Content $DMA_Path"DMA_DB_Type.txt"
  If ($Sol_Opt -eq 'Single DB')
 {
  
  $ComboList_Items = @("Export", "Import" ,"Extract")
  #$Sol_Opt=$null
  
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
    write-host "$targetcheck is the selection of solution" -ForegroundColor Yellow
    if ( $targetcheck -eq 'Export')
    {
     #$textBoxIn1.Enabled = $false
     #$textBoxIn1.Visible = $false
     #$mainFormSI.Controls.Add($titleLabe2);
     $mainFormSI.Controls.Add($titleLabe2)
     $mainFormSI.Controls.Add($titleLabel)
     $titleLabeIM2.Visible = $false
     $titleLabeIM.Visible = $false
    }
    if ( $targetcheck -eq 'Import')
    {
     $mainFormSI.Controls.Add($titleLabeIM2)
     $mainFormSI.Controls.Add($titleLabeIM)
     $titleLabe2.Visible = $false
     $titleLabel.Visible = $false
    }
  }

  $the_combo.add_SelectedIndexChanged($the_combo_SelectedIndexChanged)

  # Process Button
  $buttonProcess = New-Object System.Windows.Forms.Button
  $buttonProcess.Location = "35,220"
  $buttonProcess.Size = "75, 23"
  $buttonProcess.ForeColor = "Red"
  $buttonProcess.BackColor = "White"
  $buttonProcess.Text = "Process"
  #$buttonProcess.add_Click{processsingleServer} $SolutionMasterReports
  $buttonProcess.add_Click{SQLPackageexe}
  #$buttonProcess.add_Click{$SolutionMasterReports}
  $mainFormSI.Controls.Add($buttonProcess)
 
  # Exit Button 
  $exitButton = New-Object System.Windows.Forms.Button
  $exitButton.Location = "150,220"
  $exitButton.Size = "75,23"
  $exitButton.ForeColor = "Red"
  $exitButton.BackColor = "White"
  $exitButton.Text = "Exit"
  $exitButton.add_Click{$mainFormSI.close()}
  $mainFormSI.Controls.Add($exitButton)
  #[void]$mainFormSI.ShowDialog()
  [void]$mainFormSI.ShowDialog()

  
}

Function SQLPackageexe
{
 #[System.Management.Automation.PSCredential]$SqlCredential

    #define variable
     #$varmain="C:\Program Files\Microsoft SQL Server\160\DAC\bin\SqlPackage.exe" 
            #$var = "/Action:Export /ssn:tcp:sqlvs001,1433 /sdn:AdventureWorks2016 /su:azureuser /sp:Maxwell@12345 /tf:U:\backup\test.bacpac /p:Storage=File"
            #$var1= $var.Split(" ")
            #& $varmain $var1
  if ($the_combo.text -eq 'Export')
   {
   write-host "$the_combo.text is running" -ForegroundColor Yellow
  $servername = $textBoxIn1.text #'sqlvs001'
  $servername
  $dbname = $textBoxIn2.text #'AdventureWorks2016'
  $dbname
  $action = $the_combo.text
  $user = 'sa'
  $Password = 'Maxwell@12345'
  $targetpath = 'U:\backup\'+$dbname+'.bacpac'
             $arglist = @(
	
              #'PerfDataCollection --sqlConnectionStrings "Data Source='+$S+';Initial Catalog=master;Integrated Security=True;" --outputFolder '+$path+''
              '/Action:'+$action+' /ssn:tcp:'+$servername+',1433 /sdn:'+$dbname+' /su:'+$user+' /sp:'+$Password+' /tf:'+$targetpath+' /p:Storage=File'
            )
           $arglist
            Start-Process -FilePath 'C:\Program Files\Microsoft SQL Server\160\DAC\bin\SqlPackage.exe'  -ArgumentList $arglist #-NoNewWindow
            Start-Sleep 10
            write-host "$the_combo.text completed" -ForegroundColor Yellow
            [void] [System.Windows.MessageBox]::Show( "$action databases $dbname successfully completed ", "Script completed", "OK", "Information" )
 }

 if ($the_combo.text -eq 'Import')
 {
      $targetaction = 'Import'
      $targetserver = $textBoxIn1.text
      $targetdatabase = $textBoxIn2.text
      $targetuser = 'azureadmin'
      $targetPassword = 'Maxwell@12345'
      $DatabaseEdition = 'Standard'
      $DatabaseService = 'S3'
      $targetpathimp = 'U:\backup\'+$targetdatabase+'.bacpac'
      $arglist = @(
	
              #'PerfDataCollection --sqlConnectionStrings "Data Source='+$S+';Initial Catalog=master;Integrated Security=True;" --outputFolder '+$path+''
              '/Action:'+$targetaction+' /tsn:tcp:'+$targetserver+',1433 /tdn:'+$targetdatabase+' /tu:'+$targetuser+' /tp:'+$targetPassword+' /sf:'+$targetpathimp+' /p:DatabaseEdition='+$DatabaseEdition+' /p:DatabaseServiceObjective='+$DatabaseService+' /p:Storage=File /DF:U:\backup\logaz' 
            )

            Start-Process -FilePath 'C:\Program Files\Microsoft SQL Server\160\DAC\bin\SqlPackage.exe'  -ArgumentList $arglist #-NoNewWindow
            Start-Sleep 10
            [void] [System.Windows.MessageBox]::Show( "$targetaction databases $targetdatabase successfully completed ", "Script completed", "OK", "Information" )
 }
 
}
#SQLPackageexe

function About {
    Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName System.Windows.Forms
  $statusLabel.Text = "About"
    # About Form Objects
    $aboutForm          = New-Object System.Windows.Forms.Form
    $aboutFormExit      = New-Object System.Windows.Forms.Button
    $aboutFormImage     = New-Object System.Windows.Forms.PictureBox
    $aboutFormNameLabel = New-Object System.Windows.Forms.Label
    $aboutFormText      = New-Object System.Windows.Forms.Label

    # About Form
    $aboutForm.AcceptButton  = $aboutFormExit
    $aboutForm.CancelButton  = $aboutFormExit
    $aboutForm.ClientSize    = "350, 110"
    $aboutForm.ControlBox    = $false
    $aboutForm.ShowInTaskBar = $false
    $aboutForm.StartPosition = "CenterParent"
    $aboutForm.Text          = "Migration Master"
    $aboutForm.Add_Load($aboutForm_Load)

    # About PictureBox
    $aboutFormImage.Image    = $iconPS.ToBitmap()
    $aboutFormImage.Location = "55, 15"
    $aboutFormImage.Size     = "32, 32"
    $aboutFormImage.SizeMode = "StretchImage"
    $aboutForm.Controls.Add($aboutFormImage)

    # About Name Label
    $aboutFormNameLabel.Font     = New-Object Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
    $aboutFormNameLabel.Location = "110, 20"
    $aboutFormNameLabel.Size     = "300, 18"
    $aboutFormNameLabel.Text     = "Migration Master"
    $aboutForm.Controls.Add($aboutFormNameLabel)

    # About Text Label
    $aboutFormText.Location = "100, 40"
    $aboutFormText.Size     = "300, 30"
    $aboutFormText.Text     = "  Version 1.0"
    $aboutForm.Controls.Add($aboutFormText)

    # About Exit Button
    $aboutFormExit.Location = "135, 70"
    $aboutFormExit.Text     = "OK"
    $aboutForm.Controls.Add($aboutFormExit)

    [void]$aboutForm.ShowDialog()
    $statusLabel.Text = "Ready"
}


Function Importcsvfile
{
$csvimport = Import-Csv C:\Users\azureuser\Downloads\employeeDB.CSV #-delimiter "," |
#$csvimport
$i=0
ForEach ( $f in $csvimport)
{

    $SourceAction = $f.SourceAction
    $Sourceserver = $f.Sourceserver
    $SourceDatabase = $f.SourceDatabase
    $Sourceuser = $f.Sourceuser
    $SourcePassword = $f.SourcePassword
    $TargetFile = $f.TargetFile
    $TargetAction = $f.TargetAction
    $TargetServer =$f.TargetServer
    $TargetDatabase = $f.TargetDatabase
    $TargetUser = $f.TargetUser
    $TargetPassword   = $f.TargetPassword
    $DatabaseEdition  = $f.DatabaseEdition
    $DatabaseService  = $f.DatabaseService

   $backupname = $TargetFile+$SourceDatabase+'.bacpac'
    $backupname

    $Sourceserver
    $SourceDatabase
     $TargetAction
     $DatabaseService

     if ($sourceAction -eq 'Export')
     {
       write-host "$sourceaction to sourece path action required" -ForegroundColor Yellow
       $arglist = @(
	
              #'PerfDataCollection --sqlConnectionStrings "Data Source='+$S+';Initial Catalog=master;Integrated Security=True;" --outputFolder '+$path+''
              '/Action:'+$sourceaction+' /ssn:tcp:'+$sourceserver+',1433 /sdn:'+$sourcedatabase+' /su:'+$sourceuser+' /sp:'+$sourcePassword+' /tf:'+$backupname+' /p:Storage=File'
            )
            #write-host "$arglist"
            #Start-Process -FilePath 'C:\Program Files\Microsoft SQL Server\160\DAC\bin\SqlPackage.exe'  -ArgumentList $arglist -NoNewWindow
            #Start-Sleep 30

             $varmain="C:\Program Files\Microsoft SQL Server\160\DAC\bin\SqlPackage.exe" 
            $var = "/Action:$sourceaction /ssn:tcp:$sourceserver,1433 /sdn:$sourcedatabase /su:$sourceuser /sp:$sourcePassword /tf:$backupname /p:Storage=File"
            #$var = $arglist
            $var1= $var.Split(" ")
            write-host "Export value is-------$var1"
            & $varmain $var1

            start-sleep 10
      #($TargetAction -eq 'Import')
     
       #write-host "Import to azure sql action required" -ForegroundColor Magenta
        $arglist1 = @(
	
              #'PerfDataCollection --sqlConnectionStrings "Data Source='+$S+';Initial Catalog=master;Integrated Security=True;" --outputFolder '+$path+''
              '/Action:'+$targetaction+' /tsn:tcp:'+$targetserver+',1433 /tdn:'+$targetdatabase+' /tu:'+$targetuser+' /tp:'+$targetPassword+' /sf:'+$backupname+' /p:DatabaseEdition='+$DatabaseEdition+' /p:DatabaseServiceObjective='+$DatabaseService+' /p:Storage=File /DF:U:\backup\logaz' 
            )
            #$arglist1
            #Start-Process -FilePath 'C:\Program Files\Microsoft SQL Server\160\DAC\bin\SqlPackage.exe'  -ArgumentList $arglist1 -NoNewWindow
            #Start-Sleep 10

            $varmain1="C:\Program Files\Microsoft SQL Server\160\DAC\bin\SqlPackage.exe" 
            $var1 = "/Action:$targetaction /tsn:tcp:$targetserver,1433 /tdn:$targetdatabase /tu:$targetuser /tp:$targetPassword /sf:$backupname /p:DatabaseEdition=$DatabaseEdition /p:DatabaseServiceObjective=$DatabaseService /p:Storage=File /DF:U:\backup\logaz"
            #$var = $arglist
            $var2= $var1.Split(" ")
            write-host "Import value is------------$var2"
            & $varmain1 $var2
            write-host "completed database $targetdatabase...." -ForegroundColor Magenta
            #[void] [System.Windows.MessageBox]::Show( "Import database $targetdatabase successfully completed ", "Script completed", "OK", "Information" )

       #sqlpackage.exe /Action:Import /tsn:tcp:server-1557355658.database.windows.net,1433 /tdn:devdb /tu:sqladmin /tp:Maxwell@12345 /sf:C:\mydb\test.bacpac /p:DatabaseEdition=Standard /p:DatabaseServiceObjective=S3 /p:Storage=File /DF:C:\mydb\logaz
     }

    $i+=$i

}
 [void] [System.Windows.MessageBox]::Show( "Import databases $targetdatabase successfully completed ", "Script completed", "OK", "Information" )
}
#Importcsvfile


# WinForm Setup
################################################################## Objects
# Main Form .Net Objects
$mainForm         = New-Object System.Windows.Forms.Form
$menuMain         = New-Object System.Windows.Forms.MenuStrip

$mainToolStrip    = New-Object System.Windows.Forms.ToolStrip
$toolStripOpen    = New-Object System.Windows.Forms.ToolStripButton
$toolStripSave    = New-Object System.Windows.Forms.ToolStripButton
$toolStripSaveAs  = New-Object System.Windows.Forms.ToolStripButton
$toolStripFullScr = New-Object System.Windows.Forms.ToolStripButton
$toolStripAbout   = New-Object System.Windows.Forms.ToolStripButton
$toolStripExit    = New-Object System.Windows.Forms.ToolStripButton
$statusStrip      = New-Object System.Windows.Forms.StatusStrip
$statusLabel      = New-Object System.Windows.Forms.ToolStripStatusLabel

# Migration master menu
$menuSqlserver    = New-Object System.Windows.Forms.ToolStripMenuItem
$menuPredefineSolms = New-Object System.Windows.Forms.ToolStripMenuItem
$menums = New-Object System.Windows.Forms.ToolStripMenuItem
$menuiaasaz = New-Object System.Windows.Forms.ToolStripMenuItem
$menuiaaspaas = New-Object System.Windows.Forms.ToolStripMenuItem
$menuver = New-Object System.Windows.Forms.ToolStripMenuItem
$menuAbout = New-Object System.Windows.Forms.ToolStripMenuItem
$menuHelp = New-Object System.Windows.Forms.ToolStripMenuItem
$menupaassubazdb = New-Object System.Windows.Forms.ToolStripMenuItem
$menupaassubazmi = New-Object System.Windows.Forms.ToolStripMenuItem
$menupaassdb = New-Object System.Windows.Forms.ToolStripMenuItem
$menupaasmdb = New-Object System.Windows.Forms.ToolStripMenuItem

################################################################## Icons
# WinForms Icons
# Create Icon Extractor Assembly
$code = @"
using System;
using System.Drawing;
using System.Runtime.InteropServices;

namespace System
{
	public class IconExtractor
	{

	 public static Icon Extract(string file, int number, bool largeIcon)
	 {
	  IntPtr large;
	  IntPtr small;
	  ExtractIconEx(file, number, out large, out small, 1);
	  try
	  {
	   return Icon.FromHandle(largeIcon ? large : small);
	  }
	  catch
	  {
	   return null;
	  }

	 }
	 [DllImport("Shell32.dll", EntryPoint = "ExtractIconExW", CharSet = CharSet.Unicode, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
	 private static extern int ExtractIconEx(string sFile, int iIndex, out IntPtr piLargeVersion, out IntPtr piSmallVersion, int amountIcons);

	}
}
"@
Add-Type -TypeDefinition $code -ReferencedAssemblies System.Drawing

# Extract PowerShell Icon from PowerShell Exe
$iconPS   = [Drawing.Icon]::ExtractAssociatedIcon((Get-Command powershell).Path)

# background. This is where I need help. 
 $Image = [system.drawing.image]::FromFile("U:\solution_master\image\Migration_master.jpg") 
 #$Image = [system.drawing.image]::FromFile("C:\myDB_Assessment_Report\image\Migration_master.jpg")

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
$mainForm.Text            = "solution Master"
#$mainForm.Font = "Comic Sans MS,14"
$mainForm.Font = $header
$mainForm.ForeColor = "DarkBlue"
$mainForm.BackgroundImage = $Image
$mainForm.BackgroundImageLayout = "stretch"
#$mainForm.BackgroundImageLayout = "center"
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
#$menuSQLServer.Add_Click{$menuOracle.Enabled = $False}
[void]$menuMain.Items.Add($menuSQLServer)

# Menu Options - Microsoft SQL Server
$menums.Text = "&IaaS Migration"
$menums.Font = $header
$menums.ForeColor = "DarkBlue"
#$menuDMA.Font = Arial
#$menuDMA.Add_Click{SolutionMaster-WindowSubForm -Sol_Opt 'Azure_Databases' -SolutionMasterReports 'AzureSQLDatabases_Solutions' }
#$menums.Add_Click{SolutionMaster-WindowSubForm -Sol_Opt $menums.Text.Replace('&','') -SolutionMasterReports $menuSQLServer.Text.Replace('&','') }
[void]$menuSQLServer.DropDownItems.Add($menums)


# Menu Options - Microsoft\IaaS
$menuiaasaz.Text = "&PaaS Migration"
$menuiaasaz.Font = $header#"Comic Sans MS,14"
$menuiaasaz.ForeColor = "DarkBlue"
#$menuDMA.Add_Click{SolutionMaster-WindowSubForm -Sol_Opt $menuDMA.Text.Replace('&','') -SolutionMasterReports 'Framework' }
#$menuiaasaz.Add_Click{SolutionMaster-WindowSubForm -Sol_Opt $menuSKU.Text.Replace('&','') -SolutionMasterReports $menuiaasaz.Text.Replace('&','') }
[void]$menuSQLServer.DropDownItems.Add($menuiaasaz)



# Menu Options - Microsoft\PaaS sub menu
$menupaassubazdb.Text = "&Azure SQLDB"
$menupaassubazdb.Font = $header#"Comic Sans MS,14"
$menupaassubazdb.ForeColor = "DarkBlue"
#$menuDMA.Add_Click{SolutionMaster-WindowSubForm -Sol_Opt $menuDMA.Text.Replace('&','') -SolutionMasterReports 'Framework' }
#$menupaassubazdb.Add_Click{MigrationMaster-WindowSubForm -Sol_Opt $menupaassubazdb.Text.Replace('&','') -SolutionMasterReports $menuiaasaz.Text.Replace('&','') }
[void]$menuiaasaz.DropDownItems.Add($menupaassubazdb)


# Menu Options - Microsoft\PaaS sub menu
$menupaassdb.Text = "&Single DB"
$menupaassdb.Font = $header#"Comic Sans MS,14"
$menupaassdb.ForeColor = "DarkBlue"
#$menuDMA.Add_Click{SolutionMaster-WindowSubForm -Sol_Opt $menuDMA.Text.Replace('&','') -SolutionMasterReports 'Framework' }
$menupaassdb.Add_Click{MigrationMaster-WindowSubForm -Sol_Opt $menupaassdb.Text.Replace('&','') -SolutionMasterReports $menupaassubazdb.Text.Replace('&','') }
[void]$menupaassubazdb.DropDownItems.Add($menupaassdb)

# Menu Options - Microsoft\PaaS sub menu
$menupaasmdb.Text = "&Multiple DB"
$menupaasmdb.Font = $header#"Comic Sans MS,14"
$menupaasmdb.ForeColor = "DarkBlue"
#$menuDMA.Add_Click{SolutionMaster-WindowSubForm -Sol_Opt $menuDMA.Text.Replace('&','') -SolutionMasterReports 'Framework' }
$menupaasmdb.Add_Click{Importcsvfile}
[void]$menupaassubazdb.DropDownItems.Add($menupaasmdb)


# Menu Options - Microsoft\PaaS sub menu
$menupaassubazmi.Text      = "&Azure SQL MI"
$menupaassubazmi.Font = $header#"Comic Sans MS,14"
$menupaassubazmi.ForeColor = "DarkBlue"
#$menuDMA.Add_Click{SolutionMaster-WindowSubForm -Sol_Opt $menuDMA.Text.Replace('&','') -SolutionMasterReports 'Framework' }
#$menuiaasaz.Add_Click{SolutionMaster-WindowSubForm -Sol_Opt $menuSKU.Text.Replace('&','') -SolutionMasterReports $menuiaasaz.Text.Replace('&','') }
[void]$menuiaasaz.DropDownItems.Add($menupaassubazmi)


# Menu Options -  Microsoft\IaaS
$menuiaaspaas.Text      = "&OnPrim Migration"
$menuiaaspaas.Font = $header#"Comic Sans MS,14"
$menuiaaspaas.ForeColor = "DarkBlue"
#$menuiaaspaas.Add_Click{SolutionMaster-WindowSubForm -Sol_Opt $menumsoth.Text.Replace('&','') -SolutionMasterReports $menuSQLServer.Text.Replace('&','') }
[void]$menuSQLServer.DropDownItems.Add($menuiaaspaas)



# Menu Options - Version
$menuver =New-Object System.Windows.Forms.ToolStripMenuItem
$menuver.Text      = "&Version"
$menuver.Font = $header#"Comic Sans MS,14"
$menuver.ForeColor = "DarkBlue"
[void]$menuMain.Items.Add($menuver)

# Menu Options - Help / About
$menuAbout =New-Object System.Windows.Forms.ToolStripMenuItem
$menuAbout.Image     = [System.Drawing.SystemIcons]::Information
$menuAbout.Text      = "About Migration Master"
$menuAbout.ForeColor = "DarkBlue"
$menuAbout.Add_Click{About}
#$menuAbout.Add_Click{Start-Process ((Resolve-Path "C:\myDB_Assessment_Report\Image\myDBAssessment.pdf").Path)}
[void]$menuver.DropDownItems.Add($menuAbout)

# Menu Options - Help
$menuHelp.Text      = "&Exit"
$menuHelp.Font = $header#"Comic Sans MS,14"
$menuHelp.ForeColor = "DarkBlue"
$menuHelp.Add_Click{$mainForm.Close()}
[void]$menuMain.Items.Add($menuHelp)





################################################################## ToolBar Buttons
################################################################## Functions

#####################################

    #[void]$mainForm.Close()

 # End About

# Show Main Form
#$mainForm.add_Shown({Directory_Creation} )
[void] $mainForm.ShowDialog()