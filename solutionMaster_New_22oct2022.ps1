<#
    ############################################
    #  Created by-----Mukesh Kumar             #
    #                                          #
    # Emailid--------kum-mukesh@hcl.com        #
    #                                          #
    # Program Name---Solution Master           #
    #                                          #
    # Version--------V1.0                      #
    #                                          #
    # Date-----------05-Sept-2022              #
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
# Install .Net Assemblies
Function Directory_Creation
{
  Add-Type -AssemblyName PresentationFramework
  #$Dir_Path ="C:\Solution_Master"+'\'
       
        if(!(Test-Path -Path $Dir_Path)) 
        {
         <#
          #New-Item "$InstallDir" -type directory -Force | out-null
            new-item -type directory -path $Dir_Path -Force
            new-item -type directory -path $Dir_Path"DMA" -Force
            new-item -type directory -path $Dir_Path"SKU" -Force
            new-item -type directory -path $Dir_Path"Oracle" -Force
            new-item -type directory -path $Dir_Path"SingleInstance" -Force
            new-item -type directory -path $Dir_Path"SingleInstance_IaaS" -Force
            new-item -type directory -path $Dir_Path"Reports" -Force
            new-item -type file -path $Dir_Path"MS_Instancelist.txt" -Force
            new-item -type file -path $Dir_Path"server_not_connect.txt" -Force
            new-item -type file -path $Dir_Path"server_connect.txt" -Force
            new-item -type file -path $Dir_Path"serverlist.txt" -Force
            new-item -type file -path $Dir_Path"inventory_list.txt" -Force
            new-item -type file -path $Dir_Path"single_Instance.txt" -Force
            new-item -type file -path $Dir_Path"single_Instance_IaaS.txt" -Force
            new-item -type file -path $Dir_Path"Assess-for-AzureSQLMI.xml" -Force
            
            Start-Sleep 5
            write-host "Content adding on XML file " -ForegroundColor Yellow
            #Add-Content $Dir_Path"MS_Instancelist.txt" $sname
            start-sleep 3
            Add-Content $Dir_Path"Assess-for-AzureSQLMI.xml" -Value '<?xml version="1.0" encoding="UTF-8"?>
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
          write-host "Directory $Dir_Path and file created sucessfully " -ForegroundColor Yellow
          [void] [System.Windows.MessageBox]::Show( "Directory $Dir_Path Created,Please enter Server name ", "Script completed", "OK", "Information" )
          start-sleep 2
          #exit;
          #>
          #write-host "Folder $Dir_Path not exits.First create directory" -ForegroundColor Green
          [void] [System.Windows.MessageBox]::Show( "Folder $Dir_Path not exits.Please check File share ", "Script completed", "OK", "Information" )
        }
        else 
        {
          write-host "Folder $Dir_Path already exists" -ForegroundColor Green
          #[void] [System.Windows.MessageBox]::Show( "Directory $Dir_Path Already Created,Please check Servers on serverlist.txt before run DMA Report ", "Script completed", "OK", "Information" )
          

        }

}
Function Reset_button
{
  $the_comboaz.Enabled=$True
  $the_comboazvm.Visible=$False
  $textBoxInaz.Text =""
  #$the_comboaz.Items.Clear()
  $the_comboaz.Text=""
  #$the_comboaz.Text=""

}
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
    $aboutForm.Text          = "Solution Master"
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
    $aboutFormNameLabel.Text     = "Solution Master"
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



Function Message
{
  Add-Type -AssemblyName PresentationFramework
  [void] [System.Windows.MessageBox]::Show( "Coming Soon.... ", "Script completed", "OK", "Information" )
}


#Add-Type -AssemblyName System.Windows.Forms
#Add-Type -AssemblyName System.Drawing
#[Windows.Forms.Application]::EnableVisualStyles()



 Function SolutionMaster-WindowSubForm
{
  param (
        [Parameter(Mandatory=$true)][string]$mmenuopt,
        [Parameter(Mandatory=$true)][string]$smenuopt1,
        [Parameter(Mandatory=$true)][string]$smenuopt2
    )
    #$menurmssql.Add_Click{SolutionMaster-WindowSubForm -mmenuopt $menuSQLServer.Text.Replace('&','') -smenuopt1 $menuDMA.Text.Replace('&','') -smenuopt2 $menurmssql.Text.Replace('&','') }
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
  $mainFormSI.width = 600
  $mainFormSI.height = 250
 
  # Title Label
  $titleLabel = New-Object System.Windows.Forms.Label
  $titleLabel.Font = $header#"Comic Sans MS,14"
  $titleLabel.ForeColor = "Yellow"
  $titleLabel.Location = "30,20"
  $titleLabel.Size = "500,50"
  $titleLabel.Text = "Select any Solution from below list"
  $mainFormSI.Controls.Add($titleLabel);
  #$mainFormSI.Controls.Add($titleLabel)

  # Input Box
  #$textBoxIn = New-Object System.Windows.Forms.TextBox
  #$textBoxIn.Location = "35, 70"
  #$textBoxIn.Size = "500, 20"
  #$textBoxIn.Text = ""
  #$mainFormSI.Controls.Add($textBoxIn)
 
  #Combobox
  $the_combo = New-Object system.Windows.Forms.ComboBox
  $the_combo.location = "35, 80"

  $the_combo.Size = "400, 100"

  $the_combo.DropDownStyle = "Dropdownlist"
  #$ComboList_Items = Get-Content $DMA_Path"DMA_DB_Type.txt"
  $path ="U:\solution_master\$mmenuopt\$smenuopt1\$smenuopt2"
  #$path ="U:\solution_master\$mmenuopt\$smenuopt1"
  $ComboList_Items = $null
 
  $ComboList_Items = Get-ChildItem -Path $path -filter "*.pdf"
  if($ComboList_Items -eq $NULL)
  {
     write-host "No file exist" -ForegroundColor Yellow
     [void] [System.Windows.MessageBox]::Show( "There is no File exist ", "Script completed", "OK", "Information" )
     #break
  }
  else
  {
  $ComboList_Itemsrep = $ComboList_Items.name.replace('.pdf','')
  #$path

  
 
  #AzureSqlVirtualMachine
  #Loop thru the text file or the array
  #and add the contents to the combobox for selection
  ForEach ($Server in $ComboList_Itemsrep) 
  {

    
  
     $the_combo.Items.Add($Server)
     
   
  }
  
  $mainFormSI.controls.add($the_combo)

  #action that will capture every time a value is selected on the combobox
  $the_combo_SelectedIndexChanged=
  {

    $targetcheck= $the_combo.text
    write-host "$targetcheck is the selection of solution" -ForegroundColor Yellow
    $path
  }

  $the_combo.add_SelectedIndexChanged($the_combo_SelectedIndexChanged)

  # Process Button
  $buttonProcess = New-Object System.Windows.Forms.Button
  $buttonProcess.Location = "35,150"
  $buttonProcess.Size = "75, 23"
  $buttonProcess.ForeColor = "Red"
  $buttonProcess.BackColor = "White"
  $buttonProcess.Text = "View"
  $buttonProcess.Font = $header
  #$buttonProcess.add_Click{processsingleServer} $SolutionMasterReports
  $buttonProcess.add_Click{AzureSQLDatabases_Solutions -mainmenu $mmenuopt -submenu $smenuopt1 -submenu1 $smenuopt2}
  #$buttonProcess.add_Click{$SolutionMasterReports}
  $mainFormSI.Controls.Add($buttonProcess)
 
  # Exit Button 
  $exitButton = New-Object System.Windows.Forms.Button
  $exitButton.Location = "450,150"
  $exitButton.Size = "75,23"
  $exitButton.Font = $header
  $exitButton.ForeColor = "Red"
  $exitButton.BackColor = "White"
  $exitButton.Text = "Exit"
  $exitButton.add_Click{$mainFormSI.close()}
  $mainFormSI.Controls.Add($exitButton)
  #[void]$mainFormSI.ShowDialog()
  [void]$mainFormSI.ShowDialog()
  }
  
}

Function AzureSQLDatabases_Solutions
{
 param (
       
        [Parameter(Mandatory=$true)][string]$mainmenu,
        [Parameter(Mandatory=$true)][string]$submenu,
        [Parameter(Mandatory=$true)][string]$submenu1
    )
  
  $combovalue = $the_combo.text
  #$combovalue 
  #$sm
  #$menuoptions 
  #Write-Host "This is ........$combovalue.............."
  Start-Process ((Resolve-Path "U:\solution_master\$mainmenu\$submenu\$submenu1\$combovalue.pdf").Path)
  
  
 }
 
# WinForm Setup
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
$menusql2012      = New-Object System.Windows.Forms.ToolStripMenuItem
$menusql2014      = New-Object System.Windows.Forms.ToolStripMenuItem
$menusql2019LIN   = New-Object System.Windows.Forms.ToolStripMenuItem
$menusql2019      = New-Object System.Windows.Forms.ToolStripMenuItem
$menusql2017LIN   = New-Object System.Windows.Forms.ToolStripMenuItem
$menusql2017     = New-Object System.Windows.Forms.ToolStripMenuItem
$menusql2016     = New-Object System.Windows.Forms.ToolStripMenuItem
$menusql2019LINvm   = New-Object System.Windows.Forms.ToolStripMenuItem
$menusql2019vm     = New-Object System.Windows.Forms.ToolStripMenuItem
$menusql2017LINvm   = New-Object System.Windows.Forms.ToolStripMenuItem
$menusql2017vm     = New-Object System.Windows.Forms.ToolStripMenuItem
$menusql2016vm     = New-Object System.Windows.Forms.ToolStripMenuItem
$menuExit         = New-Object System.Windows.Forms.ToolStripMenuItem
$menusins         = New-Object System.Windows.Forms.ToolStripMenuItem
$menuExitDM       = New-Object System.Windows.Forms.ToolStripMenuItem

$menuHelp         = New-Object System.Windows.Forms.ToolStripMenuItem
$menuAbout        = New-Object System.Windows.Forms.ToolStripMenuItem
$mainToolStrip    = New-Object System.Windows.Forms.ToolStrip
$toolStripOpen    = New-Object System.Windows.Forms.ToolStripButton
$toolStripSave    = New-Object System.Windows.Forms.ToolStripButton
$toolStripSaveAs  = New-Object System.Windows.Forms.ToolStripButton
$toolStripFullScr = New-Object System.Windows.Forms.ToolStripButton
$toolStripAbout   = New-Object System.Windows.Forms.ToolStripButton
$toolStripExit    = New-Object System.Windows.Forms.ToolStripButton
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
$menuGCP = New-Object System.Windows.Forms.ToolStripMenuItem

$menurazrinvm=New-Object System.Windows.Forms.ToolStripMenuItem

$menurazlog=New-Object System.Windows.Forms.ToolStripMenuItem

$menurazlogsucess =New-Object System.Windows.Forms.ToolStripMenuItem
$menurazlogfail =New-Object System.Windows.Forms.ToolStripMenuItem

$menurazGR=New-Object System.Windows.Forms.ToolStripMenuItem

$menurazFM=New-Object System.Windows.Forms.ToolStripMenuItem
$menurazMath=New-Object System.Windows.Forms.ToolStripMenuItem

$menurazFMSQL =New-Object System.Windows.Forms.ToolStripMenuItem
$menurazFMOra =New-Object System.Windows.Forms.ToolStripMenuItem
$menurazFMMySQL =New-Object System.Windows.Forms.ToolStripMenuItem
$menurazFMPSQL = New-Object System.Windows.Forms.ToolStripMenuItem
$menurazFMExaD = New-Object System.Windows.Forms.ToolStripMenuItem
$menurazMathSQL = New-Object System.Windows.Forms.ToolStripMenuItem
$menurazMathOra = New-Object System.Windows.Forms.ToolStripMenuItem
$menurazMathExaD = New-Object System.Windows.Forms.ToolStripMenuItem
$menurazMathPSQL = New-Object System.Windows.Forms.ToolStripMenuItem
$menurazMathMySQL = New-Object System.Windows.Forms.ToolStripMenuItem
$menurazOrRep = New-Object System.Windows.Forms.ToolStripMenuItem
$menurazpgrRep = New-Object System.Windows.Forms.ToolStripMenuItem
$menurazSamRep = New-Object System.Windows.Forms.ToolStripMenuItem
$menurazMSRep = New-Object System.Windows.Forms.ToolStripMenuItem
# solution master menu
$menuPredefineSolms = New-Object System.Windows.Forms.ToolStripMenuItem
$menumsoth = New-Object System.Windows.Forms.ToolStripMenuItem
$menuawsrds = New-Object System.Windows.Forms.ToolStripMenuItem
$menuawsoth = New-Object System.Windows.Forms.ToolStripMenuItem
$menuawsdp = New-Object System.Windows.Forms.ToolStripMenuItem
$menuoradp = New-Object System.Windows.Forms.ToolStripMenuItem
$menuoraoth = New-Object System.Windows.Forms.ToolStripMenuItem
$menuoraapps = New-Object System.Windows.Forms.ToolStripMenuItem
$menurosdbs = New-Object System.Windows.Forms.ToolStripMenuItem
$menurosdbspsql = New-Object System.Windows.Forms.ToolStripMenuItem
$menurosdbsmongo = New-Object System.Windows.Forms.ToolStripMenuItem
$menurmssql  = New-Object System.Windows.Forms.ToolStripMenuItem
$menurpssql  = New-Object System.Windows.Forms.ToolStripMenuItem
$menurazoth  = New-Object System.Windows.Forms.ToolStripMenuItem
$menurazdlake  = New-Object System.Windows.Forms.ToolStripMenuItem
$menurazhadoop = New-Object System.Windows.Forms.ToolStripMenuItem

######
#second menu

$menuDDoc = New-Object System.Windows.Forms.ToolStripMenuItem
$menuDDocAz = New-Object System.Windows.Forms.ToolStripMenuItem
$menuDDocMS = New-Object System.Windows.Forms.ToolStripMenuItem
$menuDDocPgsql = New-Object System.Windows.Forms.ToolStripMenuItem
$menurosdbsmongo = New-Object System.Windows.Forms.ToolStripMenuItem
$menuDDocorc2 = New-Object System.Windows.Forms.ToolStripMenuItem
$menuDDocorc = New-Object System.Windows.Forms.ToolStripMenuItem
$menuDDocoracld = New-Object System.Windows.Forms.ToolStripMenuItem
$menuDDocaws = New-Object System.Windows.Forms.ToolStripMenuItem
$menuDDocothrs = New-Object System.Windows.Forms.ToolStripMenuItem
$menuDDochadoop = New-Object System.Windows.Forms.ToolStripMenuItem
$menuDDocdplate = New-Object System.Windows.Forms.ToolStripMenuItem
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
 $Image = [system.drawing.image]::FromFile("U:\solution_master\image\solution_master.jpg") 
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
$menuSQLServer.Text = "&Technical Documents"
$menuSQLServer.Font = $header#"Comic Sans MS,14"
$menuSQLServer.ForeColor = "DarkBlue"
#$menuDMA.Font = Arial
#$menuSQLServer.Add_Click{$menuOracle.Enabled = $False}
[void]$menuMain.Items.Add($menuSQLServer)

# Menu Options - File
$menuDMA.Text = "&Azure"
$menuDMA.Font = $header
$menuDMA.ForeColor = "DarkBlue"
#$menuDMA.Font = Arial
#$menuDMA.Add_Click{SolutionMaster-WindowSubForm -Sol_Opt 'Azure_Databases' -SolutionMasterReports 'AzureSQLDatabases_Solutions' }
#$menuDMA.Add_Click{SolutionMaster-WindowSubForm -Sol_Opt $menuDMA.Text.Replace('&','') -SolutionMasterReports $menuSQLServer.Text.Replace('&','') }
[void]$menuSQLServer.DropDownItems.Add($menuDMA)

# Menu Options - MSSQL
$menurmssql.Text = "&MSSQL"
$menurmssql.Font = $header#"Comic Sans MS,14"
$menurmssql.ForeColor = "DarkBlue"
#$menurmssql.Add_Click{DMA_files}
#$menurmssql.Font = Arial
$menurmssql.Add_Click{SolutionMaster-WindowSubForm -mmenuopt $menuSQLServer.Text.Replace('&','') -smenuopt1 $menuDMA.Text.Replace('&','') -smenuopt2 $menurmssql.Text.Replace('&','') }
#$menurmssql.Add_Click{SolutionMaster-WindowSubForm $menuSQLServer.Text.Replace('&','') -Sol_Opt $menuDMA.Text.Replace('&','') -SolutionMasterReports $menurmssql.Text.Replace('&','') }
[void]$menuDMA.DropDownItems.Add($menurmssql)

# Menu Options - Postgresql
$menurpssql.Text = "&PostgreSQL"
$menurpssql.Font = $header#"Comic Sans MS,14"
$menurpssql.ForeColor = "DarkBlue"
#$menurpssql.Add_Click{DMA_files}
#$menurpssql.Font = Arial
$menurpssql.Add_Click{SolutionMaster-WindowSubForm -mmenuopt $menuSQLServer.Text.Replace('&','') -smenuopt1 $menuDMA.Text.Replace('&','') -smenuopt2 $menurpssql.Text.Replace('&','') }
[void]$menuDMA.DropDownItems.Add($menurpssql)

# Menu Options - Open source DB
$menurosdbsmongo.Text = "&No SQL"
$menurosdbsmongo.Font = $header#"Comic Sans MS,14"
$menurosdbsmongo.ForeColor = "DarkBlue"
#$menurazdbdoc.Add_Click{DMA_files}
#$menuDMA.Font = Arial
$menurosdbsmongo.Add_Click{SolutionMaster-WindowSubForm -mmenuopt $menuSQLServer.Text.Replace('&','') -smenuopt1 $menuDMA.Text.Replace('&','') -smenuopt2 $menurosdbsmongo.Text.Replace('&','') }
[void]$menuDMA.DropDownItems.Add($menurosdbsmongo)

# Menu Options - reference
$menurazmidoc.Text = "&Oracle"
$menurazmidoc.Font = $header#"Comic Sans MS,14"
$menurazmidoc.ForeColor = "DarkBlue"
#$menurazmidoc.Add_Click{SKU_files}
#$menuDMA.Font = Arial
$menurazmidoc.Add_Click{SolutionMaster-WindowSubForm -mmenuopt $menuSQLServer.Text.Replace('&','') -smenuopt1 $menuDMA.Text.Replace('&','') -smenuopt2 $menurazmidoc.Text.Replace('&','') }
[void]$menuDMA.DropDownItems.Add($menurazmidoc)

# Menu Options - reference
$menurazmidoc2.Text = "&Data Lake"
$menurazmidoc2.Font = $header#"Comic Sans MS,14"
$menurazmidoc2.ForeColor = "DarkBlue"
#$menurazmidoc2.Add_Click{Tool_Inventory}
#$menuDMA.Font = Arial
$menurazmidoc2.Add_Click{SolutionMaster-WindowSubForm -mmenuopt $menuSQLServer.Text.Replace('&','') -smenuopt1 $menuDMA.Text.Replace('&','') -smenuopt2 $menurazmidoc2.Text.Replace('&','') }
[void]$menuDMA.DropDownItems.Add($menurazmidoc2)

# Menu Options - DataLake
$menurazdlake.Text = "&Data Platform"
$menurazdlake.Font = $header#"Comic Sans MS,14"
$menurazdlake.ForeColor = "DarkBlue"
#$menurazmidoc2.Add_Click{Tool_Inventory}
#$menuDMA.Font = Arial
$menurazdlake.Add_Click{SolutionMaster-WindowSubForm -mmenuopt $menuSQLServer.Text.Replace('&','') -smenuopt1 $menuDMA.Text.Replace('&','') -smenuopt2 $menurazdlake.Text.Replace('&','') }
[void]$menuDMA.DropDownItems.Add($menurazdlake)

# Menu Options - others
$menurazoth.Text = "&Hadoop"
$menurazoth.Font = $header#"Comic Sans MS,14"
$menurazoth.ForeColor = "DarkBlue"
#$menurazmidoc2.Add_Click{Tool_Inventory}
#$menuDMA.Font = Arial
$menurazoth.Add_Click{SolutionMaster-WindowSubForm -mmenuopt $menuSQLServer.Text.Replace('&','') -smenuopt1 $menuDMA.Text.Replace('&','') -smenuopt2 $menurazoth.Text.Replace('&','') }
[void]$menuDMA.DropDownItems.Add($menurazoth)

# Menu Options - others
$menurazhadoop.Text = "&Others"
$menurazhadoop.Font = $header#"Comic Sans MS,14"
$menurazhadoop.ForeColor = "DarkBlue"
#$menurazmidoc2.Add_Click{Tool_Inventory}
#$menuDMA.Font = Arial
$menurazhadoop.Add_Click{SolutionMaster-WindowSubForm -mmenuopt $menuSQLServer.Text.Replace('&','') -smenuopt1 $menuDMA.Text.Replace('&','') -smenuopt2 $menurazhadoop.Text.Replace('&','') }
[void]$menuDMA.DropDownItems.Add($menurazhadoop)


# Menu Options - Data plateform
$menuSKU.Text      = "&AWS"
$menuSKU.Font = $header#"Comic Sans MS,14"
$menuSKU.ForeColor = "DarkBlue"
#$menuDMA.Add_Click{SolutionMaster-WindowSubForm -Sol_Opt $menuDMA.Text.Replace('&','') -SolutionMasterReports 'Framework' }
$menuSKU.Add_Click{SolutionMaster-WindowSubForm -Sol_Opt $menuSKU.Text.Replace('&','') -SolutionMasterReports $menuSQLServer.Text.Replace('&','') }
[void]$menuSQLServer.DropDownItems.Add($menuSKU)

# Menu Options - Data plateform
$menumsoth.Text      = "&Oracle Cloud"
$menumsoth.Font = $header#"Comic Sans MS,14"
$menumsoth.ForeColor = "DarkBlue"
$menumsoth.Add_Click{SolutionMaster-WindowSubForm -Sol_Opt $menumsoth.Text.Replace('&','') -SolutionMasterReports $menuSQLServer.Text.Replace('&','') }
[void]$menuSQLServer.DropDownItems.Add($menumsoth)

######################Second Menu############################
<#
# Menu Options - AWS
$menuAWS.Text = "&Design Documents"
$menuAWS.Font = $header#"Comic Sans MS,14"
$menuAWS.ForeColor = "DarkBlue"
#$menuDMA.Font = Arial
#$menuAWS.Add_Click{message}
$menuAWS.Enabled = $TRUE
[void]$menuMain.Items.Add($menuAWS)

#######################

# Menu Options - File
$menuDMA.Text = "&Azure"
$menuDMA.Font = $header
$menuDMA.ForeColor = "DarkBlue"
#$menuDMA.Font = Arial
#$menuDMA.Add_Click{SolutionMaster-WindowSubForm -Sol_Opt 'Azure_Databases' -SolutionMasterReports 'AzureSQLDatabases_Solutions' }
#$menuDMA.Add_Click{SolutionMaster-WindowSubForm -Sol_Opt $menuDMA.Text.Replace('&','') -SolutionMasterReports $menuSQLServer.Text.Replace('&','') }
[void]$menuAWS.DropDownItems.Add($menuDMA)

# Menu Options - MSSQL
$menurmssql.Text = "&MSSQL"
$menurmssql.Font = $header#"Comic Sans MS,14"
$menurmssql.ForeColor = "DarkBlue"
#$menurmssql.Add_Click{DMA_files}
#$menurmssql.Font = Arial
$menurmssql.Add_Click{SolutionMaster-WindowSubForm -mmenuopt $menuAWS.Text.Replace('&','') -smenuopt1 $menuDMA.Text.Replace('&','') -smenuopt2 $menurmssql.Text.Replace('&','') }
#$menurmssql.Add_Click{SolutionMaster-WindowSubForm $menuSQLServer.Text.Replace('&','') -Sol_Opt $menuDMA.Text.Replace('&','') -SolutionMasterReports $menurmssql.Text.Replace('&','') }
[void]$menuDMA.DropDownItems.Add($menurmssql)

# Menu Options - Postgresql
$menurpssql.Text = "&PostgreSQL"
$menurpssql.Font = $header#"Comic Sans MS,14"
$menurpssql.ForeColor = "DarkBlue"
#$menurpssql.Add_Click{DMA_files}
#$menurpssql.Font = Arial
$menurpssql.Add_Click{SolutionMaster-WindowSubForm -mmenuopt $menuAWS.Text.Replace('&','') -smenuopt1 $menuDMA.Text.Replace('&','') -smenuopt2 $menurpssql.Text.Replace('&','') }
[void]$menuDMA.DropDownItems.Add($menurpssql)

# Menu Options - Open source DB
$menurosdbsmongo.Text = "&No SQL"
$menurosdbsmongo.Font = $header#"Comic Sans MS,14"
$menurosdbsmongo.ForeColor = "DarkBlue"
#$menurazdbdoc.Add_Click{DMA_files}
#$menuDMA.Font = Arial
$menurosdbsmongo.Add_Click{SolutionMaster-WindowSubForm -mmenuopt $menuAWS.Text.Replace('&','') -smenuopt1 $menuDMA.Text.Replace('&','') -smenuopt2 $menurosdbsmongo.Text.Replace('&','') }
[void]$menuDMA.DropDownItems.Add($menurosdbsmongo)

# Menu Options - reference
$menurazmidoc.Text = "&Oracle"
$menurazmidoc.Font = $header#"Comic Sans MS,14"
$menurazmidoc.ForeColor = "DarkBlue"
#$menurazmidoc.Add_Click{SKU_files}
#$menuDMA.Font = Arial
$menurazmidoc.Add_Click{SolutionMaster-WindowSubForm -mmenuopt $menuAWS.Text.Replace('&','') -smenuopt1 $menuDMA.Text.Replace('&','') -smenuopt2 $menurazmidoc.Text.Replace('&','') }
[void]$menuDMA.DropDownItems.Add($menurazmidoc)

# Menu Options - reference
$menurazmidoc2.Text = "&Data Lake"
$menurazmidoc2.Font = $header#"Comic Sans MS,14"
$menurazmidoc2.ForeColor = "DarkBlue"
#$menurazmidoc2.Add_Click{Tool_Inventory}
#$menuDMA.Font = Arial
$menurazmidoc2.Add_Click{SolutionMaster-WindowSubForm -mmenuopt $menuAWS.Text.Replace('&','') -smenuopt1 $menuDMA.Text.Replace('&','') -smenuopt2 $menurazmidoc2.Text.Replace('&','') }
[void]$menuDMA.DropDownItems.Add($menurazmidoc2)

# Menu Options - DataLake
$menurazdlake.Text = "&Data Platform"
$menurazdlake.Font = $header#"Comic Sans MS,14"
$menurazdlake.ForeColor = "DarkBlue"
#$menurazmidoc2.Add_Click{Tool_Inventory}
#$menuDMA.Font = Arial
$menurazdlake.Add_Click{SolutionMaster-WindowSubForm -mmenuopt $menuAWS.Text.Replace('&','') -smenuopt1 $menuDMA.Text.Replace('&','') -smenuopt2 $menurazdlake.Text.Replace('&','') }
[void]$menuDMA.DropDownItems.Add($menurazdlake)

# Menu Options - others
$menurazoth.Text = "&Hadoop"
$menurazoth.Font = $header#"Comic Sans MS,14"
$menurazoth.ForeColor = "DarkBlue"
#$menurazmidoc2.Add_Click{Tool_Inventory}
#$menuDMA.Font = Arial
$menurazoth.Add_Click{SolutionMaster-WindowSubForm -mmenuopt $menuAWS.Text.Replace('&','') -smenuopt1 $menuDMA.Text.Replace('&','') -smenuopt2 $menurazoth.Text.Replace('&','') }
[void]$menuDMA.DropDownItems.Add($menurazoth)

# Menu Options - others
$menurazhadoop.Text = "&Others"
$menurazhadoop.Font = $header#"Comic Sans MS,14"
$menurazhadoop.ForeColor = "DarkBlue"
#$menurazmidoc2.Add_Click{Tool_Inventory}
#$menuDMA.Font = Arial
$menurazhadoop.Add_Click{SolutionMaster-WindowSubForm -mmenuopt $menuAWS.Text.Replace('&','') -smenuopt1 $menuDMA.Text.Replace('&','') -smenuopt2 $menurazhadoop.Text.Replace('&','') }
[void]$menuDMA.DropDownItems.Add($menurazhadoop)

# Menu Options - Data plateform
$menuSKU.Text      = "&AWS"
$menuSKU.Font = $header#"Comic Sans MS,14"
$menuSKU.ForeColor = "DarkBlue"
#$menuDMA.Add_Click{SolutionMaster-WindowSubForm -Sol_Opt $menuDMA.Text.Replace('&','') -SolutionMasterReports 'Framework' }
$menuSKU.Add_Click{SolutionMaster-WindowSubForm -Sol_Opt $menuSKU.Text.Replace('&','') -SolutionMasterReports $menuAWS.Text.Replace('&','') }
[void]$menuAWS.DropDownItems.Add($menuSKU)

# Menu Options - Data plateform
$menumsoth.Text      = "&Oracle Cloud"
$menumsoth.Font = $header#"Comic Sans MS,14"
$menumsoth.ForeColor = "DarkBlue"
$menumsoth.Add_Click{SolutionMaster-WindowSubForm -Sol_Opt $menumsoth.Text.Replace('&','') -SolutionMasterReports $menuAWS.Text.Replace('&','') }
[void]$menuAWS.DropDownItems.Add($menumsoth)



# Menu Options - Oracle
$menuOracle.Text = "&Methodology"
$menuOracle.Font = $header#"Comic Sans MS,14"
$menuOracle.ForeColor = "DarkBlue"
$menuOracle.Enabled = $false
#$menuOracle.On_Click{$menureferdoc.Enable = $false}
#$menuDMA.Font = Arial
#$menuOracle.Add_Click{Start-Process C:\myDB_Assessment_Report\Oracle\putty.exe}
#$menuOracle.Add_Click{$mainForm.Close()}
[void]$menuMain.Items.Add($menuOracle)

# Menu Options - AWS part1
#$menuMIDMA.Image        = [System.IconExtractor]::Extract("shell32.dll", 36, $true)
#$menuMIDMA.ShortcutKeys = "F2"
$menuoradp.Text         = "&Azure"
$menuoradp.ForeColor = "DarkBlue"
#$menuIaaSAss.Add_Click{AzureDBDMA -Target 'ManagedSqlServer'}
[void]$menuOracle.DropDownItems.Add($menuoradp)


# Menu option - ora part2
$menuoraapps.Text         = "&AWS"
$menuoraapps.ForeColor = "DarkBlue"
#$menuIaaSAss.Add_Click{AzureDBDMA -Target 'ManagedSqlServer'}
[void]$menuOracle.DropDownItems.Add($menuoraapps)

# Menu option - ora part3
$menuoraoth.Text         = "&Oracle Cloud"
$menuoraoth.ForeColor = "DarkBlue"
#$menuIaaSAss.Add_Click{AzureDBDMA -Target 'ManagedSqlServer'}
[void]$menuOracle.DropDownItems.Add($menuoraoth)
#>

# Menu Options - Design Documents
$menuDDoc.Text = "&Design Documents"
$menuDDoc.Font = $header#"Comic Sans MS,14"
$menuDDoc.ForeColor = "DarkBlue"
#$menuDDocAz.Font = Arial
#$menuDDoc.Add_Click{$menuOracle.Enabled = $False}
[void]$menuMain.Items.Add($menuDDoc)

# Menu Options - DDocAz
$menuDDocAz.Text = "&Azure"
$menuDDocAz.Font = $header
$menuDDocAz.ForeColor = "DarkBlue"
#$menuDDocAz.Font = Arial
#$menuDDocAz.Add_Click{SolutionMaster-WindowSubForm -Sol_Opt 'Azure_Databases' -SolutionMasterReports 'AzureSQLDatabases_Solutions' }
#$menuDDocAz.Add_Click{SolutionMaster-WindowSubForm -Sol_Opt $menuDDocAz.Text.Replace('&','') -SolutionMasterReports $menuDDoc.Text.Replace('&','') }
[void]$menuDDoc.DropDownItems.Add($menuDDocAz)

# Menu Options - DDocMS
$menuDDocMS.Text = "&MSSQL"
$menuDDocMS.Font = $header#"Comic Sans MS,14"
$menuDDocMS.ForeColor = "DarkBlue"
#$menuDDocMS.Add_Click{DMA_files}
#$menuDDocMS.Font = Arial
$menuDDocMS.Add_Click{SolutionMaster-WindowSubForm -mmenuopt $menuDDoc.Text.Replace('&','') -smenuopt1 $menuDDocAz.Text.Replace('&','') -smenuopt2 $menuDDocMS.Text.Replace('&','') }
#$menuDDocMS.Add_Click{SolutionMaster-WindowSubForm $menuDDoc.Text.Replace('&','') -Sol_Opt $menuDDocAz.Text.Replace('&','') -SolutionMasterReports $menuDDocMS.Text.Replace('&','') }
[void]$menuDDocAz.DropDownItems.Add($menuDDocMS)

# Menu Options - DDocPgsql
$menuDDocPgsql.Text = "&PostgreSQL"
$menuDDocPgsql.Font = $header#"Comic Sans MS,14"
$menuDDocPgsql.ForeColor = "DarkBlue"
#$menuDDocPgsql.Add_Click{DMA_files}
#$menuDDocPgsql.Font = Arial
$menuDDocPgsql.Add_Click{SolutionMaster-WindowSubForm -mmenuopt $menuDDoc.Text.Replace('&','') -smenuopt1 $menuDDocAz.Text.Replace('&','') -smenuopt2 $menuDDocPgsql.Text.Replace('&','') }
[void]$menuDDocAz.DropDownItems.Add($menuDDocPgsql)

# Menu Options - DDocmgo
$menurosdbsmongo.Text = "&No SQL"
$menurosdbsmongo.Font = $header#"Comic Sans MS,14"
$menurosdbsmongo.ForeColor = "DarkBlue"
#$menurazdbdoc.Add_Click{DMA_files}
#$menuDDocAz.Font = Arial
$menurosdbsmongo.Add_Click{SolutionMaster-WindowSubForm -mmenuopt $menuDDoc.Text.Replace('&','') -smenuopt1 $menuDDocAz.Text.Replace('&','') -smenuopt2 $menurosdbsmongo.Text.Replace('&','') }
[void]$menuDDocAz.DropDownItems.Add($menurosdbsmongo)

# Menu Options - DDocorc
$menuDDocorc.Text = "&Oracle"
$menuDDocorc.Font = $header#"Comic Sans MS,14"
$menuDDocorc.ForeColor = "DarkBlue"
#$menuDDocorc.Add_Click{SKU_files}
#$menuDDocAz.Font = Arial
$menuDDocorc.Add_Click{SolutionMaster-WindowSubForm -mmenuopt $menuDDoc.Text.Replace('&','') -smenuopt1 $menuDDocAz.Text.Replace('&','') -smenuopt2 $menuDDocorc.Text.Replace('&','') }
[void]$menuDDocAz.DropDownItems.Add($menuDDocorc)

# Menu Options - DDocorc2
$menuDDocorc2.Text = "&Data Lake"
$menuDDocorc2.Font = $header#"Comic Sans MS,14"
$menuDDocorc2.ForeColor = "DarkBlue"
#$menuDDocorc2.Add_Click{Tool_Inventory}
#$menuDDocAz.Font = Arial
$menuDDocorc2.Add_Click{SolutionMaster-WindowSubForm -mmenuopt $menuDDoc.Text.Replace('&','') -smenuopt1 $menuDDocAz.Text.Replace('&','') -smenuopt2 $menuDDocorc2.Text.Replace('&','') }
[void]$menuDDocAz.DropDownItems.Add($menuDDocorc2)

# Menu Options - DDocdplate
$menuDDocdplate.Text = "&Data Platform"
$menuDDocdplate.Font = $header#"Comic Sans MS,14"
$menuDDocdplate.ForeColor = "DarkBlue"
#$menuDDocorc2.Add_Click{Tool_Inventory}
#$menuDDocAz.Font = Arial
$menuDDocdplate.Add_Click{SolutionMaster-WindowSubForm -mmenuopt $menuDDoc.Text.Replace('&','') -smenuopt1 $menuDDocAz.Text.Replace('&','') -smenuopt2 $menuDDocdplate.Text.Replace('&','') }
[void]$menuDDocAz.DropDownItems.Add($menuDDocdplate)

# Menu Options - DDochadoop
$menuDDochadoop.Text = "&Hadoop"
$menuDDochadoop.Font = $header#"Comic Sans MS,14"
$menuDDochadoop.ForeColor = "DarkBlue"
#$menuDDocorc2.Add_Click{Tool_Inventory}
#$menuDDocAz.Font = Arial
$menuDDochadoop.Add_Click{SolutionMaster-WindowSubForm -mmenuopt $menuDDoc.Text.Replace('&','') -smenuopt1 $menuDDocAz.Text.Replace('&','') -smenuopt2 $menuDDochadoop.Text.Replace('&','') }
[void]$menuDDocAz.DropDownItems.Add($menuDDochadoop)

# Menu Options - DDocothrs
$menuDDocothrs.Text = "&Others"
$menuDDocothrs.Font = $header#"Comic Sans MS,14"
$menuDDocothrs.ForeColor = "DarkBlue"
#$menuDDocorc2.Add_Click{Tool_Inventory}
#$menuDDocAz.Font = Arial
$menuDDocothrs.Add_Click{SolutionMaster-WindowSubForm -mmenuopt $menuDDoc.Text.Replace('&','') -smenuopt1 $menuDDocAz.Text.Replace('&','') -smenuopt2 $menuDDocothrs.Text.Replace('&','') }
[void]$menuDDocAz.DropDownItems.Add($menuDDocothrs)


# Menu Options - DDocaws
$menuDDocaws.Text      = "&AWS"
$menuDDocaws.Font = $header#"Comic Sans MS,14"
$menuDDocaws.ForeColor = "DarkBlue"
#$menuDDocAz.Add_Click{SolutionMaster-WindowSubForm -Sol_Opt $menuDDocAz.Text.Replace('&','') -SolutionMasterReports 'Framework' }
$menuDDocaws.Add_Click{SolutionMaster-WindowSubForm -Sol_Opt $menuDDocaws.Text.Replace('&','') -SolutionMasterReports $menuDDoc.Text.Replace('&','') }
[void]$menuDDoc.DropDownItems.Add($menuDDocaws)

# Menu Options - DDocoracld
$menuDDocoracld.Text      = "&Oracle Cloud"
$menuDDocoracld.Font = $header#"Comic Sans MS,14"
$menuDDocoracld.ForeColor = "DarkBlue"
$menuDDocoracld.Add_Click{SolutionMaster-WindowSubForm -Sol_Opt $menuDDocoracld.Text.Replace('&','') -SolutionMasterReports $menuDDoc.Text.Replace('&','') }
[void]$menuDDoc.DropDownItems.Add($menuDDocoracld)

#####################

# Menu Options - Version
$menuver.Text      = "&Version"
$menuver.Font = $header#"Comic Sans MS,14"
$menuver.ForeColor = "DarkBlue"
[void]$menuMain.Items.Add($menuver)

# Menu Options - Help / About
$menuAbout.Image     = [System.Drawing.SystemIcons]::Information
$menuAbout.Text      = "About Solution Master"
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
$mainForm.add_Shown({Directory_Creation} )
Add-Type -AssemblyName System.Speech
$synthesizer = New-Object -TypeName System.Speech.Synthesis.SpeechSynthesizer
$synthesizer.Speak('Hey, Welcome to Solution Master')
#$synthesizer.Speak('Good Bye and see you')
[void] $mainForm.ShowDialog()
Add-Type -AssemblyName System.Speech
$synthesizer = New-Object -TypeName System.Speech.Synthesis.SpeechSynthesizer
#$synthesizer.Speak('Hey, Welcome to myDB Assessment Tool')
$synthesizer.Speak('Good Bye and see you')