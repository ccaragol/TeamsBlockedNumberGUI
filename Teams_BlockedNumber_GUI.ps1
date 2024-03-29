<# 
.Synopsis 
   The purpose of this tool is to give you an easy front end for managing blocked numbers within Microsoft
   Teams at an administrative level.  This is tenant-wide blocked number patterns which are avaialble at 
   this time only through PowerShell.
 
.DESCRIPTION 
   PowerShell GUI script which allows for ability to manage blocked number patterns
 
.Notes 
     NAME:      Teams_BlockedNumber_GUI.ps1
     VERSION:   1.0 
     AUTHOR:    C. Anthony Caragol 
     LASTEDIT:  03/15/2024
      
   V 1.0 - March 15 2024 - Initial release 
    
.Link 
   Github: http://www.github.com/ccaragol
   Twitter: http://www.twitter.com/canthonycaragol
   LinkedIn: http://www.linkedin.com/in/canthonycaragol
 
.EXAMPLE 
   .\Teams_BlockedNumber_GUI.ps1

.TODO
  1) I should probably add some comments, see .APOLOGY

.APOLOGY
  Please excuse the sloppy coding for now, I don't use a development environment, IDE or ISE.  I use notepad, 
  not even Notepad++, just notepad.  I am not a developer, just an enthusiast so some code may be redundant or
  inefficient.

  #https://learn.microsoft.com/en-us/microsoftteams/block-inbound-calls
#>

Function CheckForInstalledModules
{
    #Save-Module -Name MicrosoftTeams -Path . -RequiredVersion 0.9.0
    #Install-Module -Name MicrosoftTeams -RequiredVersion 0.9.0
    if (-not(get-module -ListAvailable "MicrosoftTeams")) {
        [Microsoft.VisualBasic.Interaction]::MsgBox("Microsoft Teams PowerShell Module Not Found.  Please run 'Install-Module -Name MicrosoftTeams' before continuing if you have not done so." ,'OKOnly,Information', "Teams PowerShell Module Not Installed!")
    }
}

Function ShutDownForm()
{
	$objForm.Close()
}


function LoadBlockedFromCurrent()
{

	$BlockedNumberGridView.Rows.Clear()
    
	foreach ($entry in (Get-CsInboundBlockedNumberPattern)) 
	{ 
				$BlockedNumberGridView.Rows.Add("Blocked",$entry.identity,$entry.name,$entry.enabled,$entry.pattern,$entry.description)
	}
	foreach ($entry in (Get-CsInboundExemptNumberPattern)) 
	{ 
				$BlockedNumberGridView.Rows.Add("Exempt",$entry.identity,$entry.name,$entry.enabled,$entry.pattern,$entry.description)
	}
}


$CAC_FormSizeChanged = { 
	
	$BlockedNumberGridView.Width=($objForm.Width - 40) 
    $BlockedNumberGridView.Columns[5].Width = $BlockedNumberGridView.width-469
} 
 
   
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.IO.Compression.FileSystem") 
[void] [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')

function LoadMainForm()
{
$objForm = New-Object System.Windows.Forms.Form 
$objForm.Text = "Teams Blocked Number Management Tool"
$objForm.Size = New-Object System.Drawing.Size(960,600) 
$objForm.StartPosition = "CenterScreen"
$ObjForm.Add_SizeChanged($CAC_FormSizeChanged)
$objForm.KeyPreview = $True

$TitleLabel = New-Object System.Windows.Forms.Label
$TitleLabel.Location = New-Object System.Drawing.Size(10,10) 
$TitleLabel.Size = New-Object System.Drawing.Size(900,60) 
$TitleLabel.Text = "The purpose of this tool is to give you an easy front end for managing blocked numbers within Microsoft Teams at an administrative level.  This is tenant-wide blocked number patterns which are accessible at the time of this script's creation only through PowerShell.  Use only at your own risk."
$objForm.Controls.Add($TitleLabel) 

$BlockedNumberGridView = New-Object System.Windows.Forms.DataGridView
$BlockedNumberGridView.Location = New-Object System.Drawing.Size(10,80) 
$BlockedNumberGridView.Size = New-Object System.Drawing.Size(920,400) 
$BlockedNumberGridView.Anchor = 'Top, Bottom, Left'
$BlockedNumberGridView.ColumnCount = 6
$BlockedNumberGridView.Columns[0].Width = 100
$BlockedNumberGridView.Columns[1].Width = 100
$BlockedNumberGridView.Columns[2].Width = 100
$BlockedNumberGridView.Columns[3].Width = 75
$BlockedNumberGridView.Columns[4].Width = 150
$BlockedNumberGridView.Columns[5].Width = $BlockedNumberGridView.width-469
$BlockedNumberGridView.Columns[0].Name = "Type"
$BlockedNumberGridView.Columns[1].Name = "Identity"
$BlockedNumberGridView.Columns[2].Name = "Name"
$BlockedNumberGridView.Columns[3].Name = "Enabled"
$BlockedNumberGridView.Columns[4].Name = "Pattern"
$BlockedNumberGridView.Columns[5].Name = "Description"
$objForm.Controls.Add($BlockedNumberGridView) 

$ButtonSpacing=5

$Step1Button = New-Object System.Windows.Forms.Button
$Step1Button.Location = New-Object System.Drawing.Size(10,490)
$Step1Button.Size = New-Object System.Drawing.Size(150,25)
$Step1Button.Text = "Connect"
$Step1Button.Add_Click({
Connect-MicrosoftTeams
LoadBlockedFromCurrent
})
$Step1Button.Anchor = 'Bottom, Left'
$objForm.Controls.Add($Step1Button)


$Step2Button = New-Object System.Windows.Forms.Button
$Step2Button.Location = New-Object System.Drawing.Size(($Step1Button.Location.X + $Step1Button.Width + $ButtonSpacing),490)
$Step2Button.Size = New-Object System.Drawing.Size(150,25)
$Step2Button.Text = "Refresh"
$Step2Button.Add_Click({
    LoadBlockedFromCurrent
})
$Step2Button.Anchor = 'Bottom, Left'
$objForm.Controls.Add($Step2Button)


$Step3Button = New-Object System.Windows.Forms.Button
$Step3Button.Location = New-Object System.Drawing.Size(($Step2Button.Location.X + $Step2Button.Width + $ButtonSpacing),490)
$Step3Button.Size = New-Object System.Drawing.Size(150,25)
$Step3Button.Text = "Add Pattern"
$Step3Button.Add_Click({
		LoadAddBlockedForm
})
$Step3Button.Anchor = 'Bottom, Left'
$objForm.Controls.Add($Step3Button)

$Step4Button = New-Object System.Windows.Forms.Button
$Step4Button.Location = New-Object System.Drawing.Size(($Step3Button.Location.X + $Step3Button.Width + $ButtonSpacing),490)
$Step4Button.Size = New-Object System.Drawing.Size(150,25)
$Step4Button.Text = "Remove Pattern"
$Step4Button.Add_Click({
    $message = "Are you sure you want to delete $($BlockedNumberGridView.CurrentRow.Cells["Identity"].Value) ? "
    $result = [System.Windows.Forms.MessageBox]::Show($message , "Info" , 4)
    if ($result -eq 'Yes') {
        
        if ($BlockedNumberGridView.CurrentRow.Cells["Type"].Value -eq "Blocked") 
            {
                write-host "Deleting Blocked Pattern $($BlockedNumberGridView.CurrentRow.Cells["Identity"].Value)"
                Remove-CsInboundBlockedNumberPattern -Identity $BlockedNumberGridView.CurrentRow.Cells["Identity"].Value
            }
        else
            {
                write-host "Deleting Exempt Pattern $($BlockedNumberGridView.CurrentRow.Cells["Identity"].Value)"
                Remove-CsInboundExemptNumberPattern -Identity $BlockedNumberGridView.CurrentRow.Cells["Identity"].Value
            }
        LoadBlockedFromCurrent
    }		
})
$Step4Button.Anchor = 'Bottom, Left'
$objForm.Controls.Add($Step4Button)

$Step5Button = New-Object System.Windows.Forms.Button
$Step5Button.Location = New-Object System.Drawing.Size(($Step4Button.Location.X + $Step4Button.Width + $ButtonSpacing),490)
$Step5Button.Size = New-Object System.Drawing.Size(150,25)
$Step5Button.Text = "Test Pattern"
$Step5Button.Add_Click({
		LoadTestForm $($BlockedNumberGridView.CurrentRow.Cells["Name"].Value) $($BlockedNumberGridView.CurrentRow.Cells["Pattern"].Value)
})
$Step5Button.Anchor = 'Bottom, Left'
$objForm.Controls.Add($Step5Button)

$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Size(($Step5Button.Location.X + $Step5Button.Width + $ButtonSpacing),490)
$CancelButton.Size = New-Object System.Drawing.Size(150,25)
$CancelButton.Text = "Quit"
$CancelButton.Add_Click({ShutDownForm})
$CancelButton.Anchor = 'Bottom, Left'
$objForm.Controls.Add($CancelButton)


#LyncFix LinkLabel
$LinkedInLinkLabel = New-Object System.Windows.Forms.LinkLabel
$LinkedInLinkLabel.Location = New-Object System.Drawing.Size(10,538) 
$LinkedInLinkLabel.Size = New-Object System.Drawing.Size(250,20)
$LinkedInLinkLabel.text = "https://www.linkedin.com/in/canthonycaragol"
$LinkedInLinkLabel.add_Click({Start-Process $LinkedInLinkLabel.text})
$LinkedInLinkLabel.Anchor = 'Bottom, Left'
$objForm.Controls.Add($LinkedInLinkLabel)



$objForm.Add_Shown({$objForm.Activate()})
[void] $objForm.ShowDialog()

}

function LoadAddBlockedForm()
{
$AddBlockedForm = New-Object System.Windows.Forms.Form
$AddBlockedForm.Text = 'Add Blocked Number Pattern'
$AddBlockedForm.Size = New-Object System.Drawing.Size(300,200)
$AddBlockedForm.StartPosition = 'CenterScreen'

$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(75,120)
$okButton.Size = New-Object System.Drawing.Size(75,23)
$okButton.Text = 'Add'
$oKButton.Add_Click({
	New-CsInboundBlockedNumberPattern -Name $AddBlockedNameText.text -Enabled $True -Description $AddBlockedDescText.text -Pattern $AddBlockedPatternText.text   
})
$AddBlockedForm.AcceptButton = $okButton
$AddBlockedForm.Controls.Add($okButton)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(150,120)
$cancelButton.Size = New-Object System.Drawing.Size(75,23)
$cancelButton.Text = 'Cancel'
$CancelButton.Add_Click({
	$AddBlockedForm.Close()
})
$AddBlockedForm.CancelButton = $cancelButton
$AddBlockedForm.Controls.Add($cancelButton)

$AddBlockedNameLabel = New-Object System.Windows.Forms.Label
$AddBlockedNameLabel.Location = New-Object System.Drawing.Point(10,20)
$AddBlockedNameLabel.Size = New-Object System.Drawing.Size(70,20)
$AddBlockedNameLabel.Text = 'Name:'
$AddBlockedForm.Controls.Add($AddBlockedNameLabel)

$AddBlockedNameText = New-Object System.Windows.Forms.TextBox
$AddBlockedNameText.Location = New-Object System.Drawing.Point(80,20)
$AddBlockedNameText.Size = New-Object System.Drawing.Size(120,20)
$AddBlockedForm.Controls.Add($AddBlockedNameText)
$AddBlockedForm.Topmost = $true


$AddBlockedPatternLabel = New-Object System.Windows.Forms.Label
$AddBlockedPatternLabel.Location = New-Object System.Drawing.Point(10,40)
$AddBlockedPatternLabel.Size = New-Object System.Drawing.Size(70,20)
$AddBlockedPatternLabel.Text = 'Pattern:'
$AddBlockedForm.Controls.Add($AddBlockedPatternLabel)

$AddBlockedPatternText = New-Object System.Windows.Forms.TextBox
$AddBlockedPatternText.Location = New-Object System.Drawing.Point(80,40)
$AddBlockedPatternText.Size = New-Object System.Drawing.Size(120,20)
$AddBlockedForm.Controls.Add($AddBlockedPatternText)
$AddBlockedForm.Topmost = $true


$AddBlockedDescLabel = New-Object System.Windows.Forms.Label
$AddBlockedDescLabel.Location = New-Object System.Drawing.Point(10,60)
$AddBlockedDescLabel.Size = New-Object System.Drawing.Size(70,20)
$AddBlockedDescLabel.Text = 'Description:'
$AddBlockedForm.Controls.Add($AddBlockedDescLabel)

$AddBlockedDescText = New-Object System.Windows.Forms.TextBox
$AddBlockedDescText.Location = New-Object System.Drawing.Point(80,60)
$AddBlockedDescText.Size = New-Object System.Drawing.Size(120,20)
$AddBlockedForm.Controls.Add($AddBlockedDescText)
$AddBlockedForm.Topmost = $true

$AddBlockedExLabel = New-Object System.Windows.Forms.Label
$AddBlockedExLabel.Location = New-Object System.Drawing.Point(10,90)
$AddBlockedExLabel.Size = New-Object System.Drawing.Size(250,20)
$AddBlockedExLabel.Text = 'Pattern Example: ""^\+?1312555\d{4}$""'
$AddBlockedForm.Controls.Add($AddBlockedExLabel)

#"^\+?1312555\d{4}$"


$AddBlockedForm.Add_Shown({$textBox.Select()})
$result = $AddBlockedForm.ShowDialog()
}

function LoadTestForm([string]$name, [string]$pattern)
{
$testForm = New-Object System.Windows.Forms.Form 
$testForm.Text = "Testing Pattern $($name)"
$testForm.Size = New-Object System.Drawing.Size(400,260) 
$testForm.StartPosition = "CenterScreen"
$testForm.KeyPreview = $True

$patternlabel = New-Object System.Windows.Forms.Label
$patternlabel.Location = New-Object System.Drawing.Size(10,20)
$patternlabel.Size = New-Object System.Drawing.Size(350,50)
$patternlabel.Text = "Enter Test Number for Pattern $($pattern) in the below box and click the Test button. Ideally, test the number in E.164 format (ex: +15556667777 or +44...)"
$patternlabel.Anchor = 'Bottom, Left'
$testForm.Controls.Add($patternlabel)

$patterntextbox = New-Object System.Windows.Forms.textbox
$patterntextbox.Location = New-Object System.Drawing.Size(10,100)
$patterntextbox.Size = New-Object System.Drawing.Size(200,50)
$patterntextbox.Anchor = 'Bottom, Left'
$testForm.Controls.Add($patterntextbox)

$patternpasslabel = New-Object System.Windows.Forms.Label
$patternpasslabel.Location = New-Object System.Drawing.Size(10,140)
$patternpasslabel.Size = New-Object System.Drawing.Size(200,20)
$patternpasslabel.Text = ""
$patternpasslabel.Anchor = 'Bottom, Left'
$testForm.Controls.Add($patternpasslabel)

$Step1Button = New-Object System.Windows.Forms.Button
$Step1Button.Location = New-Object System.Drawing.Size(10,180)
$Step1Button.Size = New-Object System.Drawing.Size(150,25)
$Step1Button.Text = "Test"
$Step1Button.Add_Click({
$x=Test-csinboundBlockedNumberPattern -PhoneNumber $patterntextbox.Text
$patternpasslabel.Text = "Pattern $($patterntextbox.Text) result = $($x)"
})
$Step1Button.Anchor = 'Bottom, Left'
$testForm.Controls.Add($Step1Button)


$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Size(($Step1Button.Location.X + $Step1Button.Width + $ButtonSpacing),180)
$CancelButton.Size = New-Object System.Drawing.Size(150,25)
$CancelButton.Text = "Quit"
$CancelButton.Add_Click({
	$testForm.Close()
})
$CancelButton.Anchor = 'Bottom, Left'
$testForm.Controls.Add($CancelButton)


$testForm.Add_Shown({$testForm.Activate()})
[void] $testForm.ShowDialog()

}



CheckForInstalledModules
LoadMainForm