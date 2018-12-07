## Code to hide the powershell command window when GUI is running
$t = '[DllImport("user32.dll")] public static extern bool ShowWindow(int handle, int state);'
add-type -name win -member $t -namespace native
[native.win]::ShowWindow(([System.Diagnostics.Process]::GetCurrentProcess() | Get-Process).MainWindowHandle, 0)
  
  #ERASE ALL THIS AND PUT XAML BELOW between the @" "@ 
$inputXML = @"
<Window x:Class="BlogPostIII.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:BlogPostIII"
        mc:Ignorable="d"
        Title="MRG Quick Ticketing Tool" Height="350" Width="616.976">
    <Grid x:Name="background" Background="#FF1D3245">
        <TextBlock x:Name="Title_Subject" HorizontalAlignment="Left" Height="31" Margin="10,24,0,0" TextWrapping="Wrap" Text="Subject: " VerticalAlignment="Top" Width="51" Foreground="White"/>
        <TextBox x:Name="SubjectField" HorizontalAlignment="Left" Height="21" Margin="61,24,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="406"/>
        <TextBlock x:Name="Title_Body" HorizontalAlignment="Left" Height="21" Margin="10,60,0,0" TextWrapping="Wrap" Text="Details:" VerticalAlignment="Top" Width="51" Foreground="White"/>
        <TextBox x:Name="BodyField" HorizontalAlignment="Left" Height="194" Margin="61,60,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="406"/>
        <TextBlock x:Name="Title_Email" HorizontalAlignment="Left" Height="21" Margin="10,276,0,0" TextWrapping="Wrap" Text="Email:" VerticalAlignment="Top" Width="34" Foreground="White"/>
        <TextBox x:Name="EmailField" HorizontalAlignment="Left" Height="21" Margin="49,276,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="231"/>
        <TextBlock x:Name="Title_Password" HorizontalAlignment="Left" Height="21" Margin="296,276,0,0" TextWrapping="Wrap" Text="Password:" VerticalAlignment="Top" Width="59" Foreground="White"/>
        <PasswordBox x:Name="passwordBox" HorizontalAlignment="Left" Height="21" Margin="360,276,0,0" VerticalAlignment="Top" Width="191"/>
        <TextBlock x:Name="Title_TimeSpent" HorizontalAlignment="Left" Height="21" Margin="482,24,0,0" TextWrapping="Wrap" Text="Time Spent:" VerticalAlignment="Top" Width="69" Foreground="White"/>
        <RadioButton x:Name="Selection5" Content="5m" HorizontalAlignment="Left" Margin="502,45,0,0" VerticalAlignment="Top" Foreground="White"/>
        <RadioButton x:Name="Selection15" Content="15m" HorizontalAlignment="Left" Margin="502,68,0,0" VerticalAlignment="Top" Foreground="White"/>
        <RadioButton x:Name="Selection30" Content="30m" HorizontalAlignment="Left" Margin="502,89,0,0" VerticalAlignment="Top" Foreground="White"/>
        <RadioButton x:Name="Selection1" Content="1h" HorizontalAlignment="Left" Margin="502,110,0,0" VerticalAlignment="Top" Foreground="White"/>
        <RadioButton x:Name="Selection2" Content="2h" HorizontalAlignment="Left" Margin="502,131,0,0" VerticalAlignment="Top" Foreground="White"/>
        <Button x:Name="Button" Content="Quick Ticket!" HorizontalAlignment="Left" Height="68" Margin="472,186,0,0" VerticalAlignment="Top" Width="127"/>
    </Grid>
</Window>
"@        

$inputXML = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N'  -replace '^<Win.*', '<Window'


[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = $inputXML
#Read XAML

    $reader=(New-Object System.Xml.XmlNodeReader $xaml) 
  try{$Form=[Windows.Markup.XamlReader]::Load( $reader )}
catch{Write-Host "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .net is installed."}

#===========================================================================
# Store Form Objects In PowerShell
#===========================================================================

$xaml.SelectNodes("//*[@Name]") | %{Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name)}

Function Get-FormVariables{
if ($global:ReadmeDisplay -ne $true){Write-host "If you need to reference this display again, run Get-FormVariables" -ForegroundColor Yellow;$global:ReadmeDisplay=$true}
write-host "Found the following interactable elements from our form" -ForegroundColor Cyan
get-variable WPF*
}

#Get-FormVariables

#===========================================================================
# Actually make the objects work
#===========================================================================
##Create a Windows Scripting host shell instance to support interactive popups.
$wshell = New-Object -ComObject Wscript.Shell
## Global Required Support Items
function Calculate+Append_WorkTime{
    if($WPFSelection5.IsChecked -eq $true){
        $WPFBodyField.Text = $WPFBodyField.Text + "`n #add 5m"
    }
    if($WPFSelection15.IsChecked -eq $true){
        $WPFBodyField.Text = $WPFBodyField.Text + "`n #add 15m"
    }
    if($WPFSelection30.IsChecked -eq $true){
        $WPFBodyField.Text = $WPFBodyField.Text + "`n #add 30m"
    }
    if($WPFSelection1.IsChecked -eq $true){
        $WPFBodyField.Text = $WPFBodyField.Text + "`n #add 1h"
    }
    if($WPFSelection2.IsChecked -eq $true){
        $WPFBodyField.Text = $WPFBodyField.Text + "`n #add 2h"
    }
}
function CleanUpRadioButtons{
    $WPFSelection5.IsChecked = $false
    $WPFSelection15.IsChecked = $false
    $WPFSelection30.IsChecked = $false
    $WPFSelection1.IsChecked = $false
    $WPFSelection2.IsChecked = $false

}
function Validate_Fields{
     if($WPFSubjectField.Text -eq "" -or $WPFBodyField.Text -eq "" -or $WPFpasswordbox.Password -eq "" -or $WPFEmailField.Text -eq ""){
        return $false
     }
     else{
        return $true
     }
}

#Give the buttons Actions
$WPFButton.Add_Click({
    if(Validate_Fields){
        #Append Spiceworks commands
        Calculate+Append_WorkTime
        $WPFBodyField.Text = $WPFBodyField.Text + "`n #close"
    
        #Geneate Credentals
        $SecuredPassword = ConvertTo-SecureString $WPFpasswordbox.Password -AsPlainText -Force
        $MyCredentals = New-Object System.Management.Automation.PSCredential($WPFEmailField.Text, $SecuredPassword)

        #Setup Email Structure
        $SmtpServer = 'smtp.office365.com'
        $MailtTo = 'helpdesk@medfordradiology.com'
        $MailFrom = $MyCredentals.UserName
        $MailSubject = $WPFSubjectField.Text
        $MailBody = $WPFBodyField.Text
    
        #Send Email to Spiceworks to Create Ticket
        Send-MailMessage -To "$MailtTo" -from "$MailFrom" -Subject $MailSubject -Body $MailBody -SmtpServer $SmtpServer -UseSsl -Port 587 -Credential $MyCredentals 
    
        #Alert user ticket was compleated.
        $wshell.Popup("Ticket Created.", 5, "MRG Quick Ticketing Tool", 0x30)
   
        ##Clear old Fields & Clean Up
        CleanUpRadioButtons
        $WPFBodyField.Text = ""
        $WPFSubjectField.Text = ""
        }
        else{
            $wshell.Popup("ERROR: all fields are NOT filled.", 5, "MRG Quick Ticketing Tool", 0x30)
        }
})


#===========================================================================
# Shows the form
#===========================================================================

function Show-Form{
$Form.ShowDialog() | out-null
}

Show-Form