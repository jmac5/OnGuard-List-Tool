# This program is meant to be used with the Lenel OnGuard application, and its
# MS SQL Database. You will need the hostnames\instances of your servers, 
# as well as the ACCLVLID of an empty (no doors) access level. 
# I hope it helps you as much as it helped me! 
#
# README - https://github.com/jmac5/Portfolio/blob/master/README.md
#
# Jim MacDonald - 2020 - Worcester Polytechnic Institute

####################### SERVER DECLARATIONS ######################

# Define the servers\instances you would like to use for your OnGuard environment here.
$prodServer = 'PROD_DB_SERVER\LENEL'
$testServer = 'TEST_DB_SERVER\LENEL'

# Here, add the ACCLVLID of the empty access level used to select the cardholders.
$emptyLevelID = '0000'


####################### FORCES X64 ###############################
if (($pshome -like "*syswow64*") -and ((Get-WmiObject Win32_OperatingSystem).OSArchitecture -like "64*")) {
    write-warning "Restarting script under 64 bit powershell"
 
    # relaunch this script under 64 bit shell
    & (join-path ($pshome -replace "syswow64", "sysnative")\powershell.exe) -file $myinvocation.mycommand.Definition @args
 
    # This will exit the original powershell process. This will only be done in case of an x86 process on a x64 OS.
    exit
}
##################### END FORCE X64 ##############################


[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
 
#$c = Get-Credential - only necessary if the current user has no privs to SQL Server

# .Net methods for hiding/showing the console in the background
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'

############################################# Create Window
$Form = New-Object System.Windows.Forms.Form
$Form.Size = New-Object System.Drawing.Size(500,400) 
$Form.text = "OnGuard List Building Tool"   
############################################# End Create Window



############################################## Start Add1

# Here, the queries are built, and the input is sanitized to resemble ID numbers. My site
# uses 9-digit IDs, but you can tailor your sanitizing settings to your site. 
function Add1 {

    $clist = @()
    
        if ($RadioButton2.Checked -and (!($RadioButton1.Checked)))  {
            $server = $prodServer
            }
        else{ 
            $server = $testServer
        }
    
        $list = $Inputbox.Lines.Where({ $_ -ne ""})             #populate the list here, Removes blank lines 
        $list2 = $list -replace '[^0-9]' | Where-Object {$_.trim() -ne "" } #Get rid of everything but numbers, blanks created by trim.
        
        foreach ($ID in $List2) {
            $ID = $ID.trimstart('0')
            $IDcheck = Invoke-Sqlcmd -ServerInstance $server -Query "IF EXISTS (SELECT * FROM Accesscontrol.dbo.BADGE WHERE ID ='$ID') BEGIN SELECT 1 END ELSE BEGIN SELECT 0 END " | Select-Object -ExpandProperty Column1
            
                if ($IDcheck -eq 0) {                           # 0 would mean the ID does not exist in the DB
                    $warning = ', Check list for Invalid IDs'
                } else {
                    $clist = $clist += $ID                      #add IDs to sub-list to become csvList
                }
        }
        $csvList = $clist -join ","                             # creates CSV of IDs for the SQL Query to use
        $OutputBox2.Text = $csvList                             #Prints output for review
        Invoke-Sqlcmd -ServerInstance $server -Query "INSERT INTO [accesscontrol].[dbo].[BADGELINK] (BADGEKEY,ACCLVLID) SELECT BADGEKEY, $emptyLevelID FROM [accesscontrol].[dbo].[BADGE] WHERE ID IN ($csvList)"
        $OutputBox.Text = "List Built Successfully" + $warning
    }
############################################## end Add1



############################################## Start Remove1
function Remove1 {
    
    if ($RadioButton2.Checked -and (!($RadioButton1.Checked)))  {
            $server = $prodServer
            }
    else{ 
            $server = $testServer
        }
        Invoke-Sqlcmd -ServerInstance $server -Query "DELETE FROM [accesscontrol].[dbo].BADGELINK WHERE ACCLVLID like $emptyLevelID" #1912 being the ACCLVID of Access Level "1"
        $OutputBox.Text = "Placeholder Access Level Removed from All Cardholders."
        $OutputBox2.Text = ""
        $InputBox.Text = ""
    }   
############################################### #end Remove1



##################################### Create the collection of radio buttons
$RadioButton1 = New-Object System.Windows.Forms.RadioButton
$RadioButton1.Location = '250,10'
$RadioButton1.size = '75,20'
$RadioButton1.Checked = $true 
$RadioButton1.Text = "TEST"
$RadioButton1.Cursor = [System.Windows.Forms.Cursors]::Hand

$RadioButton2 = New-Object System.Windows.Forms.RadioButton
$RadioButton2.Location = '330,10'
$RadioButton2.size = '100,20'
$RadioButton2.Checked = $false
$RadioButton2.Text = "PROD"
$RadioButton2.Cursor = [System.Windows.Forms.Cursors]::Hand
$form.Controls.AddRange(@($RadioButton1,$RadioButton2))
####################################### End collection of radio buttons



############################################## Text Box for List
$InputBox = New-Object System.Windows.Forms.TextBox 
$InputBox.Location = New-Object System.Drawing.Size(10,45) 
$InputBox.Size = New-Object System.Drawing.Size(200,175) 
$InputBox.MultiLine = $True
$InputBox.ScrollBars = "Vertical" 
$Form.Controls.Add($InputBox) 

$InputBoxlabel = New-Object System.Windows.Forms.Label
$InputBoxlabel.Location = New-Object System.Drawing.Point(9,30)
$InputBoxlabel.Size = New-Object System.Drawing.Size(280,20)
$InputBoxlabel.Text = 'Enter Cardholder IDs Below:'
$form.Controls.Add($InputBoxlabel)
############################################## end List Box



##################################### OutputBox for formatted text testing

# This field either displays error messages, or shows the 'clear' message
# when the level is cleared off cardholders. It remains blank on a successful
# attempt. 

$OutputBox = New-Object System.Windows.Forms.TextBox 
$OutputBox.Location = New-Object System.Drawing.Size(10,250) 
$OutputBox.Size = New-Object System.Drawing.Size(430,50) 
$OutputBox.MultiLine = $True
$Form.Controls.Add($OutputBox) 

$OutputBoxlabel = New-Object System.Windows.Forms.Label
$OutputBoxlabel.Location = New-Object System.Drawing.Point(9,235)
$OutputBoxlabel.Size = New-Object System.Drawing.Size(280,20)
$OutputBoxlabel.Text = 'Result:'
$form.Controls.Add($OutputBoxlabel)
############################################## end OutputBox



####################################### OutputBox for formatted text testing

# This field is just to check that the IDs you wanted were formatted
# into a CSV correctly. 

$OutputBox2 = New-Object System.Windows.Forms.TextBox 
$OutputBox2.Location = New-Object System.Drawing.Size(10,320) 
$OutputBox2.Size = New-Object System.Drawing.Size(430,30) 
$OutputBox2.MultiLine = $True
$Form.Controls.Add($OutputBox2) 

$OutputBox2label = New-Object System.Windows.Forms.Label
$OutputBox2label.Location = New-Object System.Drawing.Point(9,305)
$OutputBox2label.Size = New-Object System.Drawing.Size(280,20)
$OutputBox2label.Text = 'CSV Verification:'
$form.Controls.Add($OutputBox2label)
############################################## end OutputBox



############################################## Buttons
$Button = New-Object System.Windows.Forms.Button 
$Button.Location = New-Object System.Drawing.Size(300,50) 
$Button.Size = New-Object System.Drawing.Size(130,80) 
$Button.Text = "ADD PLACEHOLDER TO LIST" 
$Button.Add_Click({Add1})
$Button.Cursor = [System.Windows.Forms.Cursors]::Hand
$Button.BackColor = [System.Drawing.Color]::Red
$Button.Font = New-Object System.Drawing.Font("Verdana",10,[System.Drawing.FontStyle]::Bold) 
$Form.Controls.Add($Button) 

$Button2 = New-Object System.Windows.Forms.Button 
$Button2.Location = New-Object System.Drawing.Size(300,140) 
$Button2.Size = New-Object System.Drawing.Size(130,80) 
$Button2.Text = "REMOVE PLACEHOLDER FROM ALL" 
$Button2.Add_Click({Remove1})
$Button2.Cursor = [System.Windows.Forms.Cursors]::Hand
$Button2.BackColor = [System.Drawing.Color]::Cyan
$Button2.Font = New-Object System.Drawing.Font("Verdana",10,[System.Drawing.FontStyle]::Bold) 
$Form.Controls.Add($Button2) 
############################################## end buttons



############################################## Start Hide Console
function Hide-Console
{
    $consolePtr = [Console.Window]::GetConsoleWindow()
    #0 hide
    [Console.Window]::ShowWindow($consolePtr, 0)
}
############################################## end Hide Console



$Form.Add_Shown({$Form.Activate()})
Hide-Console
[void] $Form.ShowDialog()
