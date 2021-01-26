# INFX @ <Evig nedtur> - Firemaw EU
# Use with addon UnitScan
# UnitScan must be in top-left corner (hold control and click to drag it all the way into the corner)

#############
# FUNCTIONS #
#############

function Get-ScreenColor {

    [CmdletBinding(DefaultParameterSetName='None')]

    param(
        [Parameter(
            Mandatory=$true,
            ParameterSetName="Pos"
        )]
        [Int]
        $X,
        [Parameter(
            Mandatory=$true,
            ParameterSetName="Pos"
        )]
        [Int]
        $Y
    )
    Add-Type -Assembly system.drawing
    if ($PSCmdlet.ParameterSetName -eq 'None') {
        $pos = [System.Windows.Forms.Cursor]::Position
    } else {
        $pos = New-Object psobject
        $pos | Add-Member -MemberType NoteProperty -Name "X" -Value $X
        $pos | Add-Member -MemberType NoteProperty -Name "Y" -Value $Y
    }
    $map = [System.Drawing.Rectangle]::FromLTRB($pos.X, $pos.Y, $pos.X + 1, $pos.Y + 1)
    $bmp = New-Object System.Drawing.Bitmap(1,1)
    $graphics = [System.Drawing.Graphics]::FromImage($bmp)
    $graphics.CopyFromScreen($map.Location, [System.Drawing.Point]::Empty, $map.Size)
    $pixel = $bmp.GetPixel(0,0)
    $red = $pixel.R
    $green = $pixel.G
    $blue = $pixel.B
    $result = New-Object psobject
    if ($PSCmdlet.ParameterSetName -eq 'None') {
        $result | Add-Member -MemberType NoteProperty -Name "X" -Value $([System.Windows.Forms.Cursor]::Position).X
        $result | Add-Member -MemberType NoteProperty -Name "Y" -Value $([System.Windows.Forms.Cursor]::Position).Y
    }
    $result | Add-Member -MemberType NoteProperty -Name "Red" -Value $red
    $result | Add-Member -MemberType NoteProperty -Name "Green" -Value $green
    $result | Add-Member -MemberType NoteProperty -Name "Blue" -Value $blue
    return $result
}

#get API key from api.ocr.space and edit that out here
function Upload-Screenshot {
PARAM (
    [string]$apiKey = "1234567890",
    [string]$image = "C:\temp\text.png"
)

    #POWERSHELL OCR API CALL - V2.0, May 30, 2020
    #In this demo we send an image link to the OCR API and download the text result and the searchable PDF

    #Enter your api key here
    $apiUrl = "https://api.ocr.space/parse/image" 

    #Call API with CURL
    $shutUp = curl.exe -X POST $apiurl -H "apikey:$apikey" -F "file=@$image" -F "language=eng" -F "isOverlayRequired=false" -F "iscreatesearchablepdf=true" | ConvertFrom-Json -OutVariable response

    #Done, write OCR'ed text to log file
    $text = $response.ParsedResults.ParsedText
    Write-Host $text

    $response = $response | select SearchablePDFURL

    return $text
}

function Get-ScreenCapture {
    Add-Type -AssemblyName System.Windows.Forms,System.Drawing

    $screens = [Windows.Forms.Screen]::AllScreens  | Where-Object {$_.Primary -eq "True"}

    $top    = ($screens.Bounds.Top    | Measure-Object -Minimum).Minimum
    $left   = ($screens.Bounds.Left   | Measure-Object -Minimum).Minimum
    $width  = 160
    $height = 80

    $bounds   = [Drawing.Rectangle]::FromLTRB($left, $top, $width, $height)
    $bmp      = New-Object System.Drawing.Bitmap ([int]$bounds.width), ([int]$bounds.height)
    $graphics = [Drawing.Graphics]::FromImage($bmp)

    $graphics.CopyFromScreen($bounds.Location, [Drawing.Point]::Empty, $bounds.size)

    if(!(Test-Path "C:\Temp")) {
        New-Item -ItemType Directory C:\Temp -ErrorAction SilentlyContinue
    }
    $bmp.Save("C:\temp\text.png")

    $graphics.Dispose()
    $bmp.Dispose()
}

function Start-Scouting {

PARAM (
    [switch]$Azuregos,
    [switch]$NMD,
    [switch]$Kazzak,
    [string]$RefreshSeconds = 15,
    [string]$Layer,
    [string]$Player,
    [string]$Secret,
    [switch]$Alert,
    [switch]$ScoutingInformation

)
    #Scouting Information
    if($ScoutingInformation) { 
        #Store webhook url pointing to discord
        $webHookUrl = ""

        #Store embed values
        if($Kazzak) {
            $content = "[+] **$Player** has started scouting **Kazzak** on Layer $Layer"
        }
        if($NMD) {
            $content = "[+] **$Player** has started scouting **NMD** on Layer $Layer"
        }
        if($Azuregos) {
            $content = "[+] **$Player** has started scouting **Azuregos** on Layer $Layer"
        }

        #Create the payload
        $payload = [PSCustomObject]@{

            #embeds = $embedArray
            content = $content
        }
        #Send over payload, converting it to JSON
        Invoke-RestMethod -Uri $webHookUrl -Body ($payload | ConvertTo-Json -Depth 4) -Method Post -ContentType 'application/json'
    }
    
    $stopLoop = $false
    Do {
        #These are the coordinates the script looks for popups
        #RGB values for Unitscan (3 points to ensure less wrong reporting)

        #Create an array to hold sums
        $sumArray = @()

        #Scan upper left area on Y 20, X range from 20-100 (we know this is in range)
        $x = @()
        $y = @()
        70..100 | ForEach-Object { $x += $_ }
        11..16 | ForEach-Object { $y += $_ }
        foreach($n in $x) {
            foreach($i in $y) {
                $color = Get-ScreenColor -X $n -Y $i
                $colorSum = $color.Red + $color.Blue + $color.Green
                $sumArray += $colorSum
            }
        }
        $sortedArray = $sumArray | Sort-Object
        if($sortedArray[0] -eq 0 -and ($sortedArray[1] -eq 0) -and ($sortedArray[185] -eq 765) -and ($sortedArray[184] -eq 765) -and ($sortedArray[183] -eq 765) -and ($sortedArray[182] -eq 765)) {
            Get-ScreenCapture
            Write-Host "POPUP DETECTED - CHECKING WITH API"
            $boss = Upload-Screenshot -image "C:\temp\TradeChatScouting.png"
        } else {
            Write-Host "[+] Still scanning..." -ForegroundColor DarkGreen
        }

        if($stopLoop -eq $false) {
          if($Kazzak) {
          #This is the part that looks for Kazzak
                #The red color needs to be above 100 (which in hindsight is a shit metric and can be tuned better)
                if($boss -like "LORD KAZZA*"){
                    if($Alert) {
                        $description = "**KAZZAK UP** -- **LAYER $Layer** -- @everyone -- /w $Player $Secret"
                        $stopLoop = $true
                    } else {
                        $description = "**KAZZAK UP** -- **LAYER $Layer** -- /w $Player $Secret"
                        $stopLoop = $true
                    }  
                } else {
                    #Sleeps for 15 seconds if it doesn't find anything
                    Start-Sleep $RefreshSeconds
                }
            }
            if($Azuregos) {
            #This is the part that looks for Azu
                #The blue color needs to be above 200 (which in hindsight is a shit metric and can be tuned better)
                if($boss -like "AZUREGOS*"){
                    if($Alert) {
                        $description = "**AZUREGOS UP** -- **LAYER $Layer** -- @everyone -- /w $Player $Secret"
                        $stopLoop = $true
                    } else {
                        $description = "**AZUREGOS UP** -- **LAYER $Layer** -- /w $Player $Secret"
                        $stopLoop = $true
                    }
                } else {
                    #Sleeps for 15 seconds if it doesn't find anything
                    Start-Sleep $RefreshSeconds
                }
            }
            if($NMD){
            #This is the part that looks for NMD
                #The green color needs to be above 220 (which in hindsight is a shit metric and can be tuned better)
                if($boss -like "YSONDRE*" -or ($boss -like "TAERAR*") -or ($boss -like "EMERISS*") -or ($boss -like "LETHON*")){
                    if($Alert) {
                        $description = "**NMD UP ($boss)** -- **LAYER $Layer** -- @everyone -- /w $Player $Secret"
                        $stopLoop = $true
                    } else {
                        $description = "**NMD UP ($boss)** -- **LAYER $Layer** -- /w $Player $Secret"
                        $stopLoop = $true
                    }
                } else {
                    #Sleeps for 15 seconds if it doesn't find anything
                    Start-Sleep $RefreshSeconds
                }
            }
        }
        } Until ($stopLoop -eq $true)

    # Embed with title, description, color, and a thumbnail

    #Store webhook url
    $webHookUrl = ""

    #Store embed values
    $content = "[+] **World Boss:** "
    $content = $content + " " + $description

    #Create the payload
    $payload = [PSCustomObject]@{

        #embeds = $embedArray
        content = $content
    }
    #Send over payload, converting it to JSON
    Invoke-RestMethod -Uri $webHookUrl -Body ($payload | ConvertTo-Json -Depth 4) -Method Post -ContentType 'application/json'
}

#function for testing
function Get-MouseCoordinates {
    Add-Type -AssemblyName System.Windows.Forms
    Start-Sleep 3
    $X = [System.Windows.Forms.Cursor]::Position.X
    $Y = [System.Windows.Forms.Cursor]::Position.Y
    Write-Output "X: $X | Y: $Y"
}

################
#   UI CODE    #
################

Add-Type -AssemblyName PresentationFramework
add-type -Assembly System.Drawing

[xml]$xaml = @'
<Window
xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
Title="Trade Chat Coalition Alarm System" WindowStartupLocation="CenterScreen" Height="400" Width="450" Background="#2f80ed">
    <Grid Margin="0,5,0,0">
        <Grid.RowDefinitions> 
            <RowDefinition Height="auto"/>
            <RowDefinition Height="auto"/>
            <RowDefinition Height="auto"/>
            <RowDefinition Height="auto"/>
            <RowDefinition Height="auto"/>
            <RowDefinition Height="auto"/>
            <RowDefinition Height="auto"/>
            <RowDefinition Height="auto"/>
            <RowDefinition Height="auto"/>
            <RowDefinition Height="auto"/>
            <RowDefinition Height="auto"/>
            <RowDefinition Height="auto"/>
        </Grid.RowDefinitions>
            <Label Grid.Row="0" Name="Character_Label" Content="Type character name" Foreground="#ffc0cb"/>
            <TextBox Grid.Row="1" Name="Character_Input"/>
            <Label Grid.Row="2" Name="Invite_Label" Content="Type autoinvite secret" Foreground="#ffc0cb"/>
            <TextBox Grid.Row="3" Name="Invite_Input"/>
            <Label Grid.Row="4" Name="Boss_Label" Content="Select BOSS" Foreground="#ffc0cb"/>
            <ComboBox Grid.Row="5" Name="Boss"/>
            <Label Grid.Row="6" Name="Layer_Label" Content="Select LAYER" Foreground="#ffc0cb"/>
            <ComboBox Grid.Row="7" Name="Layer"/>
            <Label Grid.Row="8" Name="Alert_Label" Content="Tag @everyone" Foreground="#ffc0cb"/>
            <ComboBox Grid.Row="9" Name="Alert"/>
            <Button Grid.Row="10" Name="startScouting" Content="Start Scouting" Background="#faffa7" Foreground="#fd625e" Height="40"/>
            <Label Grid.Row="12" Name="Feedback_Label" Content="Current status: INACTIVE" Foreground="#333333"/>
    </Grid>
</Window>
'@

$reader = (New-Object System.Xml.XmlNodeReader $xaml)
try {
    $window = [Windows.Markup.XamlReader]::Load( $reader )
}
catch {
    Write-Warning $_.Exception
    throw
}

#Create variables based on form control names.
#Variable will be named as 'var_<control name>'

$xaml.SelectNodes("//*[@Name]") | ForEach-Object {
    #"trying item $($_.Name)";
    try {
        Set-Variable -Name "var_$($_.Name)" -Value $window.FindName($_.Name) -ErrorAction Stop
    } catch {
        throw
   }
}

#This is for dev purposes
#Get-Variable var_*

#Add bosses to Combobox
$var_Boss.Items.add("Azuregos")
$var_Boss.Items.add("NMD")
$var_Boss.Items.add("Kazzak")

#Add layers to Combobox
$var_Layer.Items.add("1")
$var_Layer.Items.add("2")

#Add alert options to Combobox
$var_Alert.Items.Add("Yes")
$var_Alert.Items.Add("No")

$window.Add_Loaded({
    $Window.Icon = "https://cdn.discordapp.com/attachments/634238750964187147/727995811488596018/goose_2.jpg"
})

$var_startScouting.Add_Click( {
        If($var_Boss.SelectedItem -eq "Azuregos") {
            $var_Feedback_Label.Content = "Current status: SCOUTING AZUREGOS"
            $var_Feedback_Label.Foreground = "#00FF66"
            Write-Host "[+] ----------------" -ForegroundColor DarkGreen
            Write-Host "[+] Scouting started" -ForegroundColor DarkGreen
            Write-Host "[+] Main window will freeze" -ForegroundColor DarkGreen
            Write-Host "[+] This is working as intended" -ForegroundColor DarkGreen
            Write-Host "[+] ----------------" -ForegroundColor DarkGreen
            Update-Gui
            if($var_Alert.SelectedItem -eq "Yes") {
                Start-Scouting -Azuregos -Layer $var_Layer.SelectedItem -Player $var_Character_Input.Text -Secret $var_Invite_Input.Text -Alert -ScoutingInformation
                $var_Feedback_Label.Content = "Current status: AZUREGOS FOUND"
                $var_Feedback_Label.Foreground = "#00FF66"
            } else {
                Start-Scouting -Azuregos -Layer $var_Layer.SelectedItem -Player $var_Character_Input.Text -Secret $var_Invite_Input.Text -ScoutingInformation
                $var_Feedback_Label.Content = "Current status: AZUREGOS FOUND"
                $var_Feedback_Label.Foreground = "#FD625E"
            }
        }
        If($var_Boss.SelectedItem -eq "NMD") {
            $var_Feedback_Label.Content = "Current status: SCOUTING NMD"
            $var_Feedback_Label.Foreground = "#00FF66"
            Write-Host "[+] ----------------" -ForegroundColor DarkGreen
            Write-Host "[+] Scouting started" -ForegroundColor DarkGreen
            Write-Host "[+] Main window will freeze" -ForegroundColor DarkGreen
            Write-Host "[+] This is working as intended" -ForegroundColor DarkGreen
            Write-Host "[+] ----------------" -ForegroundColor DarkGreen
            Update-Gui
            if($var_Alert.SelectedItem -eq "Yes") {
                Start-Scouting -NMD -Layer $var_Layer.SelectedItem -Player $var_Character_Input.Text -Secret $var_Invite_Input.Text -Alert -ScoutingInformation
                $var_Feedback_Label.Content = "Current status: NMD FOUND"
                $var_Feedback_Label.Foreground = "#FD625E"
            } else {
                Start-Scouting -NMD -Layer $var_Layer.SelectedItem -Player $var_Character_Input.Text -Secret $var_Invite_Input.Text -ScoutingInformation
                $var_Feedback_Label.Content = "Current status: NMD FOUND"
                $var_Feedback_Label.Foreground = "#FD625E"
            }
        }
        If($var_Boss.SelectedItem -eq "Kazzak") {
            $var_Feedback_Label.Content = "Current status: SCOUTING KAZZAK"
            $var_Feedback_Label.Foreground = "#00FF66"
            Write-Host "[+] ----------------" -ForegroundColor DarkGreen
            Write-Host "[+] Scouting started" -ForegroundColor DarkGreen
            Write-Host "[+] Main window will freeze" -ForegroundColor DarkGreen
            Write-Host "[+] This is working as intended" -ForegroundColor DarkGreen
            Write-Host "[+] ----------------" -ForegroundColor DarkGreen
            Update-Gui
            if($var_Alert.SelectedItem -eq "Yes") {
                Start-Scouting -Kazzak -Layer $var_Layer.SelectedItem -Player $var_Character_Input.Text -Secret $var_Invite_Input.Text -Alert -ScoutingInformation
                $var_Feedback_Label.Content = "Current status: KAZZAK FOUND"
                $var_Feedback_Label.Foreground = "#FD625E"
            } else {
                Start-Scouting -Kazzak -Layer $var_Layer.SelectedItem -Player $var_Character_Input.Text -Secret $var_Invite_Input.Text -ScoutingInformation
                $var_Feedback_Label.Content = "Current status: KAZZAK FOUND"
                $var_Feedback_Label.Foreground = "#FD625E"
            }
        }
   })

function Update-Gui {
    # Basically WinForms Application.DoEvents()
    $Window.Dispatcher.Invoke([Windows.Threading.DispatcherPriority]::Background, [action]{})
}
$Null = $window.ShowDialog()
