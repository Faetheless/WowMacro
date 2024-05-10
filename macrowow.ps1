# Make Verbose Work
[CmdletBinding()]
param (
    [Parameter(Mandatory=$true)]
    [string]$SpellName,

    [Parameter(Mandatory=$false)]
    [bool]$Targeted = $true,

    [Parameter(Mandatory = $false)]
    [ValidateSet(13, 14)]  # Ensure only valid trinket slots are accepted
    [int[]]$TrinketSlot = @(),

    [Parameter(Mandatory=$false)]
    [string[]]$Consumable = @(),

    [Parameter(Mandatory=$false)]
    [ValidateSet("harm", "help")]  # Ensure only valid target types are accepted
    [string]$TargetType = "harm",

    [Parameter(Mandatory=$false)]
    [bool]$AtCursor = $false,

    [Parameter(Mandatory=$false)]
    [bool]$BattleRez = $false,

    [Parameter(Mandatory=$false)]
    [ValidateSet("dead", "nodead")]  # Ensure only valid death statuses are accepted
    [string]$DeathStatus = "nodead",

    [Parameter(Mandatory=$false)]
    [ValidateSet("attack", "follow", "stay")]  # Ensure only valid pet commands are accepted
    [string]$PetCommand,

    [Parameter(Mandatory=$false)]
    [string]$ChatMessage,

    [Parameter(Mandatory=$false)]
    [ValidateSet("say", "yell")]  # Ensure only valid chat commands are accepted
    [string]$ChatCommand = "say",

    [Parameter(Mandatory=$false)]
    [bool]$AtPlayer = $false,

    [Parameter(Mandatory=$false)]
    [bool]$BigCD = $false,

    [Parameter(Mandatory=$false)]
    [bool]$stopcast = $false

)

# Function to create the macro based on user inputs
function New-Macro {
    param (
        [string]$spellName,
        [bool]$Targeted,
        [int[]]$trinketSlot,
        [string[]]$Consumable,
        [string]$TargetType,
        [bool]$AtCursor,
        [string]$DeathStatus,
        [string]$PetCommand,
        [string]$ChatMessage,
        [string]$ChatCommand,
        [bool]$AtPlayer,
        [bool]$BigCD,
        [bool]$stopcast
    )

    Write-Verbose "Starting $spellName Macro with a Tooltip"
    $macroText = "#showtooltip $spellName"

    if ($stopcast -eq $true) {
        Write-Verbose "Stopcast is true, adding stopcasting line"
        $macroText += "`n/stopcasting"
    }

    foreach ($item in $Consumable) {
        Write-Verbose "Appending line for $item"
        $macroText += "`n/use $item"
    }
        
    foreach ($slot in $trinketSlot) {
        Write-Verbose "Appending line for trinket slot $slot"
        if ($slot -eq 13 -or $slot -eq 14) {
            $macroText += "`n/use $slot"
        }
    }

    if ($PetCommand -eq "attack") {
        write-verbose "pet command is attack, adding pet attack line"
        $macroText += "`n/cast [nopet] Call Pet 1`n/cast [nopet,@pet2] Call Pet 2`n/petattack [@target,exists]"
    } elseif ($PetCommand) {
        write-verbose "pet command is $PetCommand, adding pet $PetCommand line"
        $macroText += "`n/cast [nopet] Call Pet 1`n/cast [nopet,@pet2] Call Pet 2`n/pet$PetCommand"
    } 

    if ($BattleRez -eq $true) {
        Write-Verbose "BattleRez is true, modifying the target type and death status"
        $TargetType = "help"
        $DeathStatus = "dead"
    }
    
    if ($BigCD -eq $false) {
        if ($AtCursor -eq $false) {
            Write-Verbose "Main Cast Line"
            if ($AtPlayer -eq $false) {
                Write-Verbose "AtPlayer is $AtPlayer, adding targeting conditions"
                $macroText += "`n/cast [@mouseover,exists,$TargetType,$DeathStatus] $spellName; [@focus,$TargetType,$DeathStatus] $spellName; $spellName"
            } else {
                Write-Verbose "AtPlayer is $AtPlayer, casting at player"
                $macroText += "`n/cast [@mouseover,exists,$TargetType,$DeathStatus] $spellName; [@focus,$TargetType,$DeathStatus] $spellName; [@player] $spellName"
            }
        } else {
            Write-Verbose "Main Cast Line at cursor"
            $macroText += "`n/cast [@cursor] $spellName"
            $Targeted = $false
        }
    } else {
        Write-Verbose "Big CD Line"
        $macroText += "`n/cast $spellName"
        $Targeted = $false
    }

    if ($Targeted) {
        Write-Verbose "Appending the targeting line"
        $macroText += "`n/target [@mouseover,$TargetType,nodead]"
    }

    if ($PetCommand -eq "attack") {
        Write-Verbose "Appending pet attack lines"
        $macroText += "`n/cast [@pettarget]Claw"
        $macroText += "`n/cast [@pettarget]Bite"
        $macroText += "`n/cast [@pettarget]Smack"
    } 

    if ($ChatMessage) {
        Write-Verbose "chat message is specified, appending it to the macro"
        $macroText += "`n/$ChatCommand $ChatMessage"
    }

    return $macroText
}
function Convert-FilePath {
    param (
        [string]$UnsanitizedSpell
    )

    # List of invalid characters in a file path
    $invalidChars = '[<>:"/\\|?*]'

    # Replace invalid characters with an underscore
    $sanitizedFilePath = $UnsanitizedSpell -replace $invalidChars, '_'

    return $sanitizedFilePath
}

# Function to prompt for spell name
function Get-SpellName {
    return Read-Host "Enter the spell name"
}

# Function to save the macro to a text file
function Save-MacroToFile {
    param (
        [string]$macroText,
        [string]$spellName
    )

    # Sanitize the file Name
    $SanSpell = Convert-FilePath -UnsanitizedSpell $spellName

    # Specify the file path
    $directoryPath = "C:\WoW-Macros\"
    $filePath = "$directoryPath$SanSpell-Macro.txt"


    # Check if the directory exists, if not, create it
    if (-Not (Test-Path -Path $directoryPath)) {
        Write-Verbose "Creating directory: $directoryPath"
        New-Item -Path $directoryPath -ItemType Directory
    }

    write-verbose "Saving macro to: $filePath"
    $macroText | Out-File -FilePath $filePath

    Write-Host "Macro saved to: $filePath"
}

# Main Script Logic
if (-not $spellName) {
    write-verbose "Spell name not provided, prompting user for input"
    $spellName = Get-SpellName
}

$macroText = New-Macro -spellName $spellName -Targeted $Targeted -trinketSlot $TrinketSlot -Consumable $Consumable -targetType $TargetType -AtCursor $AtCursor -deathStatus $DeathStatus -petCommand $PetCommand -battleRez $BattleRez -ChatMessage $ChatMessage -ChatCommand $ChatCommand -AtPlayer $AtPlayer -BigCD $BigCD -stopcast $stopcast
Save-MacroToFile -macroText $macroText -spellName $spellName
Write-Host $macroText
