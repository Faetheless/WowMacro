<#
.SYNOPSIS
A PowerShell script to generate World of Warcraft macros.

.DESCRIPTION
The `macrowow.ps1` script creates customizable macros for World of Warcraft. It accepts a variety of parameters to specify the spell, target, and other options for the macro. Users can generate simple macros with just a spell cast or more complex macros that include targeting conditions, consumable usage, trinket activation, and pet commands. The script ensures that all macros are properly formatted and saved to the default directory `C:\WoW-Macros\` for easy access.

.PARAMETER spellName
The name of the spell to use in the macro. -example "Death Strike"

.PARAMETER Targeted
A boolean value indicating whether the spell should be targeted. /example: $true or $false

.PARAMETER trinketSlot
The slot number of the trinket to use. -example 13 or 14

.PARAMETER Consumabled
The name of the consumable to use. -example "Healthstone"

.PARAMETER targetType
The type of target for the spell. -example "harm" or "help"

.PARAMETER AtCursor
A boolean value indicating whether the spell should be cast at the cursor. /example: $true or $false

.PARAMETER deathStatus
The death status of the target. /example: dead, nodead

.PARAMETER petCommand
The command to give to the pet. /example: attack, follow, stay

.PARAMETER battleRez
A boolean value indicating whether to use a battle resurrection. example: $true or $false 

.PARAMETER ChatMessage
The chat message to send. example: "Hello World!"

.PARAMETER ChatCommand
The chat command to use. example: /say, /yell

.PARAMETER Simple
Modifies the macro to be a simple cast line. example: /cast spellName

.PARAMETER DirectoryPath
The directory path where the macro file will be saved. If not specified, the default directory is "C:\WoW-Macros\". -example: "D:\CustomMacros\"

.EXAMPLE
.\macrowow.ps1 -spellName "Death Strike" -Targeted $true -ChatMessage "Healing"
This example creates a macro for the spell "Death Strike" that targets a specific unit and includes a chat message "Healing".
Macro Output:
#showtooltip Death Strike
/stopcasting
/cast [@mouseover,exists,help,nodead] Death Strike; [@focus,help,nodead] Death Strike; Death Strike
/say Healing

.EXAMPLE
.\macrowow.ps1 -spellName "Fireball" -Simple $true
This example creates a simple macro for the spell "Fireball" that only includes the cast command without any additional targeting or conditions.
Macro Output:
#showtooltip Fireball
/cast Fireball

.EXAMPLE
.\macrowow.ps1 -spellName "Arcane Intellect" -Consumable "Mana Potion" -trinketSlot 13
This example creates a macro for the spell "Arcane Intellect" that also uses a "Mana Potion" and activates the trinket in slot 13.
Macro Output:
#showtooltip Arcane Intellect
/stopcasting
/use Mana Potion
/use 13
/cast [@mouseover,exists,help,nodead] Arcane Intellect; [@focus,help,nodead] Arcane Intellect; Arcane Intellect

.EXAMPLE
.\macrowow.ps1 -spellName "Revive Pet" -PetCommand "attack" -BattleRez $true
This example creates a macro for the spell "Revive Pet" that commands the pet to attack and includes battle resurrection conditions.
Macro Output:
#showtooltip Revive Pet
/stopcasting
/cast [nopet] Call Pet 1
/cast [nopet,@pet2] Call Pet 2
/petattack [@target,exists]
/cast [@mouseover,exists,help,dead] Revive Pet; [@focus,help,dead] Revive Pet; Revive Pet

.NOTES
Ensure that the spell names and other parameters are valid in the current version of World of Warcraft. 
Author: Faetheless
Date: 2024-10-12
Version: 1.2
This script is intended for use with World of Warcraft and may need updates for compatibility with future game patches.
#>
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
    [bool]$Simple = $false,

    [Parameter(Mandatory=$false)]
    [bool]$stopcast = $false,

    [Parameter(Mandatory=$false)]
    [string]$DirectoryPath = "C:\WoW-Macros\"  # Default directory path
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
        [bool]$stopcast,
        [bool]$Simple
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
        $Targeted = $false
    }
    
    if ($Simple -eq $false) {
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
        Write-Verbose "Simple Cast Line"
        $macroText += "`n/cast $spellName"
        $Targeted = $false
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
    
    if ($Targeted) {
        Write-Verbose "Appending the targeting line"
        $macroText += "`n/target [@mouseover,$TargetType,nodead]"
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
        [string]$spellName,
        [string]$DirectoryPath
    )

    # Sanitize the file Name
    $SanSpell = Convert-FilePath -UnsanitizedSpell $spellName

    # Specify the file path
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

$macroText = New-Macro -spellName $spellName -Targeted $Targeted -trinketSlot $TrinketSlot -Consumable $Consumable -targetType $TargetType -AtCursor $AtCursor -deathStatus $DeathStatus -petCommand $PetCommand -battleRez $BattleRez -ChatMessage $ChatMessage -ChatCommand $ChatCommand -AtPlayer $AtPlayer -Simple $Simple -stopcast $stopcast
Save-MacroToFile -macroText $macroText -spellName $spellName -DirectoryPath $DirectoryPath
Write-Host $macroText
