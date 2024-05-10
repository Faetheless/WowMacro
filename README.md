# WowMacro

Script to generate macros for World of Warcraft.

## Parameters

The script accepts several parameters to customize the macro:

- **`-spellName`**: The name of the spell to use in the macro.
- **`-Targeted`**: Whether the spell should be targeted.
- **`-trinketSlot`**: The slot number of the trinket to use.
- **`-Consumable`**: The name of the consumable to use.
- **`-targetType`**: The type of target for the spell.
- **`-AtCursor`**: Whether the spell should be cast at the cursor.
- **`-deathStatus`**: The death status of the target.
- **`-petCommand`**: The command to give to the pet.
- **`-battleRez`**: Whether to use a battle resurrection.
- **`-ChatMessage`**: The chat message to send.
- **`-ChatCommand`**: The chat command to use.

## Example Usage

To create a macro with a spell named "Heal", a targeted spell, and a chat message, you would enter:

```powershell
.\wowmacro.ps1 -SpellName "Heal" -Targeted $true -ChatMessage "Healing incoming!"
```

 For another example of an acceptable command:
 ```powershell
.\wowmacro.ps1 -SpellName "Incarnation: Guardian of Ursoc" -BigCD $true -Consumable "Elemental Potion of Ultimate Power" -ChatCommand "yell" -ChatMessage "BEAR DOWN FOR MIDTERMS"
```
If you encounter an error about a parameter being specified more than once, ensure no duplicate parameters exist in your command.

## Troubleshooting

If you're having trouble, you can use the -Verbose parameter. Remember that it should be used when calling the script, not within the script itself.

```powershell
.\wowmacro.ps1 -Verbose
```

