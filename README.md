# Cursor Position

## Installation
You may install the module from [PowerShell Gallery](https://www.powershellgallery.com/packages/CursorPosition).

`Install-Module -Name CursorPosition`

or

`Install-Module -Name CursorPosition -Scope CurrentUser` # for current user only.


## Examples
* `Get-Cursor`  # Returns current cursor position
* `Get-Cursor -List` # Returns the list of saved cursor position
* `Get-Cursor | Add-Cursor -Name "My favorite Position"` # Returns and Saves Current position with a the name.
* `Get-Cursor | Add-Cursor | Out-Null` # Saves current position, doesn't return anything
* `Get-Cursor -Id 2` # Returns the position with Id of 2
* `Get-Cursor -Name "My favorite Position"` # Returns the position with the name
* `Set-Cursor -X 125 -Y 250` # Sets the cursor to (125,250)
* `Get-Cursor -id 2 | Set-Cursor`  # Sets the cursor to the position with id of 2
* `Get-Cursor -Name "My favorite Position" | Set-Cursor` # Sets the cursor the position
* `Add-Cursor -X 125 -Y 250 -Name "X Button"` # Saves the cursor with the name
* `Remove-Cursor -id 2` # Removes the position from list
* `Remove-Cursor -All`  #Clears the list
* `Get-Cursor -Id 2 | Remove-Cursor`  # Removes the position with Id of 2 from the list
* `Get-Cursor -List | Export-Cursor -Path .\ -Name "My Positions"` # Exports the list to a CSV file with the name of "My Positions.csv" in the current directory. 
* `Get-Cursor -Name "My Favorite" | Export-Cursor -Path .\ -OpenLocation` # Exports the position with the id of 2 to a CSV file with the name of "Year-Month-Day_Hour-Minute-Secons.csv" in the current location and opens the folder.
* `Import-Cursor -Path ".\My Positions.csv"` # imports positions from list and appends them to the current list
* `Import-Cursor -Path ".\My Positions.csv" -Overwrite` # imports positions from list and overwrites the current lists.


## Todo
* Switch to PSobjects from Classes to make it compatible to Version < 5
* Find an alternative to "$Script:CursorRepository.Id" as it doesn't work in version < 5
* Implement [xdotool](https://github.com/jordansissel/xdotool) to make it available in Linux and Mac OS 