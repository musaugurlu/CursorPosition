$Script:CursorRepository = @()

function Get-Cursor
{
    [CmdletBinding(DefaultParameterSetName = "Id")]
    param (
        [Parameter(Mandatory = $false, ParameterSetName = "Id")]
        [int] $Id,
        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [string] $Name,
        [Parameter(Mandatory = $false, ParameterSetName = "List")]
        [switch] $List
    )
    
    begin
    {
        # Get current location
        $CurrentPosition = [System.Windows.Forms.Cursor]::Position
    }
    
    process
    {
        # create an instance of CursorPosition just to standardize the output
        $Position = New-CursorObject -X $CurrentPosition.X -Y $CurrentPosition.Y
    }
    
    end
    {
        if ($Id)
        {
            Write-Output ($Script:CursorRepository | Where-Object {$_.id -eq $Id})
        } elseif ($Name) 
        {
            Write-Output ($Script:CursorRepository | Where-Object {$_.Name -eq $Name})
        }
        elseif ($List) 
        {
            Write-Output $Script:CursorRepository
        }
        else 
        {
            Write-Output $Position
        }
    }
}

function Set-Cursor
{
    [CmdletBinding(DefaultParameterSetName = "Position")]
    param (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = "Id")]
        [Int] $id,
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = "Position")]
        [int] $X,
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = "Position")]
        [Int] $Y
    )
    
    begin
    {

    }
    
    process
    {
        # I usually assign variables in begin block. However, if X and Y info comes from pipeline, they would be processed in process block and i would see the info in begin block
        # Convert type of Int values to type of point, so the system can understant what all those numbers are for
        if ($id)
        {
            $Item = Get-Cursor -Id $id
            $Position = [System.Drawing.Point]::new($Item.X, $Item.Y)
        }
        else 
        {
            $Position = [System.Drawing.Point]::new($X, $Y)   
        }

        # set the cursor to new position provided
        [System.Windows.Forms.Cursor]::Position = $Position
    }
    
    end
    {
        Write-Verbose "Cursor position is set to $($Position.X)X$($Position.Y)"
    }
}

function Add-Cursor
{
    [CmdletBinding()]
    param (
        [string] $Name,
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [int] $X,
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [Int] $Y
    )
    
    begin
    {
    }
    
    process
    {
        if ($Name) 
        {
            $Cursor = New-CursorObject -Name $Name -X $X -Y $Y
        }
        else 
        {
            $Cursor = New-CursorObject -X $X -Y $Y
        }

        $Script:CursorRepository += $Cursor
    }
    
    end
    {
        Write-Output $Cursor
    }
}

function Remove-Cursor
{
    [CmdletBinding(DefaultParameterSetName = "Id")]
    param (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = "Id")]
        [int] $Id,
        [Parameter(Mandatory = $true, ParameterSetName = "List")]
        [switch] $All
    )
    
    begin
    {
    }
    
    process
    {
        if ($Id)
        {
            $Item = Get-Cursor -Id $Id
            $Script:CursorRepository = $Script:CursorRepository | Where-Object {$_ -ne $Item}   
        }
        elseif ($All) 
        {
            $Script:CursorRepository = @()    
        }
    }
    
    end
    {
        Write-Verbose "Cursor(s) have been removed."
    }
}

function Export-Cursor
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [psobject] $CursorList,
        [Parameter(Mandatory = $true)]
        [ValidateScript( {Test-Path $_})]
        [string] $Path,
        [string] $Name,
        [switch] $OpenLocation
    )
    
    begin
    {
        if ($Name)
        {
            if ($Name -match "([a-zA-Z0-9\s_\\.\-\(\):])+(.csv)$")
            {
                $FilePath = Join-Path -Path $Path -ChildPath $Name    
            }
            else
            {
                $FilePath = Join-Path -Path $Path -ChildPath "$Name.csv"
            }
        }
        else
        {
            $FilePath = "$(Get-Date -Format "yyyy-MM-dd_HH-mm-ss").csv"
        }
    }
    
    process
    {
        try
        {
            Get-Cursor -List | Export-Csv -Path $FilePath -NoTypeInformation -Force -Confirm:$false   
        }
        catch
        {
            Write-Error "Cursor List couldn't be exported."
        }
    }
    
    end
    {
        Write-Host "Cursor List has been exported to $FilePath"
        if ($OpenLocation)
        {
            Invoke-Item "$FilePath\.."
        }
    }
}

function Import-Cursor
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [ValidateScript({Test-Path $_})]
        [string] $Path,
        [switch] $Overwrite
    )
    
    begin
    {
    }
    
    process
    {
        if ($Overwrite)
        {
            $Script:CursorRepository = @()
            Write-Verbose "Cursor List has been freed up"
        }

        try {
            $Items = Import-Csv -Path $Path
            $Items | ForEach-Object {
                $Item = New-CursorObject -X $_.X -Y $_.Y -Name $_.Name -TimeStamp [datetime]($_.TimeStamp)
                $Script:CursorRepository += $Item
            }
        }
        catch {
            Write-Error "Some or all of the items couldn't be imported"        
        }
    }
    
    end
    {
        if($Error.Count -lt 1)
        {
            Write-Host "$Path has been imported."
        }
    }
}

##############################################################################################
## Helper Functions
##############################################################################################

function New-CursorObject {
    [CmdletBinding()]
    param (
        [int] $Id = (New-CursorID),
        [string] $Name = (Get-Date -Format "yyyy-MM-dd HH:mm:ss"),
        [int] $X,
        [int] $Y,
        [datetime] $TimeStamp = (Get-Date)
    )
    
    begin {
        $CursorObject = New-Object -TypeName psobject
    }
    
    process {
        $CursorObject | Add-Member -NotePropertyName "Id" -NotePropertyValue $Id
        $CursorObject | Add-Member -NotePropertyName "Name" -NotePropertyValue $Name
        $CursorObject | Add-Member -NotePropertyName "X" -NotePropertyValue $X
        $CursorObject | Add-Member -NotePropertyName "Y" -NotePropertyValue $Y
        $CursorObject | Add-Member -NotePropertyName "TimeStamp" -NotePropertyValue $TimeStamp
    }
    
    end {
        Write-Output $CursorObject
    }
}

function New-CursorID {
    [CmdletBinding()]
    param (
        
    )
    
    begin {
        [int]$i = 1
        [bool]$Loop = $true
    }
    
    process {
        function CheckId($id)
        {
            $check = $false
            for ($j = 0; $j -lt $Script:CursorRepository.Count; $j++) {
                if ($Script:CursorRepository[$j].id -eq $id) {
                    $check = $true
                }                
            }

            return [bool] $check
        }

        while ($Loop)
        {
            if (CheckId($i)) 
            {
                $i++
            }
            else
            {
                $Loop = $false
            }
        }
    }
    
    end {
        return [int] $i
    }
}