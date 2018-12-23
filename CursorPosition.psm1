$Script:CursorRepository = @()

# we are creating this class to standardize to object because we are going to use one function's output in the other function.
class CursorPosition
{
    [int] $Id
    [string] $Name
    [int] $X
    [int] $Y
    [datetime] $TimeStamp

    CursorPosition([int]$X, [int]$Y)
    {
        $this.Id = $this.GetNewId()
        $this.Name = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        $this.X = $X
        $this.Y = $Y
        $this.TimeStamp = Get-Date
    }

    CursorPosition([int]$X, [int]$Y, [string]$Name)
    {
        $this.Id = $this.GetNewId()
        $this.Name = $Name
        $this.X = $X
        $this.Y = $Y
        $this.TimeStamp = Get-Date
    }

    CursorPosition([int]$X, [int]$Y, [string]$Name, [datetime]$TimeStamp)
    {
        $this.Id = $this.GetNewId()
        $this.Name = $Name
        $this.X = $X
        $this.Y = $Y
        $this.TimeStamp = $TimeStamp
    }

    [int] GetNewId()
    {
        [int]$i = 1
        [bool]$Loop = $true
        
        while ($Loop)
        {
            if ($Script:CursorRepository.Id -contains $i) 
            {
                $i++
            }
            else
            {
                $Loop = $false
            }
        }
        return $i
    }
}
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
        $Position = [CursorPosition]::new($CurrentPosition.X, $CurrentPosition.Y)
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
            $Cursor = [CursorPosition]::new($X, $Y, $Name)
        }
        else 
        {
            $Cursor = [CursorPosition]::new($X, $Y)
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
        [CursorPosition] $CursorList,
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
                $Item = [CursorPosition]::new($_.X,$_.Y,$_.Name,[datetime]$_.TimeStamp)
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