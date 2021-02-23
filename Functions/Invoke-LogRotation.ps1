#Requires -Version 5.1

<#
.SYNOPSIS
Compress or delete old files and folders according to your requirements.
Performs log rotation with only a single Cmdlet.

.DESCRIPTION
Compress or delete old files and folders according to your requirements.
Performs log rotation with only a single Cmdlet.
 - You can select properties from 'CreationTime', 'LastWriteTime' and 'LastAccessTime'. The default parameter is 'CreationTime'.
 - You can select the time unit from Year, Month, Day, Hour, Minute, Second, and Millisecond.

.EXAMPLE
Invoke-LogRotation -Path C:\Temp\test-1, C:\Temp\test-2 -Property CreationTime -Unit Month -Compression 6 -Deletion 9

.EXAMPLE
'C:\Work\test-1', 'C:\Work\test-2' | Invoke-LogRotation -Property LastWriteTime -Unit Month -Compression 6 -Deletion 9

.EXAMPLE
Invoke-LogRotation -Path C:\Temp\test-1, C:\Temp\test-2 -Property CreationTime -Unit Day -Compression 180 -Deletion 365

.EXAMPLE
'C:\Work\test-1', 'C:\Work\test-2' | Invoke-LogRotation -Property LastWriteTime -Unit Year -Compression 1 -Deletion 2

.NOTES
Author: nekrassov01
#>

function Invoke-LogRotation
{
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    [OutputType([psobject[]])]
    param
    (
        [Parameter(Position = 0, Mandatory = $false, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [System.IO.FileInfo[]]$Path,

        [Parameter(Position = 1, Mandatory = $false)]
        [ValidateSet('CreationTime', 'LastWriteTime', 'LastAccessTime')]
        [string]$Property = 'CreationTime',

        [Parameter(Position = 2, Mandatory = $true)]
        [ValidateSet('Year', 'Month', 'Day', 'Hour', 'Minute', 'Second', 'Millisecond')]
        [string]$Unit,

        [Parameter(Position = 3, Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [int]$Compression,

        [Parameter(Position = 4, Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [int]$Deletion,

        [Parameter(Position = 5, Mandatory = $false)]
        [System.IO.Compression.CompressionLevel]$CompressionLevel = 'Optimal'
    )

    begin
    {
        Set-StrictMode -Version Latest
    }

    process
    {
        try
        {
            foreach ($p in $path)
            {
                Get-ChildItem -LiteralPath $p | ForEach-Object -Process {

                    switch ($PSBoundParameters['Unit'])
                    {
                        'Year'
                        { 
                            $removeScript   = [scriptblock]{ Remove-PreviousItem   -Path $_.FullName -Property $property -Year $deletion -Confirm:$false -WhatIf:$false }
                            $compressScript = [scriptblock]{ Compress-PreviousItem -Path $_.FullName -Property $property -Year $compression -CompressionLevel $compressionLevel -RemoveAfterCompression -Confirm:$false -WhatIf:$false }
                        }
                        'Month'
                        {
                            $removeScript   = [scriptblock]{ Remove-PreviousItem   -Path $_.FullName -Property $property -Month $deletion -Confirm:$false -WhatIf:$false }
                            $compressScript = [scriptblock]{ Compress-PreviousItem -Path $_.FullName -Property $property -Month $compression -CompressionLevel $compressionLevel -RemoveAfterCompression -Confirm:$false -WhatIf:$false }
                        }
                        'Day'
                        {
                            $removeScript   = [scriptblock]{ Remove-PreviousItem   -Path $_.FullName -Property $property -Day $deletion -Confirm:$false -WhatIf:$false }
                            $compressScript = [scriptblock]{ Compress-PreviousItem -Path $_.FullName -Property $property -Day $compression -CompressionLevel $compressionLevel -RemoveAfterCompression -Confirm:$false -WhatIf:$false }
                        }
                        'Hour'
                        {
                            $removeScript   = [scriptblock]{ Remove-PreviousItem   -Path $_.FullName -Property $property -Hour $deletion -Confirm:$false -WhatIf:$false }
                            $compressScript = [scriptblock]{ Compress-PreviousItem -Path $_.FullName -Property $property -Hour $compression -CompressionLevel $compressionLevel -RemoveAfterCompression -Confirm:$false -WhatIf:$false }
                        }
                        'Minute'
                        {
                            $removeScript   = [scriptblock]{ Remove-PreviousItem   -Path $_.FullName -Property $property -Minute $deletion -Confirm:$false -WhatIf:$false }
                            $compressScript = [scriptblock]{ Compress-PreviousItem -Path $_.FullName -Property $property -Minute $compression -CompressionLevel $compressionLevel -RemoveAfterCompression -Confirm:$false -WhatIf:$false }
                        }
                        'Second'
                        {
                            $removeScript   = [scriptblock]{ Remove-PreviousItem   -Path $_.FullName -Property $property -Second $deletion -Confirm:$false -WhatIf:$false }
                            $compressScript = [scriptblock]{ Compress-PreviousItem -Path $_.FullName -Property $property -Second $compression -CompressionLevel $compressionLevel -RemoveAfterCompression -Confirm:$false -WhatIf:$false }
                        }
                        'Millisecond'
                        {
                            $removeScript   = [scriptblock]{ Remove-PreviousItem   -Path $_.FullName -Property $property -Millisecond $deletion -Confirm:$false -WhatIf:$false }
                            $compressScript = [scriptblock]{ Compress-PreviousItem -Path $_.FullName -Property $property -Millisecond $compression -CompressionLevel $compressionLevel -RemoveAfterCompression -Confirm:$false -WhatIf:$false }
                        }
                    }

                    # Throw TerminatingError if both Parameter 'Deletion' and 'Compression' are null or empty.
                    if ((-not $PSBoundParameters['Deletion']) -and (-not $PSBoundParameters['Compression']))
                    {
                        $errorRecord = [System.Management.Automation.ErrorRecord]::new(
                            [System.Management.Automation.PSInvalidOperationException]::new(),
                            'InvalidOperation',
                            [System.Management.Automation.ErrorCategory]::InvalidOperation,
                            $null
                        )
                        $PSCmdlet.ThrowTerminatingError($errorRecord)
                    }

                    # Throw TerminatingError if either of the parameters 'Deletion' or 'Compression' is less than or equal to 0.
                    if (($deletion -le 0) -or ($compression -le 0))
                    {
                        $errorRecord = [System.Management.Automation.ErrorRecord]::new(
                            [System.Management.Automation.PSArgumentOutOfRangeException]::new(),
                            'InvalidRange',
                            [System.Management.Automation.ErrorCategory]::InvalidArgument,
                            $null
                        )
                        $PSCmdlet.ThrowTerminatingError($errorRecord)
                    }

                    # Throw TerminatingError if the parameter 'Deletion' is less than or equal to 'Compression'.
                    if ($PSBoundParameters['Deletion'] -and $PSBoundParameters['Compression'] -and ($deletion -le $compression))
                    {
                        $errorRecord = [System.Management.Automation.ErrorRecord]::new(
                            [System.Management.Automation.PSArgumentOutOfRangeException]::new(),
                            'InvalidRange',
                            [System.Management.Automation.ErrorCategory]::InvalidArgument,
                            $null
                        )
                        $PSCmdlet.ThrowTerminatingError($errorRecord)
                    }

                    if ($PSCmdlet.ShouldProcess($_.FullName))
                    {
                        # If the parameter 'Deletion' is present, perform the deletion.
                        if ($PSBoundParameters['Deletion'])
                        {
                            Invoke-Command -ScriptBlock $removeScript
                        }

                        # If the parameter 'Compression' is present, the target has not been deleted, and the extension is not .zip, perform the compression.
                        if ($PSBoundParameters['Compression'] -and (Test-Path -Path $_.FullName) -and ([System.IO.Path]::GetExtension($_.FullName) -ne '.zip'))
                        {
                            Invoke-Command -ScriptBlock $compressScript
                        }
                    }
                }
            }
        }
        catch
        {
            $PSCmdlet.ThrowTerminatingError($PSItem)
        }
    }
}
