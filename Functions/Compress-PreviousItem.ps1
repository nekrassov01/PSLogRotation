#Requires -Version 5.1

<#
.Synopsis
Compress old files and folders.

.DESCRIPTION
Compress old files and folders.
 - You can select properties from 'CreationTime', 'LastWriteTime' and 'LastAccessTime'. The default parameter is 'CreationTime'.
 - You can select the time unit from Year, Month, Day, Hour, Minute, Second, Millisecond or directly select DateTime.
 - You can select to delete or not the item after compression

.EXAMPLE
Compress-PreviousItem -Path 'C:\Work\test-1', 'C:\Work\test-2' -Day 90

.EXAMPLE
'C:\Work\test-1', 'C:\Work\test-2' | Compress-PreviousItem -Hour 12

.EXAMPLE
Compress-PreviousItem -Path 'C:\Work\test-1', 'C:\Work\test-2' -Day 90 -Property CreationTime

.EXAMPLE
'C:\Work\test-1', 'C:\Work\test-2' | Compress-PreviousItem -Minute 30 -Property CreationTime

.EXAMPLE
Compress-PreviousItem -Path 'C:\Work\test-1', 'C:\Work\test-2' -Day 90 -Property LastWriteTime

.EXAMPLE
'C:\Work\test-1', 'C:\Work\test-2' | Compress-PreviousItem -DateTime '2020/12/01 0:0:0' -Property LastWriteTime

.EXAMPLE
Compress-PreviousItem -Path 'C:\Work\test-1', 'C:\Work\test-2' -Day 90 -Property LastWriteTime -DestinationDirectory C:\archives

.EXAMPLE
'C:\Work\test-1', 'C:\Work\test-2' | Compress-PreviousItem -Second 60 -Property LastWriteTime -DestinationDirectory C:\archives

.EXAMPLE
Compress-PreviousItem -Path 'C:\Work\test-1', 'C:\Work\test-2' -Day 90 -Property LastWriteTime -CompressionLevel 'Fastest'

.EXAMPLE
'C:\Work\test-1', 'C:\Work\test-2' | Compress-PreviousItem -MilliSecound 1 -Property LastWriteTime -CompressionLevel 'NoCompression'

.EXAMPLE
Compress-PreviousItem -Path 'C:\Work\test-1', 'C:\Work\test-2' -Day 90 -Property LastWriteTime -RemoveAfterCompression

.EXAMPLE
'C:\Work\test-1', 'C:\Work\test-2' | Compress-PreviousItem -Year 1 -Property LastWriteTime -RemoveAfterCompression

.NOTES
Author: nekrassov01
#>

function Compress-PreviousItem
{
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High', DefaultParameterSetName = 'Day')]
    [OutputType([psobject[]])]
    param
    (
        [Parameter(Position = 0, Mandatory = $false, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [System.IO.FileInfo[]]$Path,

        [Parameter(Position = 1, Mandatory = $false)]
        [ValidateSet('CreationTime', 'LastWriteTime', 'LastAccessTime')]
        [string]$Property = 'CreationTime',

        [Parameter(Position = 2, Mandatory = $true, ParameterSetName = 'Year')]
        [int]$Year,

        [Parameter(Position = 2, Mandatory = $true, ParameterSetName = 'Month')]
        [int]$Month,

        [Parameter(Position = 2, Mandatory = $true, ParameterSetName = 'Day')]
        [int]$Day,

        [Parameter(Position = 2, Mandatory = $true, ParameterSetName = 'Hour')]
        [int]$Hour,

        [Parameter(Position = 2, Mandatory = $true, ParameterSetName = 'Minute')]
        [int]$Minute,

        [Parameter(Position = 2, Mandatory = $true, ParameterSetName = 'Second')]
        [int]$Second,

        [Parameter(Position = 2, Mandatory = $true, ParameterSetName = 'Millisecond')]
        [int]$Millisecond,

        [Parameter(Position = 2, Mandatory = $true, ParameterSetName = 'DateTime')]
        [string]$DateTime,

        [Parameter(Position = 3, Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [string]$DestinationDirectory,

        [Parameter(Position = 4, Mandatory = $false)]
        [switch]$RemoveAfterCompression,

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
            $Path | ForEach-Object -Process {

                $criteria = switch ($PSCmdlet.ParameterSetName)
                {
                    'Year'        { (Get-Date).AddYears(-$Year) }
                    'Month'       { (Get-Date).AddMonths(-$Month) }
                    'Day'         { (Get-Date).AddDays(-$Day) }
                    'Hour'        { (Get-Date).AddHours(-$Hour) }
                    'Minute'      { (Get-Date).AddMinutes(-$Minute) }
                    'Second'      { (Get-Date).AddSeconds(-$Second) }
                    'Millisecond' { (Get-Date).AddMilliseconds(-$Millisecond) }
                    'DateTime'    { [datetime]$DateTime }
                }

                Get-ChildItem -LiteralPath $_ | ForEach-Object -Process {

                    $itemType = switch ($_)
                    {
                        { $null -ne $_.LinkType } { $_.LinkType }
                        { $null -eq $_.LinkType } { if ($_.PSIsContainer) { 'Directory' } else { 'File' } }
                    }

                    $filteredItem = $_ | Where-Object -FilterScript { $_.$property -le $criteria }

                    if ($null -eq $filteredItem){ $null }

                    $filteredItem | ForEach-Object -Process {

                        $fileNameWithoutExtension = [System.IO.Path]::GetFileNameWithoutExtension($_.FullName)

                        if ($PSBoundParameters.ContainsKey('DestinationDirectory'))
                        {
                            $destinationPath = Join-Path -Path $destinationDirectory -ChildPath $fileNameWithoutExtension
                        }
                        else
                        {
                            $parentDir = Split-Path -Path $_.FullName -Parent
                            $destinationPath = Join-Path -Path $parentDir -ChildPath $fileNameWithoutExtension
                        }

                        if ($PSCmdlet.ShouldProcess($_.FullName))
                        {
                            if (($PSBoundParameters.ContainsKey('DestinationDirectory')) -and (-not (Test-Path -LiteralPath $destinationDirectory)))
                            {
                                New-Item -Path $destinationDirectory -ItemType Directory -Force | Out-Null
                            }

                            if ((Test-Path -LiteralPath $_.FullName) -and ([System.IO.Path]::GetExtension($_.FullName) -ne '.zip'))
                            {
                                Compress-Archive -LiteralPath $_.FullName -DestinationPath $destinationPath -CompressionLevel $compressionLevel -Force -Confirm:$false -WhatIf:$false

                                if ($PSBoundParameters.ContainsKey('RemoveAfterCompression'))
                                {
                                    try
                                    {
                                        Remove-Item -LiteralPath $_.FullName -Force -Recurse -Confirm:$false -WhatIf:$false
                                    }
                                    catch
                                    {
                                        $PSCmdlet.WriteError($PSItem)
                                    }
                                }

                                # Using '$?' to generate a return value
                                if ($?)
                                {
                                    $obj = [PSCustomObject]@{
                                        FileInfo  = $_
                                        ItemType  = $itemType
                                        Property  = $property
                                        ItemValue = $_.$property
                                        Criteria  = $criteria
                                        Action    = $PSCmdlet.MyInvocation.MyCommand.Verb
                                    }

                                    $PSCmdlet.WriteObject($obj)
                                }
                            }
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
