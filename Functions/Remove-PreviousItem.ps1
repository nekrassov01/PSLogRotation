#Requires -Version 5.1

<#
.SYNOPSIS
Deletes old files and folders.

.DESCRIPTION
Deletes old files and folders.
 - You can select properties from 'CreationTime', 'LastWriteTime' and 'LastAccessTime'. The default parameter is 'CreationTime'.
 - You can select the time unit from Year, Month, Day, Hour, Minute, Second, Millisecond or directly select DateTime.

.EXAMPLE
Remove-PreviousItem -Path 'C:\Work\test-1', 'C:\Work\test-2' -Day 90

.EXAMPLE
'C:\Work\test-1', 'C:\Work\test-2' | Remove-PreviousItem -Day 90

.EXAMPLE
Remove-PreviousItem -Path 'C:\Work\test-1', 'C:\Work\test-2' -Day 90 -Property CreationTime

.EXAMPLE
'C:\Work\test-1', 'C:\Work\test-2' | Remove-PreviousItem -Day 90 -Property CreationTime

.EXAMPLE
Remove-PreviousItem -Path 'C:\Work\test-1', 'C:\Work\test-2' -Day 90 -Property LastWriteTime

.EXAMPLE
'C:\Work\test-1', 'C:\Work\test-2' | Remove-PreviousItem -Day 90 -Property LastWriteTime

.NOTES
Author: nekrassov01
#>

function Remove-PreviousItem
{
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High', DefaultParameterSetName = 'Day')]
    [OutputType([psobject[]])]
    param
    (
        [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
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
        [string]$DateTime
    )

    begin
    {
        Set-StrictMode -Version Latest
    }

    process
    {
        try
        {
            $path | ForEach-Object -Process {

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

                Get-ChildItem -LiteralPath $_ -Force | ForEach-Object -Process {

                    $itemType = switch ($_)
                    {
                        { $null -ne $_.LinkType } { $_.LinkType }
                        { $null -eq $_.LinkType } { if ($_.PSIsContainer) { 'Directory' } else { 'File' } }
                    }

                    $filteredItem = $_ | Where-Object -FilterScript { $_.$property -le $criteria }

                    if ($null -eq $filteredItem){ $null }

                    $filteredItem | ForEach-Object -Process {

                        if ($PSCmdlet.ShouldProcess($_.FullName))
                        {
                            if (Test-Path -LiteralPath $_.FullName)
                            {
                                try
                                {
                                    Remove-Item -LiteralPath $_.FullName -Force -Recurse -Confirm:$false -WhatIf:$false

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
                                catch
                                {
                                    $PSCmdlet.WriteError($PSItem)
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
