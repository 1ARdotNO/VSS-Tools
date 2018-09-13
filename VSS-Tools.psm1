#-----Script Info---------------------------------------------------------------------------------------------
# Name:VSS-Tools.psm1
# Author: Einar Stenberg
# Date: 24.11.14
# Version: 1
# Job/Tasks:
#--------------------------------------------------------------------------------------------------------------


#-----Changelog------------------------------------------------------------------------------------------------
#v1.  Script created ES
#
#
#
#--------------------------------------------------------------------------------------------------------------






#-----Functions---------------------------------------------------------------------------------------------


Function New-VSS {
<#
.SYNOPSIS
Creates a VSS snapshot of the selected drive
.DESCRIPTION
Can be used as pipelineinput for Remove-VSS and Add-VSSmount
Part of WHBACKUP by ES
.EXAMPLE
New-VSS
#>
Param(
$Driveletter
)

BEGIN{}

PROCESS{
    (Get-WmiObject -list win32_shadowcopy).create("$($driveletter):\","ClientAccessible") | Write-Output
}


}

Function Get-VSS {
<#
.SYNOPSIS
Gets all VSS snapshots
.DESCRIPTION
Gets all VSS snapshots, can be used as pipelineinput for Remove-VSS and Add-VSSmount
Part of WHBACKUP by ES
.EXAMPLE
Get-VSS
#>
Param(

)

Get-CimInstance -ClassName Win32_ShadowCopy

}

Function Remove-VSS {

<#
.SYNOPSIS
Deletes VSS copy with the specified ID
.DESCRIPTION
Deletes VSS copy with the specified ID
Accepts pipelineinput for the ID
Part of WHBACKUP by ES
.EXAMPLE
Remove-vss -id D6A25A6C-7DD6-47EA-B3E4-3A2BF5C8F202
Removes the VSS snapshot with ID D6A25A6C-7DD6-47EA-B3E4-3A2BF5C8F202
#>


Param(
[Parameter(Mandatory=$true,Position=0,ValueFromPipeline='True',valuefrompipelinebypropertyname='True')]
[Alias('ShadowID')]
[string]$ID
)

BEGIN{}

PROCESS{
    $shadowcopies=Get-WMIObject -Class Win32_ShadowCopy | where {$_.ID -like "*$ID*"}
    Write-Verbose "Removing VSS with ID $ID"
    $shadowcopies.Delete()
}

}

Function Add-VSSMount{

<#
.SYNOPSIS
Mounts the selected VSS to the specified folderpath
.DESCRIPTION
Mounts the selected VSS to the specified folderpath, accepts pipelineinput from new-vss
Part of WHBACKUP by ES
.EXAMPLE
Add-VSSMount -path c:\vssfolder -id {123456-123125-653534}
#>


Param(
[Parameter(Mandatory=$true,Position=0,ValueFromPipeline='True',valuefrompipelinebypropertyname='True' )]
[Alias('ShadowID')]
[string]$ID,
[string]$path
)

BEGIN{}

PROCESS{
    $shadowcopy=Get-WMIObject -Class Win32_ShadowCopy | where {$_.ID -like "*$ID*"}

    $cmdmklink="cmd /c mklink /d"

    Invoke-Expression "$cmdmklink $path $($shadowcopy.DeviceObject)\"

}

}

Function Remove-VSSMount{

<#
.SYNOPSIS
Removes the selected VSS mountpoint
.DESCRIPTION
Removes the selected VSS mountpoint
Part of WHBACKUP by ES
.EXAMPLE
Remove-VSSMount -path c:\vssfolder
#>

Param(
[string]$path
)

BEGIN{}

PROCESS{
    Remove-Item -Path $path
}


}

