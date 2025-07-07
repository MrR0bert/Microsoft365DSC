function Get-TargetResource
{
    [CmdletBinding()]
    [OutputType([System.Collections.Hashtable])]
    param (

        #region resource params

        [Parameter()]
        [System.String]
        $Id,

        [Parameter(Mandatory = $true)]
        [System.String]
        $DisplayName,

        [Parameter()]
        [Microsoft.Management.Infrastructure.CimInstance[]]
        $Assignments,

        [Parameter()]
        [System.String]
        $ContactItEmailAddress,

        [Parameter()]
        [System.String]
        $ContactItName,

        [Parameter()]
        [System.String]
        $ContactItNotes,

        [Parameter()]
        [System.String]
        $ContactItPhoneNumber,

        [Parameter()]
        [System.String]
        $CustomCanSeePrivacyMessage,

        [Parameter()]
        [System.String]
        $CustomCantSeePrivacyMessage,

        [Parameter()]
        [System.String]
        $CustomPrivacyMessage,

        [Parameter()]
        [System.Boolean]
        $DisableDeviceCategorySelection,

        [Parameter()]
        [System.String]
        [ValidateSet('availableWithPrompts', 'availableWithoutPrompts', 'unavailable')]
        $EnrollmentAvailability,

        [Parameter()]
        [System.Boolean]
        $IsFactoryResetDisabled,

        [Parameter()]
        [System.Boolean]
        $IsRemoveDeviceDisabled,

        [Parameter()]
        [System.String]
        $LandingPageCustomizedImage,

        [Parameter()]
        [System.String]
        $LightBackgroundLogo,

        [Parameter()]
        [System.String]
        $OnlineSupportSiteName,

        [Parameter()]
        [System.String]
        $OnlineSupportSiteUrl,

        [Parameter()]
        [System.String]
        $PrivacyUrl,

        [Parameter()]
        [System.String]
        $ProfileDescription,

        [Parameter()]
        [System.String]
        $ProfileName,

        [Parameter()]
        [System.String[]]
        $RoleScopeTagIds,

        [Parameter()]
        [System.Boolean]
        $SendDeviceOwnershipChangePushNotification,

        [Parameter()]
        [System.Boolean]
        $ShowAzureAdEnterpriseApps,

        [Parameter()]
        [System.Boolean]
        $ShowConfigurationManagerApps,

        [Parameter()]
        [System.Boolean]
        $ShowDisplayNameNextToLogo,

        [Parameter()]
        [System.Boolean]
        $ShowLogo,

        [Parameter()]
        [System.Boolean]
        $ShowOfficeWebApps,

        [Parameter()]
        [Microsoft.Management.Infrastructure.CimInstance]
        $ThemeColor,

        [Parameter()]
        [System.String]
        $ThemeColorLogo,

        #endregion resource params

        [Parameter()]
        [ValidateSet('Present', 'Absent')]
        [System.String]
        $Ensure = 'Present',

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $Credential,

        [Parameter()]
        [System.String]
        $ApplicationId,

        [Parameter()]
        [System.String]
        $TenantId,

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $ApplicationSecret,

        [Parameter()]
        [System.String]
        $CertificateThumbprint,

        [Parameter()]
        [Switch]
        $ManagedIdentity,

        [Parameter()]
        [System.String[]]
        $AccessTokens

    )

    Write-Verbose -Message "Getting Intune Branding Profile with ID {$Id} and DisplayName {$DisplayName}."

    try
    {
        if (-not $Script:exportedInstance -or $Script:exportedInstance.DisplayName -ne $DisplayName)
        {
            Write-Verbose "Instance is not exported or the exported displayname does not match the provided displayname"
            New-M365DSCConnection -Workload 'MicrosoftGraph' `
                -InboundParameters $PSBoundParameters | Out-Null

            #Ensure the proper dependencies are installed in the current environment.
            Confirm-M365DSCDependencies

            #region Telemetry
            $ResourceName = $MyInvocation.MyCommand.ModuleName.Replace('MSFT_', '')
            $CommandName = $MyInvocation.MyCommand
            $data = Format-M365DSCTelemetryParameters -ResourceName $ResourceName `
                -CommandName $CommandName `
                -Parameters $PSBoundParameters
            Add-M365DSCTelemetryEvent -Data $data
            #endregion

            $nullResult = $PSBoundParameters
            $nullResult.Ensure = 'Absent'

            $instance = $null
            if (-not [System.String]::IsNullOrEmpty($Id))
            {
                Write-Verbose "Getting branding profile with ID {$($Id)}"
                $instance = Get-MgBetaDeviceManagementIntuneBrandingProfile -IntuneBrandingProfileId $Id -ErrorAction SilentlyContinue
            }

            if ($null -eq $instance)
            {
                Write-Verbose -Message "Could not find Intune Branding Profile by Id {$Id}."

                if (-not [string]::IsNullOrEmpty($DisplayName))
                {
                    $instance = Get-MgBetaDeviceManagementIntuneBrandingProfile `
                        -All `
                        -Filter "DisplayName eq '$($DisplayName -replace "'", "''")'" `
                        -ErrorAction SilentlyContinue

                    if ($null -eq $instance)
                    {
                        Write-Verbose -Message "Could not find Intune Branding Profile by DisplayName {$DisplayName}."
                        return $nullResult
                    }
                }
            }
        }
        else
        {
            write-verbose "Using exported instance"
            $instance = $Script:exportedInstance
        }

        Write-Verbose "The results are in for {$($instance.Id)} / {$($instance.DisplayName)}"

        $themeColor = New-CimInstance -ClassName MSFT_MicrosoftGraphRgbColor -Namespace "root/microsoft/Windows/DesiredStateConfiguration" -ClientOnly -Property @{
            R = [uint32]($instance.ThemeColor.R -as [int] -as [uint32])
            G = [uint32]($instance.ThemeColor.G -as [int] -as [uint32])
            B = [uint32]($instance.ThemeColor.B -as [int] -as [uint32])
        }


        $results = @{
            Ensure                                      = 'Present'
            Id                                          = $instance.Id
            DisplayName                                 = $instance.DisplayName
            #Assignments                                 = $instance.Assignments
            ContactItEmailAddress                       = $instance.ContactItEmailAddress
            ContactItName                               = $instance.ContactItName
            ContactItNotes                              = $instance.ContactItNotes
            ContactItPhoneNumber                        = $instance.ContactItPhoneNumber
            CustomCanSeePrivacyMessage                  = $instance.CustomCanSeePrivacyMessage
            CustomCantSeePrivacyMessage                 = $instance.CustomCantSeePrivacyMessage
            #CustomPrivacyMessage                        = $instance.CustomPrivacyMessage
            #DisableClientTelemetry                      = $instance.DisableClientTelemetry > NOT YET SUPPORTED BY GRAPH
            DisableDeviceCategorySelection              = $instance.DisableDeviceCategorySelection
            EnrollmentAvailability                      = $instance.EnrollmentAvailability.toString()
            #IsFactoryResetDisabled                      = $instance.IsFactoryResetDisabled
            #IsRemoveDeviceDisabled                      = $instance.IsRemoveDeviceDisabled
            #LandingPageCustomizedImage                  = $instance.LandingPageCustomizedImage
            #LightBackgroundLogo                         = $instance.LightBackgroundLogo
            OnlineSupportSiteName                       = $instance.OnlineSupportSiteName
            OnlineSupportSiteUrl                        = $instance.OnlineSupportSiteUrl
            PrivacyUrl                                  = $instance.PrivacyUrl
            ProfileDescription                          = $instance.ProfileDescription
            ProfileName                                 = $instance.ProfileName
            RoleScopeTagIds                             = $instance.RoleScopeTagIds
            #SendDeviceOwnershipChangePushNotification   = $instance.SendDeviceOwnershipChangePushNotification
            ShowAzureAdEnterpriseApps                   = $instance.ShowAzureAdEnterpriseApps
            ShowConfigurationManagerApps                = $instance.ShowConfigurationManagerApps
            ShowDisplayNameNextToLogo                   = $instance.ShowDisplayNameNextToLogo
            ShowLogo                                    = $instance.ShowLogo
            ShowOfficeWebApps                           = $instance.ShowOfficeWebApps
            ThemeColor                                  = $themeColor
            #ThemeColorLogo                              = $instance.ThemeColorLogo
            Credential                                  = $Credential
            ApplicationId                               = $ApplicationId
            TenantId                                    = $TenantId
            CertificateThumbprint                       = $CertificateThumbprint
            ApplicationSecret                           = $ApplicationSecret
            ManagedIdentity                             = $ManagedIdentity.IsPresent
            AccessTokens                                = $AccessTokens
        }

        return [System.Collections.Hashtable] $results
    }
    catch
    {
        Write-Verbose -Message $_
        New-M365DSCLogEntry -Message 'Error retrieving data:' `
            -Exception $_ `
            -Source $($MyInvocation.MyCommand.Source) `
            -TenantId $TenantId `
            -Credential $Credential

        return $nullResult
    }
}

function Set-TargetResource
{
    [CmdletBinding()]
    param (

        #region resource params

        [Parameter()]
        [System.String]
        $Id,

        [Parameter(Mandatory = $true)]
        [System.String]
        $DisplayName,

        [Parameter()]
        [Microsoft.Management.Infrastructure.CimInstance[]]
        $Assignments,

        [Parameter()]
        [System.String]
        $ContactItEmailAddress,

        [Parameter()]
        [System.String]
        $ContactItName,

        [Parameter()]
        [System.String]
        $ContactItNotes,

        [Parameter()]
        [System.String]
        $ContactItPhoneNumber,

        [Parameter()]
        [System.String]
        $CustomCanSeePrivacyMessage,

        [Parameter()]
        [System.String]
        $CustomCantSeePrivacyMessage,

        [Parameter()]
        [System.String]
        $CustomPrivacyMessage,

        [Parameter()]
        [System.Boolean]
        $DisableDeviceCategorySelection,

        [Parameter()]
        [System.String]
        [ValidateSet('availableWithPrompts', 'availableWithoutPrompts', 'unavailable')]
        $EnrollmentAvailability,

        [Parameter()]
        [System.Boolean]
        $IsFactoryResetDisabled,

        [Parameter()]
        [System.Boolean]
        $IsRemoveDeviceDisabled,

        [Parameter()]
        [System.String]
        $LandingPageCustomizedImage,

        [Parameter()]
        [System.String]
        $LightBackgroundLogo,

        [Parameter()]
        [System.String]
        $OnlineSupportSiteName,

        [Parameter()]
        [System.String]
        $OnlineSupportSiteUrl,

        [Parameter()]
        [System.String]
        $PrivacyUrl,

        [Parameter()]
        [System.String]
        $ProfileDescription,

        [Parameter()]
        [System.String]
        $ProfileName,

        [Parameter()]
        [System.String[]]
        $RoleScopeTagIds,

        [Parameter()]
        [System.Boolean]
        $SendDeviceOwnershipChangePushNotification,

        [Parameter()]
        [System.Boolean]
        $ShowAzureAdEnterpriseApps,

        [Parameter()]
        [System.Boolean]
        $ShowConfigurationManagerApps,

        [Parameter()]
        [System.Boolean]
        $ShowDisplayNameNextToLogo,

        [Parameter()]
        [System.Boolean]
        $ShowLogo,

        [Parameter()]
        [System.Boolean]
        $ShowOfficeWebApps,

        [Parameter()]
        [Microsoft.Management.Infrastructure.CimInstance]
        $ThemeColor,

        [Parameter()]
        [System.String]
        $ThemeColorLogo,

        #endregion resource params

        [Parameter()]
        [ValidateSet('Present', 'Absent')]
        [System.String]
        $Ensure = 'Present',

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $Credential,

        [Parameter()]
        [System.String]
        $ApplicationId,

        [Parameter()]
        [System.String]
        $TenantId,

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $ApplicationSecret,

        [Parameter()]
        [System.String]
        $CertificateThumbprint,

        [Parameter()]
        [Switch]
        $ManagedIdentity,

        [Parameter()]
        [System.String[]]
        $AccessTokens
    )

    #Ensure the proper dependencies are installed in the current environment.
    Confirm-M365DSCDependencies

    #region Telemetry
    $ResourceName = $MyInvocation.MyCommand.ModuleName.Replace('MSFT_', '')
    $CommandName = $MyInvocation.MyCommand
    $data = Format-M365DSCTelemetryParameters -ResourceName $ResourceName `
        -CommandName $CommandName `
        -Parameters $PSBoundParameters
    Add-M365DSCTelemetryEvent -Data $data
    #endregion

    $currentInstance = Get-TargetResource @PSBoundParameters

    $setParameters = Remove-M365DSCAuthenticationParameter -BoundParameters $PSBoundParameters
    $setParameters.Remove('Id') | Out-Null

    $keys = (([Hashtable]$PSBoundParameters).Clone()).Keys
    foreach ($key in $keys)
    {
        $keyName = $key.Substring(0, 1).ToLower() + $key.Substring(1, $key.Length - 1)
        $keyValue = $PSBoundParameters.$key
        if ($null -ne $PSBoundParameters.$key -and $PSBoundParameters.$key.GetType().Name -like '*cimInstance*')
        {
            $keyValue = Convert-M365DSCDRGComplexTypeToHashtable -ComplexObject $PSBoundParameters.$key
        }
        $PSBoundParameters.Remove($key)
        $PSBoundParameters.Add($keyName, $keyValue)
    }

    $graphColor = [Microsoft.Graph.Beta.PowerShell.Models.MicrosoftGraphRgbColor]::new()
    $graphColor.R = [int]$ThemeColor.R
    $graphColor.G = [int]$ThemeColor.G
    $graphColor.B = [int]$ThemeColor.B

    $setParameters.ThemeColor = $graphColor

    Write-Verbose -Message "ThemeColor values are of type $($graphColor.GetType())"

    # CREATE
    if ($Ensure -eq 'Present' -and $currentInstance.Ensure -eq 'Absent')
    {
        Write-Verbose -Message "Creating an Intune Branding Profile with DisplayName {$DisplayName}"
        New-MgBetaDeviceManagementIntuneBrandingProfile @SetParameters
    }
    # UPDATE
    elseif ($Ensure -eq 'Present' -and $currentInstance.Ensure -eq 'Present')
    {
        Write-Verbose -Message "Updating Intune Branding Profile with DisplayName {$DisplayName} / {$($currentInstance.Id)}"

        # For unkown reasons Graph does NOT allow for PATCH requests to update an IntuneBrandingProfile resource, so we need to delete and re-create it
        #Update-MgBetaDeviceManagementIntuneBrandingProfile @SetParameters -ThemeColor $graphColor -IntuneBrandingProfileId $currentInstance.Id

        Write-Verbose -Message "Graph does NOT allow for PATCH requests to update an IntuneBrandingProfile resource, so we need to delete and re-create it"
        try {
            Remove-MgBetaDeviceManagementIntuneBrandingProfile -IntuneBrandingProfileId $currentInstance.Id -Confirm:$false -Verbose -ErrorAction Stop
            Write-Verbose -Message "Removal succes, creating new profile"
            New-MgBetaDeviceManagementIntuneBrandingProfile @SetParameters -Verbose
            Write-Verbose -Message "Done updating profile {$($DisplayName)}"
        }
        catch {
            Write-Error $_.Exception.Message
        }

    }
    # REMOVE
    elseif ($Ensure -eq 'Absent' -and $currentInstance.Ensure -eq 'Present')
    {
        Write-Verbose -Message "Removing the Intune Branding Profile with DisplayName {$DisplayName}"

        Remove-MgBetaDeviceManagementIntuneBrandingProfile -IntuneBrandingProfileId $currentInstance.Id -Confirm:$false
    }
}

function Test-TargetResource
{
    [CmdletBinding()]
    [OutputType([System.Boolean])]
    param (

        #region resource params

        [Parameter()]
        [System.String]
        $Id,

        [Parameter(Mandatory = $true)]
        [System.String]
        $DisplayName,

        [Parameter()]
        [Microsoft.Management.Infrastructure.CimInstance[]]
        $Assignments,

        [Parameter()]
        [System.String]
        $ContactItEmailAddress,

        [Parameter()]
        [System.String]
        $ContactItName,

        [Parameter()]
        [System.String]
        $ContactItNotes,

        [Parameter()]
        [System.String]
        $ContactItPhoneNumber,

        [Parameter()]
        [System.String]
        $CustomCanSeePrivacyMessage,

        [Parameter()]
        [System.String]
        $CustomCantSeePrivacyMessage,

        [Parameter()]
        [System.String]
        $CustomPrivacyMessage,

        [Parameter()]
        [System.Boolean]
        $DisableDeviceCategorySelection,

        [Parameter()]
        [System.String]
        [ValidateSet('availableWithPrompts', 'availableWithoutPrompts', 'unavailable')]
        $EnrollmentAvailability,

        [Parameter()]
        [System.Boolean]
        $IsFactoryResetDisabled,

        [Parameter()]
        [System.Boolean]
        $IsRemoveDeviceDisabled,

        [Parameter()]
        [System.String]
        $LandingPageCustomizedImage,

        [Parameter()]
        [System.String]
        $LightBackgroundLogo,

        [Parameter()]
        [System.String]
        $OnlineSupportSiteName,

        [Parameter()]
        [System.String]
        $OnlineSupportSiteUrl,

        [Parameter()]
        [System.String]
        $PrivacyUrl,

        [Parameter()]
        [System.String]
        $ProfileDescription,

        [Parameter()]
        [System.String]
        $ProfileName,

        [Parameter()]
        [System.String[]]
        $RoleScopeTagIds,

        [Parameter()]
        [System.Boolean]
        $SendDeviceOwnershipChangePushNotification,

        [Parameter()]
        [System.Boolean]
        $ShowAzureAdEnterpriseApps,

        [Parameter()]
        [System.Boolean]
        $ShowConfigurationManagerApps,

        [Parameter()]
        [System.Boolean]
        $ShowDisplayNameNextToLogo,

        [Parameter()]
        [System.Boolean]
        $ShowLogo,

        [Parameter()]
        [System.Boolean]
        $ShowOfficeWebApps,

        [Parameter()]
        [Microsoft.Management.Infrastructure.CimInstance]
        $ThemeColor,

        [Parameter()]
        [System.String]
        $ThemeColorLogo,

        #endregion resource params

        [Parameter()]
        [ValidateSet('Present', 'Absent')]
        [System.String]
        $Ensure = 'Present',

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $Credential,

        [Parameter()]
        [System.String]
        $ApplicationId,

        [Parameter()]
        [System.String]
        $TenantId,

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $ApplicationSecret,

        [Parameter()]
        [System.String]
        $CertificateThumbprint,

        [Parameter()]
        [Switch]
        $ManagedIdentity,

        [Parameter()]
        [System.String[]]
        $AccessTokens
    )

    Confirm-M365DSCDependencies

    #region Telemetry
    $ResourceName = $MyInvocation.MyCommand.ModuleName.Replace('MSFT_', '')
    $CommandName = $MyInvocation.MyCommand
    $data = Format-M365DSCTelemetryParameters -ResourceName $ResourceName `
        -CommandName $CommandName `
        -Parameters $PSBoundParameters
    Add-M365DSCTelemetryEvent -Data $data
    #endregion

    $CurrentValues = Get-TargetResource @PSBoundParameters
    $ValuesToCheck = ([Hashtable]$PSBoundParameters).Clone()
    $testResult = $true

    $ValuesToCheck = Remove-M365DSCAuthenticationParameter -BoundParameters $ValuesToCheck
    $ValuesToCheck.Remove('Id') | Out-Null

    Write-Verbose -Message "Current Values: $(Convert-M365DscHashtableToString -Hashtable $CurrentValues)"
    Write-Verbose -Message "Target Values: $(Convert-M365DscHashtableToString -Hashtable $ValuesToCheck)"

    if ($testResult)
    {
        $testResult = Test-M365DSCParameterState -CurrentValues $CurrentValues `
            -Source $($MyInvocation.MyCommand.Source) `
            -DesiredValues $PSBoundParameters `
            -ValuesToCheck $ValuesToCheck.Keys
    }

    Write-Verbose -Message "Test-TargetResource returned $testResult"

    return $testResult
}

function Export-TargetResource
{
    [CmdletBinding()]
    [OutputType([System.String])]
    param (

        [Parameter()]
        [System.String]
        $Filter,

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $Credential,

        [Parameter()]
        [System.String]
        $ApplicationId,

        [Parameter()]
        [System.String]
        $TenantId,

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $ApplicationSecret,

        [Parameter()]
        [System.String]
        $CertificateThumbprint,

        [Parameter()]
        [Switch]
        $ManagedIdentity,

        [Parameter()]
        [System.String[]]
        $AccessTokens
    )

    $ConnectionMode = New-M365DSCConnection -Workload 'MicrosoftGraph' `
        -InboundParameters $PSBoundParameters

    #Ensure the proper dependencies are installed in the current environment.
    Confirm-M365DSCDependencies

    #region Telemetry
    $ResourceName = $MyInvocation.MyCommand.ModuleName.Replace('MSFT_', '')
    $CommandName = $MyInvocation.MyCommand
    $data = Format-M365DSCTelemetryParameters -ResourceName $ResourceName `
        -CommandName $CommandName `
        -Parameters $PSBoundParameters
    Add-M365DSCTelemetryEvent -Data $data
    #endregion

    try
    {
        $Script:ExportMode = $true
        [array] $getValue = Get-MgBetaDeviceManagementIntuneBrandingProfile -All -ExpandProperty Assignments -ErrorAction Stop

        $i = 1
        $dscContent = ''
        if ($getValue.Length -eq 0)
        {
            Write-M365DSCHost -Message $Global:M365DSCEmojiGreenCheckMark -CommitWrite
        }
        else
        {
            Write-M365DSCHost -Message "`r`n" -DeferWrite
        }
        foreach ($config in $getValue)
        {
            $displayedKey = $config.Id
            Write-M365DSCHost -Message "    |---[$i/$($getValue.Count)] $displayedKey" -DeferWrite

            $params = @{
                Ensure                                      = 'Present'
                DisplayName                                 = $config.displayName
                Credential                                  = $Credential
                AccessTokens                                = $AccessTokens
                ApplicationId                               = $ApplicationId
                TenantId                                    = $TenantId
                ApplicationSecret                           = $ApplicationSecret
                CertificateThumbprint                       = $CertificateThumbprint
                ManagedIdentity                             = $ManagedIdentity.IsPresent
            }

            $Script:exportedInstance = $config
            $Results = Get-TargetResource @Params

            if ($null -ne $Results.Assignments)
            {
                $complexMapping = @(
                    @{
                        Name            = 'Assignments'
                        CimInstanceName = 'MSFT_IntuneBrandingProfileAssignments'
                            sRequired      = $False
                    }
                )
                $complexTypeStringResult = Get-M365DSCDRGComplexTypeToString `
                    -ComplexObject $Results.Assignments `
                    -CIMInstanceName 'MSFT_IntuneBrandingProfileAssignments' `
                    -ComplexTypeMapping $complexMapping

                if (-not [String]::IsNullOrWhiteSpace($complexTypeStringResult))
                {
                    $Results.Assignments = $complexTypeStringResult
                }
                else
                {
                    $Results.Remove('Assignments') | Out-Null
                }
            }

            if ($null -ne $Results.ThemeColor)
            {
                $complexTypeStringResult = Get-M365DSCDRGComplexTypeToString `
                    -ComplexObject ($Results.ThemeColor) `
                    -CIMInstanceName MSFT_MicrosoftGraphRgbColor

                if (-not [String]::IsNullOrWhiteSpace($complexTypeStringResult))
                {
                    $Results.ThemeColor = $complexTypeStringResult
                }
                else
                {
                    $Results.Remove('ThemeColor') | Out-Null
                }
            }

            $currentDSCBlock = Get-M365DSCExportContentForResource -ResourceName $ResourceName `
                -ConnectionMode $ConnectionMode `
                -ModulePath $PSScriptRoot `
                -Results $Results `
                -Credential $Credential `
                -NoEscape @('ThemeColor')

            $dscContent += $currentDSCBlock
            Save-M365DSCPartialExport -Content $currentDSCBlock `
                -FileName $Global:PartialExportFileName
            $i++
            Write-M365DSCHost -Message $Global:M365DSCEmojiGreenCheckMark -CommitWrite
        }

        return $dscContent
    }
    catch
    {
        Write-M365DSCHost -Message $Global:M365DSCEmojiRedX -CommitWrite

        New-M365DSCLogEntry -Message 'Error during Export:' `
            -Exception $_ `
            -Source $($MyInvocation.MyCommand.Source) `
            -TenantId $TenantId `
            -Credential $Credential

        return ''
    }
}

Export-ModuleMember -Function *-TargetResource
