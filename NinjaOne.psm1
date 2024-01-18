#requires -version 7
using namespace System.Net.Http
using namespace System.Management.Automation
using namespace NinjaOne.V2
$ErrorActionPreference = 'Stop'
Add-Type -Path $PSScriptRoot/out/NinjaOne.dll

#Load the default properties for each type
function Import-DefaultProperties ($defaultPropertiesPath, $namespace) {
  $defaultProperties = Import-PowerShellDataFile $defaultPropertiesPath
  foreach ($key in $defaultProperties.Keys) {
    $value = $defaultProperties[$key]
    if ($value -is [string]) {
      $value = $value.Split(',')
    }
    $value = $value.Trim()
    if ($prefix) { $key = "${namespace}.$key" }
    Update-TypeData -TypeName ([type]$key) -DefaultDisplayPropertySet $value -Force
  }
  Remove-Variable 'defaultProperties'
}
Import-DefaultProperties $PSScriptRoot/NinjaOne.defaultProperties.psd1 'NinjaOne.V2'

#Region Connect
class NinjaOne2AuthResult {
  [string]$access_token
  [string]$token_type
  [int]$expires_in
  [string]$scope
}

enum NinjaOne2AuthScopes {
  Monitoring
  Management
  Control
}

function Connect-N1 {
  <#
  .SYNOPSIS
  Connect to the NinjaOne API
  .NOTES
  Reference: https://eu.ninjarmm.com/apidocs-beta/authorization/flows/client-credentials-flow
  #>
  [CmdletBinding()]
  [OutputType([NinjaOne.V2.Client])]
  param(
    #Provide your Client ID as the username and Client Secret as the password
    [Parameter(Mandatory)][PSCredential]$Credential,
    #Scope to access. Multiple scopes can be specified. Default is Monitoring, Management, and Control
    [NinjaOne2AuthScopes[]]$Scope = @('Monitoring', 'Management', 'Control'),
    #The root of the site you are connecting to with the trailing slash included. Default is the EU ninjarmm site.
    [Parameter(ValueFromPipeline)][string]$BaseUri = 'https://eu.ninjarmm.com/',
    #Return the generated client object. If you specify this parameter, the client will not be set as the default client for the session.
    [switch]$PassThru
  )
  $ErrorActionPreference = 'Stop'

  $uri = "${BaseUri}ws/oauth/token"

  [NinjaOne2AuthResult]$authResult = Invoke-RestMethod -Uri $Uri -Method 'POST' -Body @{
    grant_type    = 'client_credentials'
    scope         = ($Scope -join ' ').ToLower()
    client_id     = $Credential.UserName
    client_secret = $Credential.GetNetworkCredential().Password
  }

  #Generate a new N1 client
  $httpClient = [System.Net.Http.HttpClient]::new()
  $httpClient.DefaultRequestHeaders.Authorization = [Headers.AuthenticationHeaderValue]::new('Bearer', $authResult.access_token)

  $client = [Client]::new($BaseUri, $httpClient)
  $client | Add-Member -NotePropertyName 'Expires' -NotePropertyValue (Get-Date).AddSeconds($authResult.expires_in)
  $client | Add-Member -NotePropertyName 'ClientId' -NotePropertyValue $Credential.UserName
  $client | Add-Member -NotePropertyName 'Scope' -NotePropertyValue $Scope

  if ($PassThru) {
    return $client
  } else {
    Set-N1DefaultClient $client
  }
}

function Set-N1DefaultClient {
  <#
  .SYNOPSIS
  Set the default N1 client for the session
  #>
  param(
    [Parameter(ValueFromPipeline)][Client]$Client
  )
  if ($client.Expires -lt (Get-Date)) {
    Write-Error "Your N1 client expired at $($SCRIPT:CurrentClient.Expires). Please reconnect using Connect-N1"
  }
  [NinjaOne.V2.Client]$SCRIPT:CurrentClient = $Client
}


function Get-N1Client {
  <#
  .SYNOPSIS
  Get the current N1 client
  .NOTES
  #>
  [CmdletBinding()]
  [OutputType([NinjaOne.V2.Client])]
  param()
  $ErrorActionPreference = 'Stop'

  if ($null -eq $SCRIPT:CurrentClient) {
    Write-Error 'You must connect to the N1 API first using Connect-N1'
  }

  if ($SCRIPT:CurrentClient.Expires -lt (Get-Date)) {
    Write-Error "Your N1 client expired at $($SCRIPT:CurrentClient.Expires). Please reconnect using Connect-N1"
  }
  return $SCRIPT:CurrentClient
}
#EndRegion Connect

#region Organization
function Get-N1Organization {
  <#
  .SYNOPSIS
  Get the current organization
  .NOTES
  #>
  [CmdletBinding(DefaultParameterSetName = 'List')]
  [OutputType([NinjaOne.V2.Organization])]
  param(
    [Parameter(ParameterSetName = 'List')]
    [ValidateRange(1, [int]::MaxValue)]
    $AfterId,
    [Parameter(ParameterSetName = 'List')]
    [ValidateRange(1, [int]::MaxValue)]
    $Count,
    [Parameter(Mandatory, ParameterSetName = 'GetObject', ValueFromPipeline)][NinjaOne.V2.Organization]$Organization,
    [Parameter(Mandatory, ParameterSetName = 'Get', ValueFromPipelineByPropertyName)][int]$Id,
    [ValidateNotNullOrEmpty()][Parameter(ValueFromPipelineByPropertyName)]
    [NinjaOne.V2.Client]$Client = (Get-N1Client)
  )
  process {
    if ($PSCmdlet.ParameterSetName -in 'Get', 'GetObject') {
      if ($PSCmdlet.ParameterSetName -eq 'GetObject') {
        $Id = $Organization.Id
      }
      $client.GetOrganization($Id)
      | Resolve-N1Error $Id
      | Add-Member -NotePropertyName 'Client' -NotePropertyValue $client -PassThru
    }
  }
  end {
    if ($PSCmdlet.ParameterSetName -eq 'List') {
      return $client.GetOrganizations($Count, $AfterId) | Add-Member -NotePropertyName 'Client' -NotePropertyValue $client -PassThru
    }
  }
}

function Get-N1OrganizationDetailed {
  <#
  .SYNOPSIS
  Get the current organization
  .NOTES
  #>
  [CmdletBinding(DefaultParameterSetName = 'List')]
  [OutputType([NinjaOne.V2.OrganizationDetailed])]
  param(
    [Parameter(ParameterSetName = 'List')]
    [ValidateRange(1, [int]::MaxValue)]
    $AfterId,
    [Parameter(ParameterSetName = 'List')]
    [ValidateRange(1, [int]::MaxValue)]
    $Count,
    [Parameter(Mandatory, ParameterSetName = 'GetObject', ValueFromPipeline)]
    [Organization]$Organization,
    [Parameter(Mandatory, ParameterSetName = 'Get', ValueFromPipelineByPropertyName)][int]$Id,
    [ValidateNotNullOrEmpty()][Parameter(ValueFromPipelineByPropertyName)]
    [NinjaOne.V2.Client]$Client = (Get-N1Client)
  )
  process {
    if ($PSCmdlet.ParameterSetName -in 'Get', 'GetObject') {
      if ($PSCmdlet.ParameterSetName -eq 'GetObject') {
        $Id = $Organization.Id
      }
      if ($PSCmdlet.ParameterSetName -eq 'ByOrganization') {
        $Id = $Organization.Id
      }
      $client.GetOrganizationsDetailed($Id)
      | Resolve-N1Error $Id
      | Add-Member -NotePropertyName 'Client' -NotePropertyValue $client -PassThru
    }
  }
  end {
    if ($PSCmdlet.ParameterSetName -eq 'List') {
      return $client.GetOrganizations($Count, $AfterId) | Add-Member -NotePropertyName 'Client' -NotePropertyValue $client -PassThru
    }
  }
}

#Help purposefully removed for dynamic param help generation, where the parameters are dynamically generated from the model.
function New-N1Organization {
  [CmdletBinding(DefaultParameterSetName = 'ByParams', SupportsShouldProcess, ConfirmImpact = 'High')]
  [OutputType([NinjaOne.V2.OrganizationDetailed])]
  param(
    [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'ByModel')]
    [NinjaOne.V2.OrganizationWithLocationsAndPolicyAssignmentsModel]$Model,

    [Parameter(ValueFromPipelineByPropertyName)]
    [ValidateRange(1, [int]::MaxValue)]
    [Alias('Id')]
    #The ID of the template organization to use as a base for the new organization, unspecified settings will be copied from here. You can also pipe in a template organization object.
    $TemplateOrganizationId,

    [NinjaOne.V2.Client]$Client = (Get-N1Client)
  )
  dynamicparam {
    $SCRIPT:modelParams = New-DynamicParametersFromTypeProperties $Model.GetType() -Mandatory Name
    $modelParams
  }
  begin {
    [bool]$SCRIPT:orgIdPipelineAlreadyReceived = $false
  }
  process {
    #HACK: This is a tricky way to support different types of objects on the pipeline. We want there to be one and only one item on the pipeline but we accept it via property name, but we can't determine this right now, so we have to save context and error if this activity gets repeated. when we get to end, it should have the var set to the "last" item in the pipeline.
    if ($orgIdPipelineAlreadyReceived) {
      $PSCmdlet.ThrowTerminatingError(
        [Management.Automation.ErrorRecord]::new(
          'Only one organization template can be provided via the pipeline.',
          'PipelineError',
          'InvalidOperation',
          $PSItem
        )
      )
    }
    if ($PSItem.Id) {
      #Assumed to be templateOrganizationId
      $orgIdPipelineAlreadyReceived = $true
      return
    }

    #Even if nothing was on the pipeline, process runs once, so we can use our byParam logic here to merge flow with the model
    if ($PSCmdlet.ParameterSetName -eq 'ByParams') {
      $createParams = @{}
      $boundKeys = $PSBoundParameters.Keys
      $boundKeys | Where-Object { $_ -in $modelParams.Keys }
      | ForEach-Object {
        $createParams.$PSItem = $PSBoundParameters[$PSItem]
      }
      $OrganizationModel = [OrganizationWithLocationsAndPolicyAssignmentsModel]$createParams
    }

    $shouldProcessMessage = $TemplateOrganizationId ? "Create a new organization based on template organization with Id $TemplateOrganizationId" : 'Create Organization'

    if (-not $PSCmdlet.ShouldProcess($OrganizationModel.ToString(), $shouldProcessMessage)) { return }
    $client.CreateOrganization($TemplateOrganizationId, $OrganizationModel)
    | Resolve-N1Error $TemplateOrganizationId
    | Add-Member -NotePropertyName 'Client' -NotePropertyValue $client -PassThru
  }
}

#Help purposefully removed for dynamic param help generation, where the parameters are dynamically generated from the model.
function Set-N1Organization {
  [CmdletBinding(DefaultParameterSetName = 'ByParams', SupportsShouldProcess, ConfirmImpact = 'High')]
  [OutputType([NinjaOne.V2.OrganizationDetailed])]
  param(
    [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'ByModel')]
    [NinjaOne.V2.OrganizationWithLocationsAndPolicyAssignmentsModel]$Model,

    [Parameter(ValueFromPipelineByPropertyName)]
    [ValidateRange(1, [int]::MaxValue)]
    [Alias('Id')]
    #The ID of the template organization to use as a base for the new organization, unspecified settings will be copied from here. You can also pipe in a template organization object.
    $TemplateOrganizationId,

    [NinjaOne.V2.Client]$Client = (Get-N1Client)
  )
  dynamicparam {
    $SCRIPT:modelParams = New-DynamicParametersFromTypeProperties $Model.GetType() -Mandatory Name
    $modelParams
  }
  begin {
    [bool]$SCRIPT:orgIdPipelineAlreadyReceived = $false
  }
  process {
    #HACK: This is a tricky way to support different types of objects on the pipeline. We want there to be one and only one item on the pipeline but we accept it via property name, but we can't determine this right now, so we have to save context and error if this activity gets repeated. when we get to end, it should have the var set to the "last" item in the pipeline.
    if ($orgIdPipelineAlreadyReceived) {
      $PSCmdlet.ThrowTerminatingError(
        [Management.Automation.ErrorRecord]::new(
          'Only one organization template can be provided via the pipeline.',
          'PipelineError',
          'InvalidOperation',
          $PSItem
        )
      )
    }
    if ($PSItem.Id) {
      #Assumed to be templateOrganizationId
      $orgIdPipelineAlreadyReceived = $true
      return
    }

    #Even if nothing was on the pipeline, process runs once, so we can use our byParam logic here to merge flow with the model
    if ($PSCmdlet.ParameterSetName -eq 'ByParams') {
      $createParams = @{}
      $boundKeys = $PSBoundParameters.Keys
      $boundKeys | Where-Object { $_ -in $modelParams.Keys }
      | ForEach-Object {
        $createParams.$PSItem = $PSBoundParameters[$PSItem]
      }
      $OrganizationModel = [OrganizationWithLocationsAndPolicyAssignmentsModel]$createParams
    }

    $shouldProcessMessage = $TemplateOrganizationId ? "Create a new organization based on template organization with Id $TemplateOrganizationId" : 'Create Organization'

    if (-not $PSCmdlet.ShouldProcess($OrganizationModel.ToString(), $shouldProcessMessage)) { return }
    $client.CreateOrganization($TemplateOrganizationId, $OrganizationModel)
    | Resolve-N1Error $TemplateOrganizationId
    | Add-Member -NotePropertyName 'Client' -NotePropertyValue $client -PassThru
  }
}

function Remove-N1Organization {
  [CmdletBinding(DefaultParameterSetName = 'List')]
  param(
    [Parameter(Mandatory, ParameterSetName = 'Get', ValueFromPipelineByPropertyName)][int]$Id,
    [ValidateNotNullOrEmpty()][Parameter(ValueFromPipelineByPropertyName)]
    [NinjaOne.V2.Client]$Client = (Get-N1Client)
  )
}

#endRegion Organization


#region Device
function Get-N1Device {
  <#
  .SYNOPSIS
  Get the current Device
  .NOTES
  #>
  [CmdletBinding(DefaultParameterSetName = 'List')]
  [OutputType([NinjaOne.V2.Device])]
  [OutputType([NinjaOne.V2.NodeWithDetailedReferences])]
  param(
    #Filter the results using the device filter syntax. More info: https://eu.ninjarmm.com/apidocs-beta/core-resources/articles/devices/device-filters
    [Parameter(ParameterSetName = 'List')][string]$Filter,
    [Parameter(ParameterSetName = 'List')]
    [Parameter(ParameterSetName = 'ByOrganization')]
    [ValidateRange(1, [int]::MaxValue)] #HACK: We dont type this value directly because PowerShell will not let it be null and will default to 0 instead of $null which we don't want.
    $AfterId,
    [Parameter(ParameterSetName = 'List')]
    [Parameter(ParameterSetName = 'ByOrganization')]
    [ValidateRange(1, [int]::MaxValue)]
    $Count,
    [Parameter(ParameterSetName = 'ByOrganization', ValueFromPipeline)]
    [NinjaOne.V2.Organization]$Organization,
    [Parameter(Mandatory, ParameterSetName = 'Get', ValueFromPipelineByPropertyName)][int]$Id,

    [ValidateNotNullOrEmpty()][Parameter(ValueFromPipelineByPropertyName)]
    [NinjaOne.V2.Client]$Client = (Get-N1Client)
  )
  process {
    if ($PSCmdlet.ParameterSetName -eq 'Get') {
      return $client.GetDevice($Id)
      | Resolve-N1Error $Id
      | Add-Member -NotePropertyName 'Client' -NotePropertyValue $client -PassThru
    }
    if ($PSCmdlet.ParameterSetName -eq 'ByOrganization') {
      $Id = $Organization.Id
      return $client.GetOrganizationDevices($Id, $Count, $AfterId)
      | Resolve-N1Error $Id
      | Add-Member -NotePropertyName 'Client' -NotePropertyValue $client -PassThru
    }
  }
  end {
    if ($PSCmdlet.ParameterSetName -eq 'List') {
      try {
        return $client.GetDevices($Filter, $Count, $Skip) #BUG: API does not return a Device but rather a NodeWithDetailedReferences
        | Resolve-N1Error $Id
        | Add-Member -NotePropertyName 'Client' -NotePropertyValue $client -PassThru
      } catch {
        #Response Errors are not parsed properly by the client library
        $PSItem | Resolve-N1Error
      }
    }
  }
}

#endRegion Device

class NinjaOneResponseError {
  [string]$resultCode
  [string]$errorMessage
  [string]$incidentId
}

filter Resolve-N1Error($Target) {
  $errorWhenResponseExpected = $PSItem.Exception?.GetBaseException().Path -eq 'resultCode'
  if (
    -not (
      $PSItem.AdditionalProperties.resultCode -or
      $errorWhenResponseExpected
    )
  ) { return $PSItem }

  if ($errorWhenResponseExpected) {
    $err = [Management.Automation.ErrorRecord]::new(
      $PSItem.Exception,
      'ResponseParseError',
      'InvalidOperation',
      $Target
    )
    $err.ErrorDetails = 'An error was received where a response was expected, as of now this module cannot read the error Code correctly, that will come in a later version. This is probably due to a poor API specification or a bad Filter syntax.'

    $PSCmdlet.ThrowTerminatingError($err)
  }

  [NinjaOneResponseError]$responseError = $PSItem.AdditionalProperties
  #Will get the calling functions PSCmdlet since this is not an advanced function.
  $err = [Management.Automation.ErrorRecord]::new(
    "$($responseError.resultCode): $($responseError.errorMessage)",
    $responseError.resultCode,
    'InvalidOperation', $Target
  )
  $err.Exception.Data.Add('incidentId', $responseError.incidentId)
  $PSCmdlet.ThrowTerminatingError($err)
}

function New-DynamicParametersFromTypeProperties {
  param(
    [type]$Type,
    [string[]]$Mandatory,
    [string[]]$Exclude,
    [string]$ParameterSetName = 'ByParams'
  )
  $paramDictionary = [RuntimeDefinedParameterDictionary]::new()
  $properties = $Type.DeclaredProperties | Where-Object Name -NotIn $Exclude
  foreach ($property in $properties) {
    $paramDictionary.Add(
      $property.Name,
      [RuntimeDefinedParameter]::new(
        $property.Name,
        $property.PropertyType,
        [ParameterAttribute]@{
          ParameterSetName = $ParameterSetName
          Mandatory        = $Mandatory -contains $property.Name
        }
      )
    )
  }
  return $paramDictionary
}
