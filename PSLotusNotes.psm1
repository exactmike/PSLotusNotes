function New-NotesDatabaseConnection
{[cmdletbinding()]
  param(
    [string]$NotesServerName
    ,
    [string]$Database #the Notes nsf file name to be accessed
    ,
    $Credential #the password will be decrypted
    ,
    [string]$Name # An arbitrary friendly name for the notes database
    ,
    [string]$Identity #An arbitrary session name for the Notes Session
    ,
    [switch]$ServerConnectionWithSpecifiedUserName #use if you are using this on a machine with domino server installed and want/need to specify a specific user. Note:untested as I don't have access to a domino server
  )
  #If called from OneShell the SessionIdentity is a GUID and we want to remove the '-' characters.
  if ($MyInvocation.PSCommandPath -like '*OneShell*')
  {$SessionIdentity = $Identity.Replace('-','')}
  else
  {
    $SessionIdentity = $Identity
  }
  $Password = $Credential.Password | Convert-SecureStringToString
  $UserName = $Credential.Username
  if (-not (Test-Path -Path variable:NotesSessions))
  {
    New-Variable -Name NotesSessions -Value @{} -Scope Global
  }
  if (-not (Test-Path -Path variable:NotesDatabaseConnections))
  {
    New-Variable -Name NotesDatabaseConnections -Value @{} -Scope Global
  }
  if (-not ($NotesSessions.ContainsKey($SessionIdentity)))
  {
    $NotesSessions.$SessionIdentity = New-Object -ComObject 'Lotus.NotesSession'
    if ($ServerConnectionWithSpecifiedUserName)
    {
      $NotesSessions.$SessionIdentity.InitializeUsingNotesUserName("$UserName","$Password")
    }
    else
    {
      $NotesSessions.$SessionIdentity.Initialize("$Password")
    }
    if (-not ($NotesDatabaseConnections.ContainsKey($Name)))
    {
        $NotesDatabaseConnections.$Name = $NotesSessions.$SessionIdentity.GetDatabase("$NotesServerName","$Database")
    }
  }
  Write-Output -InputObject $NotesDatabaseConnections.$Name
}
function Get-NotesUser
{
  [cmdletbinding()]
  param(
    [parameter()]
    [string[]]$NotesDatabase
    ,
    [parameter(ParameterSetName = 'ByIdentity',Mandatory)]
    [string[]]$Identity
    ,
    [parameter(ParameterSetName = 'All')]
    [switch]$All
    ,
    [parameter(ParameterSetName = 'All',Mandatory)]
    [string[]]$Property
  )
Begin
{
  #Create the Views Hashtable if needed
  if (-not (Test-Path -Path variable:Global:NotesViews))
  {
    New-Variable -Name NotesViews -Value @{} -Scope Global
  }
  #Create the View Object(s) if needed
  switch ($PSCmdlet.ParameterSetName)
  {
    'ByIdentity'
    {
      #Setup the Users View if needed
      foreach ($ND in $NotesDatabase)
      {
        $DatabaseView = "$($ND)Users"
        if (-not ($NotesViews.ContainsKey($DatabaseView)))
        {
            $NotesViews.$DatabaseView = $NotesDatabaseConnections.$ND.GetView('($Users)')
        }
      }
    }#ByName
    'All'
    {
      #Setup the People View if needed
      foreach ($ND in $NotesDatabase)
      {
        $DatabaseView = "$($ND)People"
        if (-not ($NotesViews.ContainsKey($DatabaseView)))
        {
            $NotesViews.$DatabaseView = $NotesDatabaseConnections.$ND.GetView('People')
        }
      }#foreach
    }#All
  }#switch
}#Begin
Process
{
  switch ($PSCmdlet.ParameterSetName)
  {
    'ByIdentity'
    {
      foreach ($ID in $Identity)
      {
        $userdocs = @()
        foreach ($ND in $NotesDatabase)
        {
          $userdoc = @($NotesViews.$DatabaseView.GetDocumentByKey($ID) | Where-Object -FilterScript {$_ -ne $null})
          switch ($userdoc.Count)
          {
            1
            {
                $userdocs += $userdoc
            }
            0
            {}
            default
            {
                throw "$ID is ambiguous in `$ND"
            }
          }
        }
        switch ($userdocs.Count)
        {
          0
          {Write-Warning -Message "No Notes User for $ID was found"}
          default
          {
              throw "$ID is ambiguous among Notes Databases: $($NotesDatabase -join ',')"
          }
          1
          {
              $rawNotesUserdoc = $userdocs[0]
              $NotesUserObject = [pscustomobject]@{}
              foreach ($item in $($rawNotesUserdoc.Items | Sort-Object -Property Name))
              {
                  if ($NotesUserObject.psobject.members.GetEnumerator().Name -notcontains $item.name)
                  {
                    $NotesUserObject | Add-Member -Name $($item.name) -value $(if ($item.values.count -gt 1) {$item.text} else {$item.values}) -MemberType NoteProperty
                  }
              }
              Write-Output -InputObject $NotesUserObject
          }
        }
      }#foreach ID in Identity
    }#ByIdentity
    'All'
    {
      $NotesUserObjects = @(
        foreach ($ND in $NotesDatabase)
        {
          $RecordCount = $NotesViews.$DatabaseView.EntryCount
          $userdoc = $NotesViews.$DatabaseView.GetFirstDocument()
          $count = 0

            While ($userdoc -ne $null)
            {
              $Count++
              $nuh = @{}
              foreach ($prop in $Property)
              {
                $nuh.$prop = $userdoc.GetItemValue($prop)
              }
              $nuo = New-Object -TypeName PSCustomObject -Property $nuh
              Write-Output -InputObject $nuo
              $userdoc = $NotesViews.$DatabaseView.GetNextDocument($userdoc)
            }
        }#foreach
      )#NotesUserObjects
    } #All
  }
}#Process
}
function Convert-SecureStringToString
{
    <#
        .SYNOPSIS
        Decrypts System.Security.SecureString object that were created by the user running the function.  Does NOT decrypt SecureString Objects created by another user. 
        .DESCRIPTION
        Decrypts System.Security.SecureString object that were created by the user running the function.  Does NOT decrypt SecureString Objects created by another user.
        .PARAMETER SecureString
        Required parameter accepts a System.Security.SecureString object from the pipeline or by direct usage of the parameter.  Accepts multiple inputs.
        .EXAMPLE
        Decrypt-SecureString -SecureString $SecureString
        .EXAMPLE
        $SecureString1,$SecureString2 | Decrypt-SecureString
        .LINK
        This function is based on the code found at the following location:
        http://blogs.msdn.com/b/timid/archive/2009/09/09/powershell-one-liner-decrypt-securestring.aspx
        .INPUTS
        System.Security.SecureString
        .OUTPUTS
        System.String
    #>

    [cmdletbinding()]
    param (
        [parameter(ValueFromPipeline=$True)]
        [securestring[]]$SecureString
    )
    
    BEGIN {}
    PROCESS {
        foreach ($ss in $SecureString)
        {
          if ($ss -is 'SecureString')
          {[Runtime.InteropServices.marshal]::PtrToStringAuto([Runtime.InteropServices.marshal]::SecureStringToBSTR($ss))}
        }
    }
    END {}
}
