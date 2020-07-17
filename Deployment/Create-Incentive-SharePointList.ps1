#Requires -Version 5.0
#Script Parameters
[cmdletbinding(SupportsShouldProcess=$True)]Param(
  [Parameter(Mandatory=$True, ValueFromPipeline=$true, ParameterSetName="currentSite")][string]$siteURL,
  [Parameter(Mandatory=$True, ValueFromPipeline=$true, ParameterSetName="newSite"    )][string]$baseURL,
  [Parameter(Mandatory=$True, ValueFromPipeline=$false, ParameterSetName="newSite"    )][string]$siteName
)

#Check Which Version of PowerShell is being used
if($host.version -ge "7.0.0"){
  $loadModuleCmd = "import-module -Name SharePointPnPPowerShellOnline -UseWindowsPowerShell"
}else{
  Write-Warning "Using PowerShell version: $($host.version); it is recommended to upgrade to 7.0.0 or above"
  $loadModuleCmd = "import-module -Name SharePointPnPPowerShellOnline"
}

#Making sure the user has this Package installed.
if(-not (Get-Module -ListAvailable -Name "SharePointPnPPowerShellOnline")){
  Install-Module SharePointPnPPowerShellOnline -Scope User
}

#Load the SharePointOnline Module
Invoke-Expression -command $loadModuleCmd

#When user enters $siteURL then it adds the list inside that site url
if($siteURL){
  $incentivesiteURL = $siteURL
}

#when user enters $baseURL and $siteName then it will create a new SharePoint site for the list
if($baseURL){
  #Connects to the sharepoint base URL
  Connect-PnPOnline -Url $baseURL -UseWebLogin

  #Creates the Data Store Incentives App 
  #Title and alias for new site needs to be unique
  $incentivesiteURL = New-PnPSite -Type TeamSite -Title $siteName -Alias $siteName
}

#Switchs to the newly created site
Connect-PnPOnline -Url $incentivesiteURL -UseWebLogin

<#####################################################################################
Creates a new list for Incentives
#####################################################################################>
New-PnPList -Title 'Incentives' -Template GenericList -Url Lists/Incentives -ErrorAction Continue

#Adds new fields(columns) for Incentives
#Add-PnPField -List "Incentives" -DisplayName "Title" -InternalName "Title" -Type Text -AddToDefaultView #This is not required as it is created automatically
Add-PnPField -List "Incentives" -DisplayName "Name" -InternalName "Name" -Type Text -AddToDefaultView
Add-PnPField -List "Incentives" -DisplayName "Points" -InternalName "Points" -Type Number -AddToDefaultView
Add-PnPField -List "Incentives" -DisplayName "IncentiveCode" -InternalName "IncentiveCode" -Type Text -AddToDefaultView
Add-PnPField -List "Incentives" -DisplayName "CreatedBy" -InternalName "CreatedBy" -Type Text -AddToDefaultView
Add-PnPField -List "Incentives" -DisplayName "CreatedOn" -InternalName "CreatedOn" -Type DateTime -AddToDefaultView
Add-PnPField -List "Incentives" -DisplayName "DueDate" -InternalName "DueDate" -Type Date -AddToDefaultView
Add-PnPField -List "Incentives" -DisplayName "UpdatedBy" -InternalName "UpdatedBy" -Type Text -AddToDefaultView
Add-PnPField -List "Incentives" -DisplayName "UpdatedOn" -InternalName "UpdatedOn" -Type DateTime -AddToDefaultView

#Attachments column is present by default in sharepoint lists, but needs to be visible
$incentivesFields = (Get-PnPView -List "Incentives").ViewFields
If($incentivesFields -notcontains "Attachments"){
  $incentivesFields += "Attachments"
  Set-PnPView -List "Incentives" -Fields $incentivesFields -Identity (Get-PnPView -List "Incentives").Id
}

Add-PnPField -List "Incentives" -DisplayName "IsIncentiveLive" -InternalName "IsIncentiveLive" -Type Boolean -AddToDefaultView


<#####################################################################################
Creates a new list for Rewards
#####################################################################################>
New-PnPList -Title 'Rewards' -Template GenericList -Url Lists/Rewards -ErrorAction Continue

#Adds new fields for Rewards
#Add-PnPField -List "Rewards" -DisplayName "Title" -InternalName "Title" -Type Text -AddToDefaultView #This is not required as it is created automatically
Add-PnPField -List "Rewards" -DisplayName "Name" -InternalName "Name" -Type Text -AddToDefaultView
Add-PnPField -List "Rewards" -DisplayName "Points" -InternalName "Points" -Type Number -AddToDefaultView
Add-PnPField -List "Rewards" -DisplayName "CreatedBy" -InternalName "CreatedBy" -Type Text -AddToDefaultView
Add-PnPField -List "Rewards" -DisplayName "CreatedOn" -InternalName "CreatedOn" -Type DateTime -AddToDefaultView
Add-PnPField -List "Rewards" -DisplayName "UpdatedBy" -InternalName "UpdatedBy" -Type Text -AddToDefaultView
Add-PnPField -List "Rewards" -DisplayName "UpdatedOn" -InternalName "UpdatedOn" -Type DateTime -AddToDefaultView

#Attachments column is present by default in sharepoint lists, this may show error like "Attachments field exists"
#Need to click on add column inside list and goto Show/hide columns and make Attachments visible
Add-PnPField -List "Rewards" -DisplayName "Attachments" -InternalName "Attachments" -Type Attachments -AddToDefaultView


<#####################################################################################
Creates a new list for UserIncentives
#####################################################################################>
New-PnPList -Title 'UserIncentives' -Template GenericList -Url Lists/UserIncentives -ErrorAction Continue

#Adds new fields for UserIncentives
#Add-PnPField -List "UserIncentives" -DisplayName "Title" -InternalName "Title" -Type Text -AddToDefaultView #This is not required as it is created automatically
Add-PnPField -List "UserIncentives" -DisplayName "UserName" -InternalName "UserName" -Type Text -AddToDefaultView
Add-PnPField -List "UserIncentives" -DisplayName "IncentiveName" -InternalName "IncentiveName" -Type Text -AddToDefaultView
Add-PnPField -List "UserIncentives" -DisplayName "IncentivePoints" -InternalName "IncentivePoints" -Type Number -AddToDefaultView
Add-PnPField -List "UserIncentives" -DisplayName "IncentiveCode" -InternalName "IncentiveCode" -Type Text -AddToDefaultView
Add-PnPField -List "UserIncentives" -DisplayName "ReceivedOn" -InternalName "ReceivedOn" -Type DateTime -AddToDefaultView
Add-PnPField -List "UserIncentives" -DisplayName "ReasonForEdit" -InternalName "ReasonForEdit" -Type Text -AddToDefaultView


<#####################################################################################
Creates a new list for UserRewards
#####################################################################################>
New-PnPList -Title 'UserRewards' -Template GenericList -Url Lists/UserRewards -ErrorAction Continue

#Adds New Fields for UserRewards
#Add-PnPField -List "UserRewards" -DisplayName "Title" -InternalName "Title" -Type Text -AddToDefaultView #This is not required as it is created automatically
Add-PnPField -List "UserRewards" -DisplayName "UserName" -InternalName "UserName" -Type Text -AddToDefaultView
Add-PnPField -List "UserRewards" -DisplayName "RewardName" -InternalName "RewardName" -Type Text -AddToDefaultView
Add-PnPField -List "UserRewards" -DisplayName "RewardPoint" -InternalName "RewardPoint" -Type Number -AddToDefaultView
Add-PnPField -List "UserRewards" -DisplayName "Status" -InternalName "Status" -Type Text -AddToDefaultView
Add-PnPField -List "UserRewards" -DisplayName "VoucherCode" -InternalName "VoucherCode" -Type Text -AddToDefaultView
Add-PnPField -List "UserRewards" -DisplayName "CreatedOn" -InternalName "CreatedOn" -Type DateTime -AddToDefaultView

#Closes the connection
Disconnect-PnPOnline