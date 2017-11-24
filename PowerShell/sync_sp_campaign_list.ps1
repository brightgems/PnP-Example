#-------------------------------------------------------------------- 
# Name: Load CSV into SharePoint List 
# NOTE: No warranty is expressed or implied by this code – use it at your 
# own risk. If it doesn't work or breaks anything you are on your own 
#--------------------------------------------------------------------


if (-not (Get-Module -ListAvailable -Name SharePointPnPPowerShellOnline)) 
{
    Install-Module SharePointPnPPowerShellOnline
}

Import-Module SharePointPnPPowerShellOnline

$tenantUrl="https://insidemedia.sharepoint.com/sites/mec_mb_dashboard"

Connect-PnPOnline -Url $tenantUrl
# read campaign from azure sql
$params = @{

  'Database' = 'datamart'

  'ServerInstance' =  'mecshsql.database.windows.net,1433'

  'Username' = 'mercedes'

  'Password' = '1qaz!QAZ'

  'OutputSqlErrors' = $true

  'Query' = "select distinct CampYear,Campaign,Product from [dbo].[Campaign_Mercedes] where isnull(Campaign,'')<>''"

  }

$Campaigns = Invoke-Sqlcmd  @params

# read sp campaigns
$fields = @{Name="ID"; Expression = {$_.ID}},@{Name="Campaign"; Expression = {$_.Title}},@{Name="Year"; Expression = {$_._x5b57__x6bb5_1}},@{Name="Product"; Expression = {$_.Product}}

$cmpgs = (Get-PnPListItem -List CampaignInfo).FieldValues |Select-Object -Property $fields

# Loop through campaign add each one to SharePoint

"Uploading data to SharePoint...."

foreach ($row in $Campaigns) {
   $key = $cmpgs |Where-Object {$_.Campaign -eq $row.Campaign.ToString() -and $_.Year.ToString() -eq $row.CampYear.ToString()}
   
   Write-Host $key
   if ($key -eq $null){
	   "Adding entry for "+$row.Campaign.ToString() 
	   $spItem = @{"Title" = $row.Campaign.ToString() ; "_x5b57__x6bb5_1"=$row.CampYear.ToString();"Product"=$row.Product.ToString()}
	   Add-PnPListItem -List "CampaignInfo" -ContentType Item -Values $spItem

   }
   ElseIf($key.Product -ne $row.Product){
	   $spItem = @{"Product"=$row.Product.ToString()}
	   Set-PnPListItem -Identity $key.ID.ToString()  -List "CampaignInfo" -ContentType Item -Values $spItem
   }
}

"---------------" 
"Upload Complete"
