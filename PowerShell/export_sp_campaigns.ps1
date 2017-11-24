# This script creates site collections and/or sub webs from csv configuration template.

# Change the policy settings if needed.
#$policy = Get-ExecutionPolicy
#if ($policy -ne 'RemoteSigned') {
#    Set-ExecutionPolicy RemoteSigned
#}

if (-not (Get-Module -ListAvailable -Name SharePointPnPPowerShellOnline)) 
{
    Install-Module SharePointPnPPowerShellOnline
}

Import-Module SharePointPnPPowerShellOnline

$tenantUrl="https://insidemedia.sharepoint.com/sites/mec_mb_dashboard"

Connect-PnPOnline -Url $tenantUrl

Write-Host "Provisioning site collection $siteUrl" -ForegroundColor Yellow

$fields = @{Name="ID"; Expression = {$_.ID}},@{Name="Campaign"; Expression = {$_.Title}},@{Name="Year"; Expression = {$_._x5b57__x6bb5_1}},@{Name="Product"; Expression = {$_.Product}}

$out_path = "C:\Digital_Report_Platform\source\mectech\datamart\mercedes\temp_files\mercedes_sp_campaign\"

$cmpgs = (Get-PnPListItem -List CampaignInfo).FieldValues

$cmpgs|Select-Object -Property $fields|Export-Csv  -Path $out_path"sp_campaign.csv" -Force -Encoding UTF8 -notype


$fields = @{Name="VideoId"; Expression = {$_.LookupId}}

# export video usage
$output = @()
$(foreach($each in $cmpgs){
	if ($each.Video){
		
		foreach($ev in $each.Video){
			$obj = "" | select cmpgID, videoID
			$obj.cmpgID = $each.ID
			$obj.videoID = $ev.LookupID
			$output += $obj
		}
	}
})
$output|export-csv -Path $out_path"sp_campaign_videos.csv" -Force -Encoding UTF8 -notype

# export image usage
$output = @()
$(foreach($each in $cmpgs){
	if ($each.Image){
		
		foreach($ev in $each.Image){
			$obj = "" | select cmpgID, imageID
			$obj.cmpgID = $each.ID
			$obj.imageID = $ev.LookupID
			$output += $obj
		}
	}
})
$output|export-csv -Path $out_path"sp_campaign_images.csv" -Force -Encoding UTF8 -notype

# export video & image

$fields = @{Name="ID"; Expression = {$_.ID}},@{Name="Name"; Expression = {$_.Title}},@{Name="FileRef"; Expression = {$_.FileRef}}

(Get-PnPListItem -List "KV Videos").FieldValues|Select-Object -Property $fields|Export-Csv  -Path $out_path"sp_video.csv" -Force -Encoding UTF8 -notype

(Get-PnPListItem -List "KV Images").FieldValues|Select-Object -Property $fields|Export-Csv  -Path $out_path"sp_image.csv" -Force -Encoding UTF8 -notype