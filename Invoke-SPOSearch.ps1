#Requires -Modules @{ ModuleName="PnP.PowerShell"; RequiredVersion="2.3.0" }
#Requires -Version 7.2

[CmdletBinding()]
param (
    [Parameter(Mandatory, Position = 0)]
    [String]
    $QueryString,

    [Parameter()]
    [ValidateRange(1, 500)]
    [Int]
    $PageSize = 500
)

begin {

    [Console]::ResetColor()

    try {
        $null = Get-PnPTenantInfo -CurrentTenant -OutVariable tenantInfo
    }
    catch {
        $_.Exception.Message | Out-Default
        continue
    }

    $lastDocId = 0
    $result = [System.Collections.Generic.List[System.Object]]@()
}
process {
    do {
        $search = Submit-PnPSearchQuery -Query "$queryString IndexDocId>$lastDocId" -SortList @{ "[DocId]" = "ascending" } -MaxResults $pageSize
        if ($search.RowCount -gt 0) {
            "$($search.TotalRows) items remaining" | Out-Default
            foreach ($row in $search.ResultRows) {
                $result.Add((
                        New-Object psobject -Property ([ordered]@{
                                SiteUrl          = "$($row.SiteName)"
                                ParentPath       = "$($row.ParentLink)"
                                FullPath         = "$($row.ParentLink)/$($row.Title).$($row.FileType)"
                                Filename         = "$($row.Title).$($row.FileType)"
                                SizeKB           = "$($row.Size / 1KB)"
                                LastModifiedTime = "$(Get-Date $($row.LastModifiedTime) -Format "yyyy-MM-dd HH:mm:ss")"
                                DocId            = "$($row.DocId)"
                                OwnerName        = ""
                                OwnerEmail       = ""
                                OwnerLoginName   = ""
                            })
                    ))
            }
            $lastDocId = $search.ResultRows[-1].DocId
        }
    }
    while ($search.RowCount -gt 0)
}
end {
    ## Filter unique site URLs
    $uniqueSites = (($result).SiteUrl | Select-Object -Unique)

    ## Compose the site owners list
    "Getting site owner details..." | Out-Default
    $siteOwnerList = [System.Collections.Generic.List[System.Object]]@()
    foreach ($siteUrl in $uniqueSites) {
        $siteInfo = Get-PnPTenantSite -Identity $siteUrl | Select-Object Owner*
        $siteOwnerList.Add((
                New-Object psobject -Property ([ordered]@{
                        SiteUrl        = $siteUrl
                        OwnerName      = $siteInfo.OwnerName
                        OwnerEmail     = $siteInfo.OwnerEmail
                        OwnerLoginName = ($siteInfo.OwnerLoginName).Split("|")[-1]
                    })
            ))
    }

    ## Update the result with owner information
    "Adding site owner details to the results..." | Out-Default
    foreach ($item in $result) {
        $currentSite = $siteOwnerList | Where-Object { $_.SiteUrl -eq $item.SiteUrl }
        $item.OwnerName = $currentSite.OwnerName
        $item.OwnerEmail = $currentSite.OwnerEmail
        $item.OwnerLoginName = $currentSite.OwnerLoginName
    }

    ## Return the results
    $result
}






