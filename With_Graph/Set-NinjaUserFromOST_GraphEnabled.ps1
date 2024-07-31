# Function to get an access token
function Get-AccessToken {
    param (
        [string]$clientId,
        [string]$clientSecret,
        [string]$tenantId
    )

    $tokenUrl = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
    $tokenBody = @{
        client_id     = $clientId
        scope         = "https://graph.microsoft.com/.default"
        client_secret = $clientSecret
        grant_type    = "client_credentials"
    }

    try {
        $tokenResponse = Invoke-RestMethod -Uri $tokenUrl -Method Post -Body $tokenBody
        return $tokenResponse.access_token
    } catch {
        throw "Failed to get access token: $_"
    }
}

# Function to get user information
function Get-UserInfo {
    param (
        [string]$accessToken,
        [string]$userPrincipalName
    )

    $graphApiUrl = "https://graph.microsoft.com/v1.0/users/$userPrincipalName"
    $headers = @{
        Authorization = "Bearer $accessToken"
    }

    try {
        $userInfo = Invoke-RestMethod -Uri $graphApiUrl -Headers $headers -Method Get
        return $userInfo
    } catch {
        if ($_.Exception.Response.StatusCode -eq 404) {
            return $null
        }
        throw "Failed to get user information: $_"
    }
}

# Function to set Ninja custom field
function Set-NinjaCustomField {
    param (
        [string]$fieldName,
        [string]$fieldValue
    )

    try {
        & ninja-property-set $fieldName $fieldValue
    } catch {
        throw "Failed to set Ninja custom field: $_"
    }
}

# Function to get email UPN from Outlook OST file
function Get-UPNFromOST {
    param (
        [string]$userProfilePath
    )

    $outlookVersion = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -ErrorAction SilentlyContinue).VersionToReport
    if (-not $outlookVersion) {
        $outlookVersion = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Office\16.0\Outlook\InstallRoot" -ErrorAction SilentlyContinue).Path
    }
    
    if (-not $outlookVersion) {
        throw "Outlook installation not found"
    }

    $ostPath = Join-Path $userProfilePath "AppData\Local\Microsoft\Outlook"
    $ostFiles = Get-ChildItem -Path $ostPath -Filter "*.ost" -ErrorAction SilentlyContinue

    if (-not $ostFiles) {
        throw "No Outlook OST files found"
    }

    $freeEmailDomains = @("gmail.com", "yahoo.com", "hotmail.com", "outlook.com", "aol.com", "icloud.com")
    $validOsts = @()

    foreach ($ost in $ostFiles) {
        if ($ost.Name -match '.*?([^\s_]+@[^\s]+)') {
            $upn = $Matches[1] -replace '\.ost$', ''
            $domain = $upn.Split('@')[1]
            if ($freeEmailDomains -notcontains $domain) {
                $validOsts += [PSCustomObject]@{
                    UPN = $upn
                    LastWriteTime = $ost.LastWriteTime
                }
            }
        }
    }

    if ($validOsts.Count -eq 0) {
        throw "No valid company UPN found in OST files"
    }

    return $validOsts | Sort-Object LastWriteTime -Descending | Select-Object -First 1
}

# Function to get all user profiles
function Get-UserProfiles {
    $userProfiles = Get-ChildItem -Path "C:\Users" -Directory | Where-Object { $_.Name -ne "Public" -and $_.Name -ne "Default" }
    return $userProfiles
}

# Main script execution
try {
    # Your multi-tenant app registration details
    $clientId = "<multitenant app id>"
    $clientSecret = "<multitenant app secret>"
    $tenantId = "<app source tenant>"

    Write-Output "Script started. Attempting to find user profiles..."

    $userProfiles = Get-UserProfiles
    if ($userProfiles.Count -eq 0) {
        throw "No user profiles found in C:\Users"
    }

    Write-Output "Found $($userProfiles.Count) user profile(s)."

    $validProfiles = @()

    foreach ($profile in $userProfiles) {
        $userProfilePath = $profile.FullName
        $userName = $profile.Name

        Write-Output "Checking profile: $userName"

        try {
            $upnInfo = Get-UPNFromOST -userProfilePath $userProfilePath
            $validProfiles += [PSCustomObject]@{
                UserName = $userName
                UPN = $upnInfo.UPN
                LastWriteTime = $upnInfo.LastWriteTime
                ProfileLastWriteTime = $profile.LastWriteTime
            }
            Write-Output "Successfully detected UPN for $userName : $($upnInfo.UPN)"
        } catch {
            Write-Output "Failed to process profile $userName : $_"
        }
    }

    if ($validProfiles.Count -eq 0) {
        throw "No valid UPN found in any user profile"
    }

    # Sort profiles by OST last write time, then by profile folder last write time
    $mostRecentProfile = $validProfiles | Sort-Object -Property LastWriteTime, ProfileLastWriteTime -Descending | Select-Object -First 1

    Write-Output "Most recent active profile: $($mostRecentProfile.UserName) with UPN: $($mostRecentProfile.UPN)"

    # Get an access token
    $accessToken = Get-AccessToken -clientId $clientId -clientSecret $clientSecret -tenantId $tenantId

    # Try to get user information with the UPN from OST
    $userInfo = Get-UserInfo -accessToken $accessToken -userPrincipalName $mostRecentProfile.UPN

    if ($null -eq $userInfo) {
        throw "User not found with UPN: $($mostRecentProfile.UPN)"
    }

    # Set the Ninja custom field
    Set-NinjaCustomField -fieldName "lastDetectedUser" -fieldValue $userInfo.userPrincipalName

    Write-Output "Successfully set lastDetectedUser to $($userInfo.userPrincipalName)"
} catch {
    $errorMessage = $_.Exception.Message
    Write-Error $errorMessage
    
    # Attempt to set the custom field with the error message
    try {
        Set-NinjaCustomField -fieldName "lastDetectedUser" -fieldValue "Error: $errorMessage"
    } catch {
        Write-Error "Failed to set error message in custom field: $_"
    }
    
    exit 1
}