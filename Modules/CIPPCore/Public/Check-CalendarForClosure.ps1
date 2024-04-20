function Check-CalendarForClosure {
    param (
        [Parameter(Mandatory=$true)]
        [string]$AccessToken,

        [string]$CalendarId = "d6f9cb8bd5da494781011da1f75051aa856952238952057021",

        [string]$DateToCheck = (Get-Date).Date.ToString("yyyy-MM-dd") # Default to today's date
    )

    $headers = @{
        "Authorization" = "Bearer $AccessToken"
        "Content-Type"  = "application/json"
    }

    $url = "https://graph.microsoft.com/v1.0/me/calendars/$CalendarId/events?`$filter=start/dateTime le '$DateToCheck' and end/dateTime ge '$DateToCheck' and contains(subject, 'Office Closed')"

    try {
        $response = Invoke-RestMethod -Uri $url -Headers $headers -Method Get
        if ($response.value.Count -gt 0) {
            return $true # Office is closed today
        } else {
            return $false # No closure events today
        }
    } catch {
        Write-Host "Error querying calendar: $($_.Exception.Message)"
        return $false
    }
}

function Update-CIPPOutOfOffice {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]$CheckEmail, # Email to check for OoO

        [Parameter(Mandatory=$true)]
        [string]$ApplyEmail, # Email to apply OoO to

        [Parameter(Mandatory=$true)]
        [string]$TenantFilter,

        [Parameter(Mandatory=$true)]
        [string]$ExecutingUser,

        [Parameter(Mandatory=$true)]
        [string]$AccessToken,

        [string]$NewState = 'Scheduled',

        [string]$NewStartTime,

        [string]$NewEndTime,

        [string]$ManualMessage, # Manually set OoO message

        [string]$AdditionalOptions # Optional parameter for additional commands or scripts
    )

    $fallbackMessage = "Thanks for submitting a helpdesk ticket.`r`n`r`n" +
                       "We are currently closed but will be back responding to tickets at 8am on our next business day. We appreciate your patience.`r`n`r`n" +
                       "For urgent situations (such as a company-wide outage or suspected security breach) we have a dedicated team on standby 7 days a week.`r`n`r`n" +
                       "Please reply back to this message with URGENT in the subject line, include any additional details and contact information and we will be in touch as soon as possible.`r`n`r`n" +
                       "Just a heads up, there might be an extra charge as after-hours rates may apply.`r`n`r`n" +
                       "Sincerely,`r`nSimplePowerIT Help Desk Team"

    $officeClosed = Check-CalendarForClosure -AccessToken $AccessToken -CalendarId "d6f9cb8bd5da494781011da1f75051aa856952238952057021" -DateToCheck (Get-Date).Date.ToString("yyyy-MM-dd")

    if ($officeClosed) {
        $fallbackMessage = "Office is closed today. Please contact us on the next business day."
    }

    $currentSettings = Get-CIPPOutOfOffice -userid $CheckEmail -TenantFilter $TenantFilter -ExecutingUser $ExecutingUser
    if ($currentSettings -is [String]) {
        $newInternalMessage = $fallbackMessage
        $newExternalMessage = $fallbackMessage
    } else {
        $currentSettings = $currentSettings | ConvertFrom-Json
        $newInternalMessage = $currentSettings.InternalMessage
        $newExternalMessage = $currentSettings.ExternalMessage
    }

    # Apply manual message if provided
    if ($ManualMessage) {
        $newInternalMessage = $ManualMessage
        $newExternalMessage = $ManualMessage
    }

    $today = Get-Date
    $dayOfWeek = $today.DayOfWeek
    $isWeekend = $dayOfWeek -eq 'Saturday' -or $dayOfWeek -eq 'Sunday'
    $defaultStartTime = $today.Date.AddHours(17) # 5 PM local time
    $defaultEndTime = $today.Date.AddHours(32) # 8 AM next day

    # Apply weekend schedule if no specific time provided
    if ($isWeekend) {
        $NewStartTime = $today.Date
        $NewEndTime = $today.Date.AddDays(1)
    } else {
        $NewStartTime = $NewStartTime -ne $null ? $NewStartTime : $defaultStartTime
        $NewEndTime = $NewEndTime -ne $null ? $NewEndTime : $defaultEndTime
    }

    $result = Set-CIPPOutOfOffice -userid $ApplyEmail -InternalMessage $newInternalMessage -ExternalMessage $newExternalMessage -TenantFilter $TenantFilter -State $NewState -ExecutingUser $ExecutingUser -StartTime $NewStartTime -EndTime $NewEndTime

    # Optionally run additional scripts or commands
    if ($AdditionalOptions) {
        Invoke-Expression $AdditionalOptions
    }

    return $result
}
