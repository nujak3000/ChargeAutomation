###Helps adjust amp draw for Tesla charge based on Enphase solar production.
##########SCRIPT PREP###########
###Follow the instructions here to generate a Tesla auth token and refresh token. The auth 
###token will expire in a day and the refresh token will not expire until you reset your 
###tesla account password.
###https://tesla-info.com/tesla-token.php
###Tesla doesn't have a documented/supported API. This is a 3rd party that reversed engineered the Tesla app.
###
###This script requires Enphase's envoy.local to be available. envoy.local API doesn't require login .
###
###Schedule to run in task scheduler every 5 minutes or less.
#################################
############################store the token here####
$refreshtoken = get-content C:\secure\token.txt
$URL_Prefix = "https://tesla-info.com/api/control_v2.php?refresh=$refreshtoken&request="

###############################################################################Logging function####
function Write-Log
{
<#
.SYNOPSIS Create well formatted record of messages for troubleshooting
.PARAMETER Message
.EXAMPLE
Write-Log -Message 'Value1'
#>
    [CmdletBinding()]
    param (
    [Parameter(Mandatory)]
    [string]$Message
    )

    try
    {
    if (!(Test-Path "c:\secure\runlog.txt"))
        {
            new-item -Path "c:\secure\runlog.txt"
        }
        $DateTime = Get-Date -Format 'MM-dd-yy HH:mm:ss'
        $Invocation = $MyInvocation.ScriptLineNumber
        Add-Content -Value "$DateTime - $PSCommandPath - $Invocation - $Message" -path "c:\secure\runlog.txt"
        Write-host "$DateTime - $PSCommandPath - $Invocation - $Message"
    }
    catch
    {
        Write-Error $_.Exception.Message
    }
}

###############################################################################check for night hours before calling any APIs

$current_hour = get-date -Format "HH"

if($current_hour -gt 17 -and $current_hour -lt 8)
{
    Write-Log -Message "No sun at night, exiting"
    exit
}
else
{
    write-host "Party time"
}

###############################################################################check car status

$command = "get_charge"

$posturl = $URL_Prefix + $command
$tesla_WebResponse = Invoke-WebRequest $posturl
$tesla_psResponse = $tesla_WebResponse.Content | ConvertFrom-Json
Write-Log -Message " battery_level  $($tesla_psResponse.response.battery_level)"
Write-Log -Message " charge_amps  $($tesla_psResponse.response.charge_amps)"
Write-Log -Message " charging_state  $($tesla_psResponse.response.charging_state)"

#possible charging states
#"charging_state":"Charging"
#"charging_state":"Disconnected"
#"charging_state":"Stopped"
#"charging_state":"Complete"

if($tesla_psResponse.response.charging_state -eq "Disconnected")
{
    Write-Log -Message "Charger disconnected, exiting script"
    Start-Sleep 5
    exit
}

###############################################################################Check solar status
try 
{
    #if run every 5 minutes, this seems to fail a few times a day for some reason
    $WebResponse = Invoke-WebRequest "http://envoy.local/production.json?details=1"
}
catch
{
    write-log -message $_.Exception
    start-sleep 5
    exit
}
$psResponse = $WebResponse.Content | ConvertFrom-Json

$culture = Get-Culture

if($tesla_psResponse.response.charging_state -ne "Charging")
{
    $CurrentChargeRateAmps = 0
}
else
{
    $CurrentChargeRateAmps = [decimal]::Parse($tesla_psResponse.response.charge_amps,$culture)
}

$curcuitVolts = [decimal]::Parse(240,$culture)
$currentChargeWatts = $CurrentChargeRateAmps * $curcuitVolts
Write-Log -Message  "current charging watts:  $currentChargeWatts"

###Convert strings to decimals so math always works
$decProductionWatts = [decimal]::Parse($psResponse.production[0].wNow, $culture)
$decConsumptionWatts = [decimal]::Parse($psResponse.consumption[0].wNow, $culture)
$decNetConsumptionWatts = [decimal]::Parse($psResponse.consumption[1].wNow, $culture)
$decCurrentChargeWatts = [decimal]::Parse($currentChargeWatts, $culture)

Write-Log -Message  "production:  $decProductionWatts"
Write-Log -Message  "consumption:  $decConsumptionWatts"
Write-Log -Message  "net:  $decNetConsumptionWatts"
$wattsLessCharging = $decNetConsumptionWatts - $decCurrentChargeWatts
Write-Log -Message  "net less charging:  $wattsLessCharging"

if($decNetConsumptionWatts -lt 0)
{
    Write-Log -Message  "producing"

    $amps = $wattsLessCharging / 240
    $amps = [math]::abs($amps) 
    $amps = [math]::Round($amps,0)
    $amps = $amps + 1

    Write-Log -Message  "set charging amps to:  $amps"

}
else
{
    $amps = ($decCurrentChargeWatts - $decNetConsumptionWatts) / 240
    $amps = [math]::abs($amps) 
    $amps = [math]::Round($amps,0)
    $amps = $amps + 1
    Write-Log -Message  "consuming, cannot charge more. Set charging amps to:  $amps"
}
###############################################################################END CHECK SOLAR STATUS

#could use this data to decide to use peak hours
#$weather = (curl http://wttr.in/Detroit?0 -UserAgent "curl" ).Content

#avoid peak rates and low amp charges, set to 0
$overridepeak = 1

###############################################################################MAIN LOGIC: Stop charge, set amps and/or start charge
$current_hour = get-date -Format "HH"

if(($amps -lt 5 -and $tesla_psResponse.response.charging_state -eq "Charging") -or ($overridepeak = 0 -and $current_hour -gt 13 -and $current_hour -lt 20 -and $tesla_psResponse.response.charging_state -eq "Charging" -and $tesla_psResponse.response.battery_level -gt 30))
{
    ####STOP CHARGING
    Write-Log -Message  "stop charging - peak rate hours or not enough production"
    $command = "charge_stop"
    $posturl = $URL_Prefix + $command
    Invoke-WebRequest $posturl
}
elseif($tesla_psResponse.response.charging_state -eq "Charging" -and $amps -gt 4 -and $amps -ne $tesla_psResponse.response.charge_amps) 
{
    $command = "set_charging_amps&value=$amps"
    $posturl = $URL_Prefix + $command
    $WebResponse = Invoke-WebRequest $posturl
    $psResponse = $WebResponse.Content | ConvertFrom-Json
    Write-Log -Message $psResponse.cause
}
elseif($tesla_psResponse.response.charging_state -eq "Stopped")
{
    $command = "set_charging_amps&value=$amps"
    $posturl = $URL_Prefix + $command
    $WebResponse = Invoke-WebRequest $posturl
    $psResponse = $WebResponse.Content | ConvertFrom-Json
    Write-Log -Message $psResponse.cause

    $command = "charge_start"

    $posturl = $URL_Prefix + $command
    $WebResponse = Invoke-WebRequest $posturl
    $psResponse = $WebResponse.Content | ConvertFrom-Json
    Write-Log -Message $psResponse.cause
}

###############################################################################cleanup variables and wait to display on screen
$refreshtoken = ""
$URL_Prefix = ""
$posturl = ""
Start-Sleep 5