$LOGIN_URL = "https://sso-cas.it.utsa.edu/cas/login?service=https%3A%2F%2Fdawson2.utsarr.net%2Fcomal%2Fosp%2Fpages%2Fcas_login.php%3Fpage%3D%2Fcomal%2Fosp%2Fpages%2Fadmin_home.php"

# Prompt for credentials
$USERNAME = Read-Host "Username"
$PASSWORD = Read-Host "Password" -AsSecureString
$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($PASSWORD)
$PLAIN_PASSWORD = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

# Grab login page tokens + CAS cookie
$session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
$loginPage = Invoke-WebRequest -Uri $LOGIN_URL -SessionVariable session -UseBasicParsing

$LT = ($loginPage.InputFields | Where-Object { $_.name -eq "lt" }).value
$EXEC = ($loginPage.InputFields | Where-Object { $_.name -eq "execution" }).value

if (-not $LT -or -not $EXEC) {
    Write-Error "Failed to grab login tokens"
    exit 1
}

# POST login
$body = @{
    username = $USERNAME
    password = $PLAIN_PASSWORD
    lt       = $LT
    execution = $EXEC
    _eventId = "submit"
}

$loginResponse = Invoke-WebRequest -Uri $LOGIN_URL -Method POST -Body $body -WebSession $session -MaximumRedirection 0 -ErrorAction SilentlyContinue -UseBasicParsing

# Follow redirect with ticket to get PHPSESSID
$ticketUrl = $loginResponse.Headers.Location
if (-not $ticketUrl) {
    Write-Error "Login failed - no redirect received"
    exit 1
}

$appSession = New-Object Microsoft.PowerShell.Commands.WebRequestSession
# follow the full redirect chain (ticket validation -> cas_login.php -> admin_home.php)
# so PHP finishes session regeneration and authorizes the OSP area
$appResponse = Invoke-WebRequest -Uri $ticketUrl -SessionVariable appSession -UseBasicParsing

$PHPSESSID = ($appSession.Cookies.GetCookies("https://dawson2.utsarr.net") | Where-Object { $_.Name -eq "PHPSESSID" }).Value

if ($PHPSESSID) {
    Write-Host "PHPSESSID: $PHPSESSID"
} else {
    Write-Error "Failed to get PHPSESSID"
}
