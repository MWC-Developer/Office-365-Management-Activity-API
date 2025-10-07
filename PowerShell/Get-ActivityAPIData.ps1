 
# Get-ActivityAPIData.ps1
#
# By David Barrett, Microsoft Ltd. Use at your own risk.  No warranties are given.
#
#  DISCLAIMER:
# THIS CODE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
# THE SOFTWARE.

# This script is called as a scheduled task (via a batch file), so the app Id and secret key are stored as a file encrypted to the user account.  Ensure the scheduled task is run under the same user context.

$tenantId = "77275f64-d6b0-4d6d-b04b-8580417d20a6"
$startDate = [DateTime]::Today.AddDays(-6) # The oldest date we can request
$rootPath = "E:\Scripts\demonmaths.co.uk"
$appSecretKeyFile = "$rootPath\AppSecret.xml"
$lastRunFile = "$rootPath\Get-ActivityAPIData.lastrun"

# Save the credential file as follows (or run the script from PS console to be prompted to create the file)
# $appId = "392fbfe0-b277-4b91-b82c-cb45d1a41c49"
# $credential = Get-Credential -UserName $appId -Message "Enter secret key for app Id $appId"
# $credential |  Export-Clixml -Path $appSecretKeyFile

# Retrieve AppId and secret key
$appAuth = Import-Clixml -Path $appSecretKeyFile
if (!$appAuth) {
	Write-Host "Failed to read authentication information from $appSecretKeyFile" -ForegroundColor Red
	if ([System.Environment]::UserInteractive) {
		# If interactive, prompt to create and save the app credential file
		$credential = Get-Credential -Message "Enter App Id (as username) and secret key"
		if ($credential) {
			$credential |  Export-Clixml -Path $appSecretKeyFile
			$appAuth = Import-Clixml -Path $appSecretKeyFile
		}
	}
	if (!$appAuth) {
		exit
	}
}

try
{
    $lastrun = (Get-Content $lastRunFile | ConvertFrom-Json).value
    if ($lastrun -gt $startDate)
    {
        $startDate = $lastrun
    }
} catch {}

$secret = [Net.NetworkCredential]::new('', $appAuth.Password)

# We want to retrieve the data for any days between the last run and now (up to seven)
while ($startDate -le [DateTime]::Today.AddDays(-1))
{
    .\Test-ManagementActivityAPI.ps1 -RetrieveContent -SaveContentPath "$rootPath\Data" -ListContentDate $startDate -MaxRetrieveContentJobs 25 -AppId $appAuth.Username -TenantId $tenantId -AppSecretKey $secret.Password -LogFile "$rootPath\Get-ActivityAPIData.log" -verbose
    $startDate = $startDate.AddDays(1)
}

$startDate | ConvertTo-Json | out-file $lastRunFile