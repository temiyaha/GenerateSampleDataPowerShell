function New-SampleWord{
	$csvdata = Import-Csv C:\Windows\System32\WindowsPowerShell\v1.0\Modules\GenerateSampleDataPowerShell\word.csv -Header "data"
	$TXT = $csvdata.data | Get-Random
	return $TXT
}

function New-SampleSentense{
	$num = Get-Random -minimum 10 -maximum 50
	$csvdata = Import-csv C:\Windows\System32\WindowsPowerShell\v1.0\Modules\GenerateSampleDataPowerShell\sentense.csv -Header "data"
	$TXT = ($csvdata.data | Get-Random -count $num) -join ""
	return $TXT       
}

function New-SampleFirstName{
	$csvdata = Import-Csv C:\Windows\System32\WindowsPowerShell\v1.0\Modules\GenerateSampleDataPowerShell\name.csv
	$firstName = $csvdata.First | Get-Random
	return $firstName
}

function New-SampleLastName{
	$csvdata = Import-Csv C:\Windows\System32\WindowsPowerShell\v1.0\Modules\GenerateSampleDataPowerShell\name.csv
	$LastName = $csvdata.Given | Get-Random
	return $LastName
}

function New-SampleDept{
	$csvdata = Import-Csv C:\Windows\System32\WindowsPowerShell\v1.0\Modules\GenerateSampleDataPowerShell\dept.csv -Header "dept"
	$dept = $csvdata.dept | Get-Random
	return $dept
}

function New-SampleAlias{
	$num = Get-Random -minimum 4 -maximum 6
	$vowle = 'aiueo'
	$char = 'aghikomnrstwy'
	for($i=1;$i -lt $num; $i++){
		$random = Get-Random -Maximum $char.length
		$alias += -join $char[$random]
		$random = Get-Random -Maximum $vowle.length
		$alias += -join $vowle[$random]
	}	
	return $alias
}

function New-SampleFileName{
	$csvdata = Import-Csv C:\Windows\System32\WindowsPowerShell\v1.0\Modules\GenerateSampleDataPowerShell\filename.csv -Header "filename"
	$filename = $csvdata.filename | Get-Random
	return $filename
}

function New-SampleWordFile{
    param(
        [parameter(mandatory=$true)] [string]$savePath,
        [parameter(mandatory=$true)] [string]$fileName,
        [parameter(mandatory=$true)] [string]$body
    )
    	$filePath = $savePath + "\" +$fileName
	[ref]$SaveFormat = "microsoft.office.interop.word.WdSaveFormat" -as [type]
        $word = New-Object -ComObject word.application
        $word.visible = $false
        $doc = $word.documents.add()
        $selection = $word.selection |Out-Null
        $selection.WholeStory |Out-Null
        #$selection.Style = "No Spacing"
        $selection.font.size = 10
        # $selection.typeText("My Document: Title")

        $selection.TypeParagraph()
        $selection.font.size = 11
        $selection.typeText("$body")
        $doc.saveas([ref] $filePath, [ref]$saveFormat::wdFormatDocument)
        $doc.close()
        $word.quit()
}

function New-SampleWordFiles{
    param(
	[parameter(mandatory=$true)] [string]$savePath,
        [parameter(mandatory=$true)] [string]$num
    )
	for($i=0;$i -lt $num; $i++){
		New-SampleWordFile -savePath $savePath -fileName (New-SampleFileName) -body (New-SampleSentense)
	}
}

function New-SampleADUsers{
    param(
	[parameter(mandatory=$true)] [string]$OUName,
        [parameter(mandatory=$true)] [string]$num
    )
	for($i=0;$i -lt $num; $i++){
		$FirstName = New-SampleFirstName
		$LastName = New-SampleLastName
		$DisplayName = $FirstName + " " + $LastName
		$alias = New-SampleAlias
		$upn = $alias + "@contoso.local"
		$password = ConvertTo-SecureString "P@ssw0rd1!" -AsPlainText -Force
		New-ADUser -Name $alias -Surname $LastName -GivenName $FirstName -DisplayName $DisplayName -UserPrincipalName $upn -AccountPassword $password -PasswordNeverExpires $true -path "OU=$OUName,OU=User,DC=contoso,DC=local" -Enabled $True
	}
}

Export-ModuleMember -Function *

