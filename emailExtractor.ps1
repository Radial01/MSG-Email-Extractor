# Modify the email pattern to match the email address for which you are searching.
$emailPattern = "([a-zA-Z0-9_\-\.]+)@gmail.com"
$msgDirectory = "C:\msgFiles"

# Parse .msg files, extracting email addresses matching the regex specified above.
# Matches are placed in temp.txt.
$Outlook = New-Object -ComObject Outlook.Application
Get-ChildItem $msgDirectory -Filter *.msg |
    ForEach-Object {
        $msg = $outlook.Session.OpenSharedItem($_.FullName)
        $msg.body
        $results = ($msg.body | Select-String $pattern -AllMatches).Matches
        foreach ($item in ($results)) { 
            Add-Content -Path .\temp.txt $item.Value
        }
    }

# Remove duplicate Entries from temp.txt and then remove temp.txt itself.
$hash = @{}
Get-Content -Path .\temp.txt | %{if($hash.$_ -eq $null) { $_ }; $hash.$_ = 1} > "List.txt"
Remove-Item -Path .\temp.txt -Force
