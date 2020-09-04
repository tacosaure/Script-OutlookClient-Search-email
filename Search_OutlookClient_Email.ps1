param(
    $skip=$false,
    $keyword=$false, 
    $mailbox_folder="", 
    $email_address="",
    [switch]$output_full_option=$false,
    $output_path=(Get-Location).tostring(),
    $output_filename=$null,
    [switch]$export_emails_html=$false,
    [switch]$deep,
    $startDate,
    $endDate=(Get-Date -format "dd/MM/yyyy HH:mm:ss")
)

if($startDate)
{$startDate = get-date $startDate -Format "MM/dd/yyyy HH:mm:ss"}
$endDate = get-date $endDate -Format "MM/dd/yyyy HH:mm:ss"

#to do: input validation => when startdate is greater than enddate

$output_path=$output_path.tostring() -replace "/","\\"

$output_name_directory=(Get-Date -Format "yyyyMMddTHHmmssffffZ") + "_email_search_results"
if (!($output_path.tostring()[$output_path.tostring().length -1] -eq "\"))
    {$output_path =$output_path.tostring() + "\"}

$output_filepath=$output_path +$output_name_directory +"\"+ $output_filename
#parameters
####### Searching parameters

if($keyword)
{
    $keyword=[regex]::escape($keyword)
    if($deep)
    {
        $body_search_deep_keyword = $keyword
    }
}


####### script variables
$count_total_mails=0
$count_total_searched_mails = 0
[system.array]$output_searched_mail=$null

function Get-MailboxFolder($folder)
{   
    if($keyword)
    {
        if($deep)
        {
            $searched_mail = $folder.items |where {(($_.subject -match $keyword) -or ($_.HTMLbody -match $body_search_deep_keyword)) -and ($folder.FullFolderPath -notmatch $skip) -and (($_.receivedtime -ge $startDate) -and ( $_.receivedtime -le $endDate))}
        }
        else
        {
            $searched_mail = $folder.items |where {($_.subject -match $keyword) -and ($folder.FullFolderPath -notmatch $skip) -and (($_.receivedtime -ge $startDate) -and ( $_.receivedtime -le $endDate))}
        }
    }
    else
    {
        $searched_mail = $null
    }
    
    [String[]]$searched_mail_array = $searched_mail
    
    #print
    if($keyword)
    {
        write-host ("`t"*$call + "{0}: {1}/{2}" -f $folder.name,$searched_mail_array.count, $folder.items.count)        
    }
    else
    {
        write-host("`t"*$call + "{0}: {1}" -f $folder.name, $folder.items.count)
    }
    
    $call++
    
    $searched_mail |sort -descending receivedtime | format-table subject,receivedtime,sendername,to,senton | Out-String |%{write-host $_ -NoNewline}
    
    foreach ($f in $folder.folders | sort -Property name |where {$_.FullFolderPath -notmatch $skip})
    {
        
        Get-MailboxFolder $f
        
    }


    $script:output_searched_mail +=($searched_mail|Select-Object @{Name = "FolderLocation"; Expression={$folder.FullFolderPath}},* )

    $script:count_total_searched_mails +=$searched_mail_array.count
    $script:count_total_mails+=$folder.items.count

    $searched_mail =$null
}


$ol = new-object -com Outlook.Application
$ns = $ol.GetNamespace("MAPI")
$mailbox = $ns.Folders($email_address)#.folders($mailbox_folder)


foreach ($object in $mailbox.folders | sort -Property name | where {($_.FullFolderPath -notmatch $skip) -and ($_.name -match $mailbox_folder)})
{
    $call=0 
    Get-MailboxFolder $object
}

write-host "------------------------"
write-host "Total of mails = $count_total_mails"
write-host "Total of searched mails = $count_total_searched_mails"

if(($output_filename) -or ($export_emails_html))
{
        New-Item -ItemType "directory" -Path $output_path -Name ($output_name_directory) | Out-Null
}

if($output_filename)
{

    if($output_full_option)
    {
        $output_searched_mail| sort -Descending -Property receivedtime | select-object * -ExcludeProperty HTMLBody| Export-Csv -Delimiter ";" -path $output_filepath -NoTypeInformation
        write-host "Output file (summary) is created with all properties: $output_filepath"
    }

    elseif (!$output_full_option)
    {
        $output_searched_mail| sort -Descending -Property receivedtime | select-object folderLocation,receivedtime,sendername,sentonbehalfofname,to,cc,bcc,subject | Export-Csv -Delimiter ";" -path $output_filepath -NoTypeInformation
        write-host "Output file (summary) is created with default properties: $output_filepath"
    }
    else
    {
        write-host "An issue occurred"
    }
}
elseif($output_filename -eq $null)
{
    write-host "No output file (summary) has been created"
}
else
{
    write-host "An issue occurred"
}

if($export_emails_html)
{
    New-Item -ItemType "directory" -Path ($output_path+$output_name_directory) -Name "emails_html" | Out-Null
    foreach ($email in $output_searched_mail ){
        $html_path = $output_path+$output_name_directory + "\emails_html"
        $mail_html_name=$email.entryID +".html"
        $mail_html_body="<b>Received date:</b> "+$email.receivedtime+"<br><b>From</b>:"+$email.senderName+" <b>Sent On Behalf Of:</b> "+$email.sentonbehalfofname+"<br><b>To:</b> "+$email.to+"<br><b>Cc:</b> "+$email.cc+"<br><b>BCC:</b> "+$email.bcc+"<br><b>Subject:</b> "+$email.subject+"<br><br>"+$email.HTMLbody
        New-Item -ItemType "file" -Path $html_path -Name $mail_html_name -Value $mail_html_body | Out-Null
    }
    write-host "Emails extract (HTML) has been created in the folder: " $output_path$output_name_directory"\emails_html\"
}
return $output_searched_mail