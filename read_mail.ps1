##-----------------------------------------------------## 
##        PICK AUTH Method                             ## 
##-----------------------------------------------------## 
 
## HARD CODING PSW    ## 
#$password = ConvertTo-SecureString "xxx" -AsPlainText -Force 
#$cred = New-Object System.Management.Automation.PSCredential "xxx@xxx.onmicrosofot.com",$password 
 
## USER PROMPT PSW    ## 
$cred = Get-Credential

$filestamp=Get-Date -Format "yyyy-dd-MM"
$shared_path_file = "c:\temp\test\BG_measurement_$filestamp.csv"
$start_time=$(Get-Date -UFormat "%Y-%m-%d %H:%M:%S")

function script_measure($server, $start_time, $end_time, $status) {
$usecasename="IIS Automation"
$execution_type="Assisted"
$GUID=[guid]::NewGuid().Guid

"$server,$GUID,$start_time,$end_time,$usecasename,$execution_type,$status"|Out-File -Append $shared_path_file

}

##-----------------------------------------------------## 
##    END PICK 
##-----------------------------------------------------## 
 
$url = "https://outlook.office365.com/api/v1.0/me/messages" 
$fromdate = "2018-11-10T00:00:00Z" 
$todate = "2018-11-19T23:59:59Z"

## Get all messages that have attachments where received date is greater than $date  
$messageQuery = "" + $url + "?`$select=Id,subject&`$filter=HasAttachments eq true and IsRead eq false and DateTimeReceived ge " + $fromdate + " and DateTimeReceived le " + $todate + "&`$top=200"
#`$search=subject:SNOW&
#$messageQuery = "" + $url + "?`$select=Id&`$search=%22Received:today%22 AND %22subject:efficiency, effectiveness and experience%22"
echo "$messageQuery"
$messages = Invoke-RestMethod $messageQuery -Credential $cred 
#echo $messages 
## Loop through each results 
foreach ($message in $messages.value) 
{ 
    #$message.subject
    if ( $message.subject -like "*Meter*") 
    {
        # get attachments and save to file system 
        $query = $url + "/" + $message.Id + "/attachments" 
        $attachments = Invoke-RestMethod $query -Credential $cred 
 
        # in case of multiple attachments in email 
        foreach ($attachment in $attachments.value) 
        { 
            #$attachment.Name
            if ( $attachment.Name -like "*.csv" )
            {
            $version=Get-Random -Maximum 100
            $path = "c:\Temp\test\" + $version + "_" + $attachment.Name 
     
            $Content = [System.Convert]::FromBase64String($attachment.ContentBytes) 
            Set-Content -Path $path -Value $Content -Encoding Byte
            #script_measure $attachment.Name "$start_time" "$(Get-Date -UFormat "%Y-%m-%d %H:%M:%S")" "Completed"
            }
        }
       #$message.IsRead = $false
    } 
}
Get-ChildItem c:\Temp\test\*.csv |ForEach-Object { Import-Csv $_ } |Export-Csv c:\Temp\test\output\newOutputFile.csv -NoTypeInformation
if ($?){
    Remove-Item "c:\Temp\test\*.csv"
}
Import-Csv C:\temp\test\output\newOutputFile.csv | Sort-Object -Property  JOBID,Start_Time -Unique |Export-Csv "C:\temp\test\result\test_data.csv" -NoClobber -NoTypeInformation -Force
if ($?){
    Remove-Item "C:\temp\test\output\newOutputFile.csv"
}
