<#
.Synopsis
   Parses html files produced by Google Takeout to present Google Voice call and text history in a useful way.
.DESCRIPTION
   When exporting Google Voice data using the Google Takeout service, the data is delivered in the form
   of many individual .html files, one for each call (placed, received, or missed), each text message 
   conversation, and each voicemail or recorded call. For heavy users of Google Voice, this could mean many
   thousands of individual files, all located in a single directory.  This script parses all of the html 
   files to collect details of each call or message and outputs them as an object which can then be 
   manipulated further within powershell or exported to a file.  
   
   This script requires the "HTML Agility Pack", available via NuGet.  The script was written and tested
   with HTML Agility Pack version 1.4.9 For more information, go to http://htmlagilitypack.codeplex.com/
   
   In order to run this script, you must have at least powershell version 3.0.  If you're running Windows 7,
   You can update powershell by installing the latest "Windows Management Framework" from microsoft, currently
   WMF 4.  
   
   About This Script
   -----------------
   Author: Steve Whitcher
   Web Site: http://www.neighborgeek.net
   Version: 1.01
   Date: 12/21/2014
   
.EXAMPLE
   import-gvhistory -path c:\temp\takeout\voice\calls -agilitypath C:\packages\HtmlAgilityPack.1.4.9\lib\net45
   
   This command parses files in c:\temp\takeout\voice\calls using the HtmlAgilityPack.dll file located 
   in 'C:\packages\HtmlAgilityPack.1.4.9\lib\net45'.  Run this way, all of the text message and call history would be
   output to the screen only.
   
.EXAMPLE
   import-gvhistory -path c:\temp\takeout\voice\calls -agilitypath C:\packages\HtmlAgilityPack.1.4.9\lib\net45\ |
        where-object {$_.Type -eq "Text"} | export-csv c:\temp\TextMessages.csv
        
    This command uses the same parameters as Example 1, but then passes the information on be filtered 
    by Where-Object to only include records of Text messages, and not calls.  After filtering, the information
    is saved to c:\temp\TextMessages.csv by passing the output of Where-Object to Export-CSV.  
.EXAMPLE
   import-gvhistory -path c:\temp\takeout\voice\calls | export-csv c:\temp\GVHistory.csv
    This command does not include the -agilitypath parameter, so the script will attempt to find 
    and use HTMLAgilityPack.dll in the current working directory. The command will process all call and text
    message information and save it to c:\temp\GVHistory.csv
    
#>
function import-gvhistory
{
    [CmdletBinding()]
    [Alias()]
    [OutputType("Selected.System.Management.Automation.PSCustomObject")]
    #Requires -Version 3.0
    Param
    (
        # Path to the "Calls" directory containing Google Voice data exported from Google Takeout.
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $Path,

        # Path to "HtmlAgilityPack.dll" if not located in the working directory.
        $AgilityPath = "." 
    )

    Begin
    {
        $option = [System.StringSplitOptions]::None
        $separator = "-"
        $Records = (get-childitem $Path) | Where-object {$_.Name -like "*.html"}
        $Calls = @()
        $Texts = @()
        $GVHistory = @()
        add-type -assemblyname system.web
        add-type -path "$($AgilityPath)\HtmlAgilityPack.dll"
        
    }
    Process
    {
        ForEach ($Record in $Records) 
        {
            Write-Verbose "Record $Record.Name" # File name being processed
            
            # Split File Name into Contact Name, Call Type, and Timestamp
            $RecordName = (($Record.Name).trimend(".html")).split($separator,3,$option)
            Write-Verbose "RecordName $RecordName"
            $Contact = $RecordName[0].trim()
            $Type = $RecordName[1].trim()
            $FileTime = $RecordName[2]
            Write-Verbose "Name $Contact"
            Write-Verbose "Type $Type"
            Write-Verbose "TimeStamp $FileTime"
            Write-Verbose ""

            $doc = New-Object HtmlAgilityPack.HtmlDocument
            $source = $doc.Load($Record.fullname)

            if ($Type -ne "Text")
            {
                # Record is of a phone call that was placed, received, or missed, or of a voicemail message.

                $GMTTime = $doc.documentnode.selectnodes("//abbr [@class='published']").InnerText.Trim()
                $CallTime = get-date $GMTTime
                $Tel = $doc.documentnode.selectnodes(".//a [@class='tel']")
                $ContactName = $tel.selectsinglenode(".//span[1]").InnerText.Trim()
                $ContactNum = $tel.GetAttributeValue("href", "Number").TrimStart("tel:+")

                If ($Type -ne "Missed" -and $Type -ne "Recorded") 
                {
                    # Missed Calls don't have a duration listed.  Some recorded calls might also be zero length.
                    # Get duration for all other call types.
                    $Duration = $doc.documentnode.selectnodes(".//abbr[@class='duration']").InnerText.Trim("(",")")
                }
                
                Else 
                {
                    $Duration = ""
                }
                If ($Type -eq "Voicemail") 
                {
                    # Get the Automated Transcription of voicemail messages as well as the name of the mp3 audio file.                   
                    $FullText = $doc.documentnode.selectnodes("//span [@class='full-text']").InnerText
                    $Fulltext = [System.Web.HttpUtility]::HtmlDecode($FullText)
                    $Audio = $doc.documentnode.selectsinglenode("//audio")
                    If ($Audio) 
                    {
                        # If there was no audio recorded, the mp3 file won't exist.  
				        $AudioFilePath = $Audio.GetAttributeValue("src", "")
			        }
                }

                Else 
                {
                    # Calls of type other than "Voicemail" won't have audio or transcription, so blank the variables.
                    $FullText = ""
                    $Audio = ""
                    $AudioFilePath = ""
                }

                # Add the details of this call record to $Calls

                $Calls += [PSCustomObject]@{
                    Contact = $ContactName
                    Time = $CallTime
                    Type = $Type
                    Number = $ContactNum
                    Duration = $Duration
                    Message = $FullText
                    AudioFile = $AudioFilePath
                    Direction = ""
                }
            }

            else 
            {
                # Record is of an SMS Conversation containing one or more messages
                
                $Messages = $doc.documentnode.selectnodes("//div[@class='message']")
                
                # Each HTML file represents a single SMS "Conversation".  A conversation could include many messages.
                # Process each individual message. 

                ForEach ($Msg in $messages) {
                    $GMTTime = $msg.selectsinglenode(".//abbr[@class='dt']").InnerText.Trim()
                    $MsgTime = get-date $GMTTime
                    $Tel = $msg.selectsinglenode(".//a [@class='tel']")
                    $SenderName = $tel.InnerText.Trim()
                    $SenderNum = $tel.GetAttributeValue("href", "Number").TrimStart("tel:+")
                    $Body = $msg.selectsinglenode(".//q").InnerText.Trim()
                    $Body = [System.Web.HttpUtility]::HtmlDecode($Body)
                    if ($SenderName -eq "Me")
                    {
                        $Direction = "Received"
                    }
                    else
                    { 
                        $Direction = "Sent"
                    }

                    # Add the details of this message to $Texts

                    $Texts += [PSCustomObject]@{
                        Contact = $Contact
                        Time = $MsgTime
                        Type = $Type
                        Direction = $Direction
                        Message = $Body
                    }
                }

            }

        }
    }
    End
    {
        # Combine all $Calls and $Texts, sort based on the timestamp.  
        $GVHistory = $Calls + $Texts
        $GVHistory | Sort Time
    }
}
