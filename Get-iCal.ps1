<#
	.SYNOPSIS
        A very simple function to read an iCal file and formats the output so that it can be read by Add-CalendarMeeting

	.DESCRIPTION

	.PARAMETER Path
        Path to file to read

	.PARAMETER URI
        URL to iCal data to read

	.PARAMETER ExtraWebRequestParams
        A hashtable with additional parameters that are added to Invoke-WebRequest.

    .PARAMETER Encoding
        Sets the encoding when reading from file. Defaults to UTF8

    .PARAMETER Reminder
        Overrides any reminder in iCal data.

    .PARAMETER AddStartTime
        Adds specified time to MeetingStart. Use a negative value to subtract.

    .PARAMETER AddEndTime
        Adds specified time to MeetingEnd. Use a negative value to subtract.

    .PARAMETER AppendEventDataToBody
        Appends the iCal Event data to meeting body

    .PARAMETER Sensitivity
        Adds Sensitivity property to the output object.

	.EXAMPLE
        Get-iCal -Path d:\calendar\mycalendar.ics | Add-CalendarMeeting -whatif

        Reads ical data from d:\calendar\mycalendar.ics and adds it to Outlook using Add-CalendarMeeting

	.EXAMPLE
        Get-iCal -Path http://ww.domain.com/mycalendar/calendar.ics | Add-CalendarMeeting -whatif

        Reads ical data from http://ww.domain.com/mycalendar/calendar.ics and adds it to Outlook using Add-CalendarMeeting

	.EXAMPLE
        Get-iCal -Path http://ww.domain.com/mycalendar/calendar.ics -ExtraWebRequestParams @{UserAgent="Mozilla/5.0 (iPad; CPU OS 6_0 like Mac OS X) AppleWebKit/536.26 (KHTML, like Gecko) Version/6.0 Mobile/10A5376e Safari/8536.25"} | Add-CalendarMeeting -whatif

        Reads ical data from http://ww.domain.com/mycalendar/calendar.ics and adds it to Outlook using Add-CalendarMeeting
        The request to http://ww.domain.com/mycalendar/calendar.ics is made using an alternate UserAgent header

	.NOTES
        Currently more of a quick hack implemented from looking at a couple of ics files.
        Should readlly be implemented by following RFC 5545. https://tools.ietf.org/html/rfc5545

	.LINK
        https://github.com/Fortika
#>
#Requires –Version 3
Function Get-iCal {
    [cmdletBinding(DefaultParameterSetName="Path")]
    Param(
		[Parameter(Mandatory=$True
                  ,ParameterSetName="Path")]
		[string]$Path

		#,[Parameter(Mandatory=$True
        #          ,ParameterSetName="StringData"
        #          ,ValueFromPipeline=$True)]
		#[string[]]$Data

		,[Parameter(Mandatory=$True
                   ,ParameterSetName="URI")]
		[string]$URI

		,[Parameter(Mandatory=$False
                   ,ParameterSetname="URI"
                   ,HelpMessage="Any additional parameters to invoke-webrequest.")]
		[hashtable]$ExtraWebRequestParams

		,[Parameter(Mandatory=$False
                   ,ParameterSetName="Path")]
		[Microsoft.PowerShell.Commands.FileSystemCmdletProviderEncoding]$Encoding=[Microsoft.PowerShell.Commands.FileSystemCmdletProviderEncoding]::UTF8

		,[Parameter(Mandatory=$False
                   ,HelpMessage="Overrides any reminder in iCal data")]
		[int]$Reminder

		,[Parameter(Mandatory=$False
                   ,HelpMessage="Pattern to use for the location info")]
		[string]$LocationPattern="%LOCATION%"

		,[Parameter(Mandatory=$False
                   ,HelpMessage="Pattern to use for the subject info")]
		[string]$SubjectPattern="%SUBJECT%"

		,[Parameter(Mandatory=$False
                   ,HelpMessage="Pattern to use for the body info")]
		[string]$BodyPattern="%BODY%"

		,[Parameter(Mandatory=$False
                   ,HelpMessage="Time format to use for expanding variables based on any datetime types in the pattern parameters.")]
		[string]$TimeFormat="yyyy-MM-dd HH:mm:ss"

		,[Parameter(Mandatory=$False
                   ,HelpMessage="Adds specified time to MeetingStart. Use a negative value to subtract.")]
		[System.TimeSpan]$AddStartTime=[System.TimeSpan]::Zero

		,[Parameter(Mandatory=$False
                   ,HelpMessage="Adds specified time to MeetingEnd. Use a negative value to subtract.")]
		[System.TimeSpan]$AddEndTime=[System.TimeSpan]::Zero

		,[Parameter(Mandatory=$False
                   ,HelpMessage="Appends the iCal Event data to meeting body")]
		[switch]$AppendEventDataToBody

	    ,[Parameter(Mandatory=$False
                  ,ValueFromPipelineByPropertyName = $True
                  ,HelpMessage="Sets the sensitivity of the calendar item")]
        [ValidateSet('Normal','Personal','Private','Confidential')]
        [string]$Sensitivity = 'Normal' # https://msdn.microsoft.com/VBA/Outlook-VBA/articles/olsensitivity-enumeration-outlook        
    )

    # # Generated with New-FortikaPSFunction -name Get-iCal -Params @{Path=@{Type="string"; Parameter=@{Mandatory=$True; ParameterSetName="Path"}; }; URI=@{Type="string"; Parameter=@{Mandatory=$True; ParameterSetName="URI"}}   }

    BEGIN {
		# If -debug is set, change $DebugPreference so that output is a little less annoying.
		#	http://learn-powershell.net/2014/06/01/prevent-write-debug-from-bugging-you/
		If ($PSBoundParameters['Debug']) {
			$DebugPreference = 'Continue'
		}

        if($Path) {

            try {
            
                $Data = Get-Content -Path $Path -ErrorAction Stop -Encoding $Encoding

            }
            catch {
                Throw "{0}" -f $_.Exception.Message
            }
        } elseif($URI) {
        
            try {
                if(-not $ExtraWebRequestParams) { $ExtraWebRequestParams = @{} }
                $WebContent = (Invoke-WebRequest -Uri $URI -UseBasicParsing @ExtraWebRequestParams -ErrorAction Stop).Content

                # content from web is sent as a string and not an array of strings which we currently use in parsing loop
                $Data = $WebContent -split "`r`n"
            }
            catch {
                Throw "{0}" -f $_.Exception.Message
            }

        }

        #if( -not ($Data[0] -eq "BEGIN:VCALENDAR" -and $Data[1] -eq "VERSION:2.0")  ) {
        #    Throw "iCal data is not on supported format!"
        #}
        if( -not ($Data[0] -eq "BEGIN:VCALENDAR" )  ) {
            Throw "iCal data is not on supported format!"
        }

        if($AppendEventDataToBody) {
            $BodyPattern += "%EVENTDATA%"
        }

        # Build a hashtable to easier be able to apply the variable replacements
        # Format:
        #   <property in $CalObject> = <pattern string>
        $PropertyPatternTable = @{
            Body=$BodyPattern;
            Location=$LocationPattern;
            Subject=$SubjectPattern;
        }
    }

    PROCESS {

        # we're doing a little state machine for the fun of it...

        $index=2
        $State="FIND_VEVENT"

        do {

            $str = $Data[$index]
            
            Write-Debug "iCal string: $str"
        
            switch($State) {
                "FIND_VEVENT" {
                    
                    if(-not ($str -eq "BEGIN:VEVENT") ) {
                        $index++
                        continue

                    } else {
                        $CalObjectStrings = @()
                        $CalObject = New-Object -TypeName PSObject

                        $State="READ_VEVENT"
                        $index++
                        continue

                    }

                }
                "READ_VEVENT" {

                    # We are reading a VEVENT

                    
                    if($str -eq "END:VEVENT") {

                        # validate object


                        #if(-not $CalObject.Subject) {
                        #
                        #}

                        # handle overrides
                        if($Reminder) {                        
                            if ( [bool]($CalObject.psobject.Properties.name -match "Reminder") ) {
                                $CalObject.Reminder = $Reminder
                            } else {
                                $CalObject | Add-Member -MemberType NoteProperty -Name "Reminder" -Value $Reminder
                            }                        
                        }

                        if($Sensitivity) {
                            if ( [bool]($CalObject.psobject.Properties.name -match "Sensitivity") ) {
                                $CalObject.Sensitivity = $Sensitivity
                            } else {
                                $CalObject | Add-Member -MemberType NoteProperty -Name "Sensitivity" -Value $Sensitivity
                            }                        

                        }

                        #if($AppendEventDataToBody) {
                        #    if ( [bool]($CalObject.psobject.Properties.name -match "Body") ) {
                        #        $CalObject.Body += $CalObjectStrings -join "`r`n"
                        #    } else {
                        #        $CalObject | Add-Member -MemberType NoteProperty -Name "Body" -Value ($CalObjectStrings -join "`r`n")
                        #    }                        
                        #}

                        


                        $PropertyPatternTable.Keys | ForEach-Object {
                            # Property name is in the hashtable key
                            $PropertyName = $_
                            $PropertyPattern = $PropertyPatternTable.Item($PropertyName)


                            # ok, so decided not to care if there's any properties missing from the ical data.
                            # For example, if Subject is missing, then it would be empty anyway, so just create it if
                            # it does not exist, and apply pattern.
                            if ( -not [bool]($CalObject.psobject.Properties.name -match $PropertyName) ) {
                                $CalObject | Add-Member -MemberType NoteProperty -Name $PropertyName -Value ""
                            }

                            $CalObject.$PropertyName = $PropertyPattern | _Expand-VariablesInString -VariableMappings @{
                                                                                BODY=$CalObject.Body;
                                                                                SUBJECT=$CalObject.Subject;
                                                                                LOCATION=$CalObject.Location;
                                                                                EVENTDATA=($CalObjectStrings -join "`r`n");
                                                                                MEETINGSTART=$CalObject.MeetingStart.ToString($TimeFormat);
                                                                                MEETINGEND=$CalObject.MeetingStart.ToString($TimeFormat);
                                                                                REMINDER=$CalObject.Reminder;
                                                                        }
                        }


                        # output object
                        $CalObject

                        $index++
                        $State="FIND_VEVENT"

                        continue
                    }

                    $CalObjectStrings += $str

                    #($iCalItem,$ItemData) = ($str -replace "([a-z]+):(.*)",'$1§$2') -split "§"
                    ($iCalItem,$ItemData) = $str -split "[:;]",2

                    # we're not handling any parameters attached to the property
                    
                    Write-Debug "Parsed item: item = $iCalItem | data = $ItemData"

                    switch ($iCalItem) {
                        "DTSTART" {
                            # https://stackoverflow.com/questions/26997511/how-can-you-test-if-an-object-has-a-specific-property
                            if ( [bool]($CalObject.psobject.Properties.name -match "MeetingStart") ) {
                                Write-Warning "There's already a property named MeetingStart in object!"
                            } else {
                                # this is a bit ugly...
                                # $ItemData might be TZID=Europe/Stockholm:20171114T100000
                                # split the $ItemData and try and parse each item as a date to find the date...
                                foreach($item in $ItemData.split(':')) {
                                    
                                    try {
                                        # somehow cant get dates on format 20171031T141500Z to parse correctly...
                                        $d = [datetime]::Parse( ($Item -replace "(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(.*)",'$1-$2-$3 $4:$5:$6') )
                                        $d = $d + $AddStartTime
                                        break
                                    } catch {
                                    }
                                }

                                $CalObject | Add-Member -MemberType NoteProperty -Name "MeetingStart" -Value $d
                            }
                        }
                        "DTEND" {
                            if ( [bool]($CalObject.psobject.Properties.name -match "MeetingEnd") ) {
                                Write-Warning "There's already a property named MeetingEnd in object!"
                            } else {

                                foreach($item in $ItemData.split(':')) {
                                    
                                    try {
                                        # somehow cant get dates on format 20171031T141500Z to parse correctly...
                                        $d = [datetime]::Parse( ($Item -replace "(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(.*)",'$1-$2-$3 $4:$5:$6') )
                                        $d = $d + $AddEndTime
                                        break
                                    } catch {
                                    }
                                }
                
                                $CalObject | Add-Member -MemberType NoteProperty -Name "MeetingEnd" -Value $d
                            }
                        
                        }
                        "LOCATION" {
                            if ( [bool]($CalObject.psobject.Properties.name -match "Location") ) {
                                Write-Warning "There's already a property named Location in object!"
                            } else {

                                # LOCATION can also be multi-line
                                # See DESCRIPTION

                                $tmpIndex = $Index
                                
                                while( $Data[++$tmpIndex] -match "^ .*" ) {
                                    # skip the first character in string
                                    $ItemData = $ItemData + $Data[$tmpIndex].SubString(1)                                    
                                }
                                
                                $CalObject | Add-Member -MemberType NoteProperty -Name "Location" -Value ${ItemData}${AppendLocation}
                            }
                        }
                        "X-GWSHOW-AS" {
                            $CalProperty = "BusyStatus" # this should be in a mapping table...
                            if ( [bool]($CalObject.psobject.Properties.name -match $CalProperty ) ) {
                                Write-Warning ("There's already a property named {0} in object!" -f $CalProperty)
                            } else {
                                $ValidBusyStatus=@('Free','Tentative','Busy','OutOfOffice')
                                if($ItemData -in $ValidBusyStatus) {
                                    $CalObject | Add-Member -MemberType NoteProperty -Name $CalProperty -Value $ItemData
                                } else {
                                    Write-Warning "Unsupported BusyStatus $ItemData"
                                }
                            
                            }
                        }
                        "SUMMARY" {
                            $CalProperty = "Subject" # this should be in a mapping table...
                            if ( [bool]($CalObject.psobject.Properties.name -match $CalProperty ) ) {
                                Write-Warning ("There's already a property named {0} in object!" -f $CalProperty)
                            } else {
                                $CalObject | Add-Member -MemberType NoteProperty -Name $CalProperty -Value $ItemData                            
                            }
                        
                        }
                        "DESCRIPTION" {
                            $CalProperty = "Body" # this should be in a mapping table...
                            if ( [bool]($CalObject.psobject.Properties.name -match $CalProperty ) ) {
                                Write-Warning ("There's already a property named {0} in object!" -f $CalProperty)
                            } else {

                                # look further into $data to see if description is multi-line-
                                # multi-line descriptions seems to start next line with a space.
                                # DESCRIPTION:This is\n
                                #  a multi-line\n
                                #  description\n
                                # \n seems to be line terminators.
                                # There also seems to be wrapped lines that does not have line terminators
                                # DESCRIPTION:This is\n
                                #  a wrapped multi-line
                                #  description\n

                                $tmpIndex = $Index

                                while( $Data[++$tmpIndex] -match "^ .*" ) {
                                    # skip the first character in string
                                    $ItemData = $ItemData + $Data[$tmpIndex].SubString(1)                                    
                                }

                                $CalObject | Add-Member -MemberType NoteProperty -Name $CalProperty -Value $ItemData.replace('\n','`r`n')
                            }
                        
                        }
                        default {
                            Write-Debug "Unhandled item $iCalItem | Data = $ItemData"
                        }
                    }

                    $index++
                }
                default {
                    # some sanity checking 
                    Throw "Illegal state {0}" -f $State
                }
            
            }



        
        } while($index -lt $Data.count)
        
    }

    END {

    }
}

Function _Expand-VariablesInString {
    [cmdletBinding()]
    Param(
        [Parameter(Mandatory=$True
                  ,ValueFromPipeline=$True)]
        [string]$Inputstring,

        [Parameter(Mandatory=$True)]
        [hashtable]$VariableMappings
    )


    foreach($key in $Variablemappings.Keys) {

        $InputString = $Inputstring.Replace("%"+$key+"%",$VariableMappings[$key])
    }


    return $Inputstring
}



