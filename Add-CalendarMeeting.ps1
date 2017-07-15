<#
	.SYNOPSIS
        Powershell function to create Outlook Calendar Meeting on the fly.

	.DESCRIPTION
        This function is originally from http://newdelhipowershellusergroup.blogspot.se/2013/08/powershell-and-outlook-create-calendar.html
        which was written by Aman Dhally, www.amandhally.net

        Code slightly modified and extended and help text is re-formatted.

	.PARAMETER Subject
		The subject of the meeting.

	.PARAMETER  Body
		The text that goes into the Body of the meeting

	.PARAMETER Location
		The location of your Meeting, for example can be meeting room1 or any country.

	.PARAMETER Importance
        The importance of the meeting.
        Default set to Normal.

	.PARAMETER AllDayEvent
		Indicates that the event is covering all day

    .PARAMETER BusyStatus
        Sets the Busy status.

	.PARAMETER DisableReminder
		Disables any reminder for the meeting.
        By default reminders are on.

	.PARAMETER MeetingStart
		Date and time when meeting starts.

    .PARAMETER MeetingEnd
        Date and time when meeting ends.
        Can not be specified together with MeetingDuration.

	.PARAMETER  MeetingDuration
		The duration of the meeting in minutes.

    .PARAMETER EventType
        Used to indicate if the meeting is all day or a "normal" meeting when sending parameters over the pipeline.
        If not piping in parameters, use the switch AllDayEvent

	.PARAMETER  Reminder
        The number of minutes befre meeting start that triggers the reminder.

	.EXAMPLE
		Add-CalendarMeeting -Subject "Powershell" -Body "Show how to use Powershell with Outlook" -Location "Conf Room 1" -AllDayEvent -DisableReminder
		
	.EXAMPLE
		Add-CalendarMeeting -Subject "Powershell" -Body "Show how to use Powershell with Outlook" -Location "Conf Room 1" -MeetingStart "08/08/2013 22:30" -Reminder 30 
	
	.EXAMPLE
		Add-CalendarMeeting -Subject "Powershell" -Body "Show how to use Powershell with Outlook" -Location "Conf Room 1" -Importance 'High'

    .EXAMPLE
        "MeetingStart;MeetingEnd;Subject`r`n2017-05-01 15:00;2017-05-01 17:00;Test`r`n" | ConvertFrom-Csv -Delimiter ";" | Add-CalendarMeeting -WhatIf -Verbose -Debug

        Adds a calendar meeting from a CSV string.
        Headers in CSV must match parameter names.

	.NOTES
        This function is from http://newdelhipowershellusergroup.blogspot.se/2013/08/powershell-and-outlook-create-calendar.html
        which was written by Aman Dhally www.amandhally.net

        Code slightly modified and extended and help text is re-formatted.

	.LINK

#>
Function Add-CalendarMeeting {
[cmdletBinding(SupportsShouldProcess=$true,
               ConfirmImpact='Medium')]
Param(
    [Parameter(
        Mandatory = $True,
        ValueFromPipelineByPropertyName = $True,
        HelpMessage="Please provide a subject of your calendar item.")]
    [Alias('sub')]
    [string]$Subject

    ,[Parameter(
        ValueFromPipelineByPropertyName = $True,
        HelpMessage="Please provide a description of your calendar item.")]
    [Alias('bod')]
    [string]$Body

    ,[Parameter(
        ValueFromPipelineByPropertyName = $True,
        HelpMessage="Please provide the location of your meeting.")]
    [Alias('loc')]
    [string]$Location

	,[Parameter(
        Mandatory=$True,
        ValueFromPipelineByPropertyName = $True)]
	[datetime]$MeetingStart

	,[Parameter(
        ValueFromPipelineByPropertyName = $True)]
	[datetime]$MeetingEnd

	,[Parameter(
        ValueFromPipelineByPropertyName = $True)]
	[int]$MeetingDuration = 30

    ,[Parameter(
        ValueFromPipelineByPropertyName = $True,
        HelpMessage="Please provide the importance of your meeting.")]
	[ValidateSet('Normal','Low','High')]
	[string]$Importance = 'Normal'

	,[Parameter(
        ValueFromPipelineByPropertyName = $True)]
	[ValidateSet('Free','Tentative','Busy','OutOfOffice')]
	[string]$BusyStatus = 'Busy'

    
	,[Parameter(Mandatory=$False
              ,ValueFromPipelineByPropertyName = $True
              ,HelpMessage="Parameter to be able to tell function to add an all day event by pipeline.")]
    [ValidateSet('AllDay','Normal')]
    [string]$EventType = 'Normal' # Switch AllDayEvent is tricky to send over pipeline.

	,[Parameter(Mandatory=$False
              ,ValueFromPipelineByPropertyName = $True
              ,HelpMessage="Sets the sensitivity of the calendar item")]
    [ValidateSet('Normal','Personal','Private','Confidential')]
    [string]$Sensitivity = 'Normal' # https://msdn.microsoft.com/VBA/Outlook-VBA/articles/olsensitivity-enumeration-outlook


	,[Parameter(Mandatory=$False)]
	[switch]$AllDayEvent = $false

	,[switch]$DisableReminder = $False

	,[Parameter(
        ValueFromPipelineByPropertyName = $True)]
	[int]$Reminder = 15

#    [Parameter(
#        HelpMessage="Rounds the specified MeetingStart and MeetingEnd to nearest Hour, half hour or quarter")]
#    [ValidateSet('Hour','HalfHour','Quarter')]
#    [string]$RoundToNearest
)
    BEGIN {
		# if -debug is set, change $DebugPreference so that outpout is a little less annoying.
		#	ref: http://learn-powershell.net/2014/06/01/prevent-write-debug-from-bugging-you/
		If ($PSBoundParameters['Debug']) {
			$DebugPreference = 'Continue'
		}

        # ref: https://msdn.microsoft.com/en-us/library/office/aa171432(v=office.11).aspx
        $ImportanceHash = @{
                            'Low' = 0;
                            'Normal' = 1;
                            'High' = 2;
                            }

        $BusyStatusHash = @{
                            'Free' = 0;
                            'Tentative' = 1;
                            'Busy' = 2;
                            'OutofOffice' = 3
                        }

        $SensitivityHash = @{
                            'Normal' = 0;
                            'Personal' = 1;
                            'Private' = 2;
                            'Confidential' = 3;        
                        }


        $outlookProcess = Get-process | ? { $_.Path -like "*outlook.exe"}
        if(-Not $outlookProcess) {
            Throw "Outlook needs to be started. Could not find a process matching '*outlook.exe'"
        }
    }

    PROCESS {
        Write-Debug $($pscmdlet.MyInvocation.BoundParameters | out-string)

        # This should perhaps be done with Parameter sets instead
        # Currently verification in PROCESS because checks also done on piped in data.
	    if ( ($AllDayEvent -or $EventType -eq 'AllDay') -And -Not $MeetingEnd) {
		    Throw "If an all day event is specified MeeintEnd must also be set!"
	    }

        # Check with BoundParameters because parameter MeetingDuration has a default value set.
        # $pscmdlet.MyInvocation.BoundParameters is used instead of $PSBoundParameters['MeetingDuration'] because param can also be piped.
        if ($($pscmdlet.MyInvocation.BoundParameters['MeetingDuration']) -And $MeetingEnd) {
            Throw "Both MeetingDuration and MeetingEnd cannot be specified at the same time!"
        }

        # Currently both -AllDayEvent and -EventType 'Normal' can be set, which is not a valid combination.
        # To work around that seems difficult if data is also sent over pipe.


        # Create a new appointments using Powershell
        $outlookApplication = New-Object -ComObject 'Outlook.Application' -Verbose:$False
        # Creating a instatance of Calenders
        $newCalenderItem = $outlookApplication.CreateItem('olAppointmentItem')

        $newCalenderItem.Subject = $Subject
        $newCalenderItem.Body = $Body
        $newCalenderItem.Location  = $Location

        if ($Reminder -ge 0) {
            # messy... 
            # .ReminderSet should be $True if not $DisableReminder is specified. $DisableReminder is $False by default.
            # Hence ($DisableReminder -eq $False) = $True when not $DisableReminder is specified and reminderSet is $True
            $newCalenderItem.ReminderSet = ($DisableReminder -eq $False) 
            $newCalenderItem.ReminderMinutesBeforeStart = $Reminder
        } else {
            # if $Reminer is less than 0 (-1) - reminder is turned off.
            # This is a quick and dirty to support sending Reminder over pipe.
            $newCalenderItem.ReminderSet = $False
            $newCalenderItem.ReminderMinutesBeforeStart = 0
        }
        $newCalenderItem.Importance = $ImportanceHash[$Importance]
        $newCalenderItem.BusyStatus = $BusyStatusHash[$BusyStatus]
        $newCalenderItem.Sensitivity = $SensitivityHash[$Sensitivity]

        $newCalenderItem.Start = $MeetingStart
    

        # Normalize end timestamp of meeting.
        # We don't check here if MeetingDuration and MeetingEnd both are set. That should have been taken care of earlier.
        # $MeetingEnd is checked if it's set, if not MeetingEnd is calculated frm MeetingDuration.
        # Reason for not using $MeetingDuration in the check is because $MeetingDuration will always be set because it has a default value.
        if ( -Not $MeetingEnd ) {            
            # MeetingEnd is calculated here with MeetingDuration so that only MeetingEnd is used frm here ob.
            $MeetingEnd = $MeetingStart.AddMinutes($MeetingDuration)
        }

        if($AllDayEvent -or ($EventType -eq 'AllDay') ) {
            # if AllDayEvent is specified it seems that end date must be one day further.
            $newCalenderItem.AllDayEvent = $True
            $newCalenderItem.End = $meetingend.AddDays(1)
        } else {
            $newCalenderItem.End = $MeetingEnd
        }

        Write-Verbose "Add Meeting. Subject: `"$Subject`". Body: `"$Body`". Location: `"$Location`". AllDayEvent = $($AllDayEvent -or ($EventType -eq 'AllDay')). Meeting start: $MeetingStart. Meeting end: $MeetingEnd. Reminder: $Reminder. Sensitivity: $Sensitivity"

        Write-Debug ($newCalenderItem | Out-string)

        if($pscmdlet.ShouldProcess("Calendar","Add Meeting. `"$Subject`". $MeetingStart to $MeetingEnd") ) {
            $newCalenderItem.Save()
        }
    }
}


Function _Test_Add-CalendarMeeting {
    
    $MeetingParams = @{
                    Subject = "Meeting created with Add-CalendarMeeting";
                    Body = "Content";
                    MeetingStart = (Get-date).AddDays(1);
                    MeetingDuration = 120;
                    }
    #New-Object -TypeName PSObject -Property $MeetingParams | Add-CalendarMeeting -Verbose
    New-Object -TypeName PSObject -Property $MeetingParams | Add-CalendarMeeting -Verbose -WhatIf

    <#
    Add-CalendarMeeting -Subject "Meeting created with Add-CalendarMeeting" -Body "Content" -MeetingStart (Get-date).AddDays(1) -MeetingDuration 120



    $MeetingParams = @{
                        Subject = "Meeting created with Add-CalendarMeeting";
                        Body = "Content";
                        MeetingStart = (Get-date).AddDays(1);
                        MeetingEnd = (Get-date).AddDays(4);
                        EventType = 'AllDay'
                    }
    New-Object -TypeName PSObject -Property $MeetingParams | Add-CalendarMeeting -Verbose


    Add-CalendarMeeting -Subject "Test 2017-01-19. " -MeetingStart "2017-01-19 00:00" -MeetingEnd "2017-01-19 00:00" -EventType AllDay -BusyStatus Free -Verbose
    #>

}
