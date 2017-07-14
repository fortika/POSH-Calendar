# dot-source files containing each function
# makes managing the module easier

. $PSScriptRoot\Add-CalendarMeeting.ps1

. $PSScriptRoot\Get-CalendarItems.ps1

Export-ModuleMember -Function Add-CalendarMeeting
Export-ModuleMember -Function Get-CalendarItems
