$ReminderHours = 6            # Number of hours to set reminder to. Unused if $ReminderSet is False.
$ReminderSet = $true          # Value for reminder on/off. $False will disable reminder.

# Set start date to today to only gather future events
$Start = (Get-Date).ToShortDateString()

# Filter all future all day events with default reminder set
$Filter = "[AllDayEvent]='True' AND [Start] > '$Start' AND [ReminderSet] = 'True' AND [ReminderMinutesBeforeStart] = '1080'"

# Build COM object and pull default calendar
$Outlook = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNamespace('MAPI')
$olFolders = "Microsoft.Office.Interop.Outlook.OlDefaultFolders" -as [type]
$Calendar = $Namespace.getDefaultFolder($olFolders::olFolderCalendar)

# Get all appointments
$Appointments = $Calendar.Items

# For testing
# $Appointments.Restrict($Filter)|select Subject, Start, End, ReminderSet, ReminderMinutesBeforeStart | ft

Foreach ($Appt in $Appointments.Restrict($Filter))
{
  $Appt.ReminderMinutesBeforeStart = "$($ReminderHours * 60)"
  $Appt.ReminderSet = "$ReminderSet"
  $Appt.Save()
}
