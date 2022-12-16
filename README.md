# POTA-Notifications
Receive SMS Text messages and Emails when certain POTA stations are on the air.
What I have here is the latest POTA Notification PowerShell script. This program requires a Gmail mailbox set up with two step verification. This will allow you to set up a separate password to allow this program to send out emails. This PS script uses a watchlist.csv file to define the criteria for notifications.
The following one or more criteria can be used to provide notifications:
activator, frequency, Band, mode, reference, locationDesc.

Activator - Callsign of the Activation station, wildcard set to *
Frequency - Really not used - default to *
Band - Band activator is operating on - default to *
Mode - Mode activator is operating on - default to *
Reference - This is the Federal/State Park identifier IE: K-1234 - default to *
LocationDesc - USA would be US-XX where XX is the 2 letter state abbreviation - default to *

This program also interfaces with N3FJP logging software. It performs 2 functions. First once an activation has met the watchlist search criteria a call is made to N3FJP to see if the activation has been already logged. If so then the notification is suppressed. If the notification was found in N3FJP then the Park and Park Name are written back to the logged entry.

This program has evolved over the basic need to complete my WAS-POTA. I want to know when certain parks are activated in states I need to complete my WAS. When I am home I can set the MODE to * because I have access to my radio and can work the station with any mode. When I am remote I set the MODE to FT8 which is the only mode I presently have available when operating my station remote.

This program runs every 15 minutes on my Task Scheduler. Often times this gives me plenty of time to find the station that met my criteria and attempt to work him.

This program is a work-in-progress. Keep an eye on the releases because I will be adding new features on a regular basis.
