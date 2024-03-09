
# Outlook to Google Calendar Sync for Notion Integration

## Motivation

I recently started to use Notion Calendar as the main source of truth for my time management. I love the way it integrates with Notion, it has a great native Mac App and the UI is awesome. One particular feature that I enjoy is the sticky push notification I get on my screen whenever a new meeting is starting.

*Pain #1*: Notion Calendar does not yet support the integration of Outlook Calendars. As a consultant who often has to manage communication using client domain emails, almost always using Outlook, this is a big pain.

### Solution #1

It turns out that Outlook allows you to easily publish your calendar via .ICS link as [explained here](https://support.microsoft.com/en-gb/office/share-your-calendar-in-outlook-com-0fc1cb48-569d-4d1e-ac20-5a9b3f5e6ff2)

It adds a new feed on your Google Calendar and it also shows up on Notion Calendar. 

*Pain #2*: This feed is only updated by Google every 8-12 hours. I cannot afford to miss a client meeting, so this solution is not for me.

### Solution #2

I never really used Google App Script, it feels like one of those technologies that make sense but you never found a usecase for it. Well, here it is.
I created a small script that reads the published ICS Outlook Calendar and creates/deletes/edits the events on a specific Google Calendar, and I can define my own refresh frequency.

## How to set up this Automation

To replicate this solution, follow these steps:

1. **Prepare your Google Calendar**
	* Create a new Calendar in your Google Calendar. This is not required, you can just use your main Calendar, but I like to keep things neat and have one separate Calendar per Outlook integration
	* Note down the Calendar ID (under Settings -> Integration)
2. **Share your Outlook Calendar**
	* Activate sharing of your Outlook calendar and obtain the ICS sharing URL through the Outlook UI. [explained here](https://support.microsoft.com/en-gb/office/share-your-calendar-in-outlook-com-0fc1cb48-569d-4d1e-ac20-5a9b3f5e6ff2)
3. **Create a new Google App Script**
	* Create a new Script
	* Copy-paste the code from [this template](https://github.com/2bd-ventures/outlook-google-sync/blob/main/Code.gs)
	* Add the Google Calendar API Service using the respective button on the left side 
	* Alter the variables *CALENDAR_ID* and *ICAL_URL*
	* Make sure the fuction set to run is the *updateCalendars* and hit Run (this will run for the first time and will require you to give permissions for the script to act on your behalf)
	* *Pain #3*: You will see the on the script I invite a dummy email to each meeting, this is due to Notion Calendar's settings which will not create the sticky push notification unless an event includes a guest 🤷
4. **Create a new trigger for the script**
	* Create a new time-based trigger for the script and adjust the update frequency that you wish.

🎉 Voilá!
