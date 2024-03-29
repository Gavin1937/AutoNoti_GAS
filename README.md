
# AutoNoti_GAS

## Automatically Send Scheduled Notification with Google Apps Script.

## Prerequisite

* Computer
* Internet
* A Google Account
* Basic knowledge about spreadsheet. (e.g. knowing how does spreadsheet range works, like !A2:D)

## Recommand Skills

* basic knowledge of JavaScript

## Google Apps Script Required Permissions

* See, edit, create, and delete all your Google Docs documents
  * For reading Google Doc email templates
* See, edit, create, and delete all your Google Sheets spreadsheets
  * For reading Data SpreadSheet & writing cache data
* Connect to an external service
  * For parsing Google Doc text into html
* Send email as you
  * For sending notification emails

## Setup

1. Open Browser and go to your Google Drive.
2. Create a new folder for this project.
3. Inside the folder, create a new Google SpreadSheet for storing all the data.

### Setup Data SpreadSheet

#### Setup Schedule Sheet

1. Inside your Data SpreadSheet, name the first sheet as **Schedule** (case & order matters)

2. Put all your schedule information into **Schedule**

3. Sheet **Schedule** must have following items:
   1. a column stores date. (formatted like mm/dd or mm/dd/yyyy)
   2. a column stores sermon person name
   3. a column stores worship person name
   4. **cells at column A stores year indicator**
      * You need to input indicators for the year of schedule
      * This script uses **column A** of Sheet **Schedule** for that
      * The indicator format for this script is: **year=xxxx** where **xxxx** is 4 digit year number

#### Setup Contact Sheet

1. Inside your Data SpreadSheet, name the second sheet as **Contact** (case & order matters)

2. Sheet **Contact** must have following items:
   1. A **name** column (person's name in schedule)
   2. A **refer_name** column (how do you want to call a person in email notification)
   3. A **email** column 
   4. A **is_admin** column

3. Put all your data into dedicated columns

#### Setup MessagesTemplate Sheet

1. Inside your Data SpreadSheet, name the second sheet as **MessagesTemplate** (case & order matters)

2. Sheet **MessagesTemplate** must formatted exactly like follow table:

| SpreadSheet Column A | SpreadSheet Column B   |
|----------------------|------------------------|
| sermon_msg_doc       | url to sermon_msg_doc  |
| worship_msg_doc      | url to worship_msg_doc |

* **sermon_msg_doc** is a Google Doc that stores the template for sermon person notification (more on that in next section)
* **worship_msg_doc** is a Google Doc that stores the template for worship person notification (more on that in next section)

3. Put all your data into dedicated columns


### Setup Template Google Docs

After your finish setting up your Data SpreadSheet, your need to set up notification templates using Google Docs

1. Create two Google Docs under your project folder for sermon notification and worship notification.

2. Inside each Google Docs, type in your template notification

3. You can use following macros in your template:
   * ```${sermon_name}``` => name of sermon person
   * ```${worship_name}``` => name of worship person
   * ```${admin_name}``` => name of admins (people who have is_admin set to 1 in **Contact** Sheet)
   * ```${admin_email}``` => email of admins (people who have is_admin set to 1 in **Contact** Sheet)

4. Template example:

```
    Hello, ${sermon_name}, you are in charge of this week's sermon.
    
    If you have any question, please contact ${admin_name}.
    
    Email: ${admin_email}
```

   * **Note that, only the info of first admin in Contact Sheet will be place into templates.**

5. After you finish your template Google Docs, be sure to paste their urls to **MessagesTemplate** Sheet.

### Setup Google Apps Script

After your finish setting up both Data SpreadSheet and template Google Docs, you need to setup the Apps Script.

1. Go to your Data SpreadSheet.

2. On the top, find and click: **Extensions -> Apps Script**

3. **[OPTIONAL]** In the new tab, rename your project and current file.
   * rename project: click on the text beside "Apps Script" logo
   * rename file: click on the tree dot button from item on the left side "Files" tab

4. Copy all the contents from **autonoti.gs** file and replace them into Apps Script editor.

5. Fill in the blank for **CONFIGURATION** variable on top of the script.

```js
   // All the columns are counting start from 0 instead of 1 
   var CONFIGURATION = {
     SPREAD_SHEET_URL: "URL to Google SpreadSheet",
     SCHEDULE_SHEET: {
       RANGE: "!A:A",         // select range of Schedule Sheet
       DATE_COLUMN: 0,        // which column is for Date
       SERMON_COLUMN: 0,      // which column is for Sermon Person
       WORSHIP_COLUMN: 0      // which column is for Worship Person
     },
     CONTACT_SHEET: {
       RANGE: "!A:A",         // select range of Contact Sheet
       NAME_COLUMN: 0,        // which column is for name
       REFER_NAME_COLUMN: 0,  // which column is for refer_name
       EMAIL_COLUMN: 0,       // which column is for email
       IS_ADMIN_COLUMN: 0     // which column is for is_admin
     },
     NOTI_MIN_INTERVAL: 0,    // minimum time inteval of notification, in weeks
     NOTI_AT_WKDAYS: ["MON"], // send notification at specific weekdays, this is a list.
                              // ["MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN"]
     NOTI_HOUR_RANGE: [0, 0], // notification hour range of a day. (0-23)
                              // Script will send notification in time range >= first and <= second 
     EMAIL_SUBJECT: "Email Subject for all emails"
   }
```

6. Click on **Run** button to test the script
   * When you click on **Run** button, Google will ask your to sign in first, and then this script will ask you for some permissions (Check out [Google Apps Script Required Permissions](#google-apps-script-required-permissions) for more detail)
   * Choose an account to sign in
   * Google may give a **Google hasn’t verified this app** warning message, just click **Advanced** and then **Go to your project name(unsafe)** 
   * After you reviewed all the permissions, click **Allow**.
   * You only need to do this once.


7. Finally, go to **Triggers** tab on the left side and add a trigger for your project
   * you can set when does script run after your click **Add Trigger** button.

### Testing

I suggest you set the script trigger to run at a short time first, so you can test your project. If everything works fine, you can change the trigger to what you want.

## Features

### Spam Guard

To avoid message spamming, the script has a Spam Guard.

The Script will use `NOTI_MIN_INTERVAL`, `NOTI_AT_WKDAYS`, and `NOTI_HOUR_RANGE` parameters in `CONFIGURATION` for Spam Guard.

* `NOTI_MIN_INTERVAL`:
  * minimum interval in weeks between two notifications
  * for example:
    * if `NOTI_MIN_INTERVAL` set to 1
      * `first notification` was sent at week 1 Thursday
      * `second notification` will be send after week 2 Monday 00 AM
      * week number changes between two notification, and this week number delta is $\ge$ `NOTI_MIN_INTERVAL`
    * if `NOTI_MIN_INTERVAL` set to 0
      * then two notifications can be send in the same week at different date.
      * you should use this with `NOTI_AT_WKDAYS`
* `NOTI_AT_WKDAYS`:
  * send notification at specified weekdays
  * this is a list of string to specify weekdays, ["MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN"]
  * if you want to send multiple notifications in the same week, you should set `NOTI_MIN_INTERVAL` to 0
  * for example:
    * if `NOTI_MIN_INTERVAL` set to 0 and `NOTI_AT_WKDAYS` set to ["MON", "TUE"]
      * this script will send a notification on Monday and Tuesday every week
* `NOTI_HOUR_RANGE`:
  * notification hour range of a day. (0-23)
  * this parameter will be applied on specified weekday
  * this is a list of two integers [first, second]
  * script will send notification in time range `>= first` and `<= second`
  * for example:
    * if `NOTI_MIN_INTERVAL` set to 0, `NOTI_AT_WKDAYS` set to ["MON", "TUE"], and `NOTI_HOUR_RANGE` set to [9, 20]
      * this script will send a notification on Monday and Tuesday between 9 AM to 20 PM every week

In order for Spam Guard to work, this script needs to store some cache data. These data will be store in Sheet **cache** in your Data SpreadSheet.

Sheet **cache** will be automatically create by the script, **DO NOT EDIT OR DELETE THIS SHEET**.

