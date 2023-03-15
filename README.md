# HackTheCalendar - a Hack-Together project

## Updates your Out-of-Office-Auto-Replies automatically

This is a simple console application built using .NET v7.0 that connects to your calendar via Microsoft Graph, to determine if their are any out-of-office events scheduled. If so, the app will automatically update the AutomaticReplySettings of the user's mailbox to schedule an Out-Of-Office notice.

## Creators 🚀


[SomebodyToLove1337](https://github.com/SomebodyToLove1337)
[maxhe87](https://github.com/maxhe87)


## How to configure the app 🚀
You will find an file names "appsettings.json".
The file looks like this:
```{
    "Connect": {
      "AzureTenantID" :  "YOUR-TENANT-ID", 
      "AzureClientID": "YOUR-CLIENT-ID",
      "AzureClientSecret" :  "YOUR-CLIENT-SECRET" 
      },
      "UserConf": {
        "MailSubject": "Vacation",
        "UserID": ["your_mail1@company.com", "your_mail2@company.com"],
        "ExternalMessage": "Greetings! Thank you for your email. Im out of the office from [start] until [end]. Ill reply to your email as soon as I can upon my return.",
        "InternalMessage": "Hello colleagues, im out of the office till [end]."
      }
  }
  ```

  **AzureTenantID**:
  Enter your TenantID from Azure Active Directory.
  You will find it under your registered App.

  **AzureClientID**:
  Enter your TenantID from Azure Active Directory.
  You will find it under your registered App.

  **AzureClientSecret**:
  Enter your TenantID from Azure Active Directory.
  You will find it under your registered App -> certificate and secret.

  **MailSubject**:
  You can enter the subject for the appointment what the app is searching for.
  The app always looks from Start. So if you enter "Holiday" and your appointment names "Holidays" then he will find it.
  When you enter "Holiday" as keyword and your appointment names like "My Cool Holiday" the app will NOT find it.

  **UserID**:
  you can enter any Mail Adress from your O365 tenant here.
  The app will check this email adress.
  You can add a comma separated list e.q. ["email1@mycompany.com", "email2@mycompany.com", "email3@mycompany.com"]

  **ExternalMessage and InternalMessage**:
  This the part where you can enter your Message for the auto response.

**Azure AD API-Permissions**
You will need the following permissions for your registered app:
- Calendars.ReadWrite (Application)
- Mail.Send (Application)
- MailboxSettings.ReadWrite (Application)

[![Hack Together: Microsoft Graph and .NET](https://img.shields.io/badge/Microsoft%20-Hack--Together-orange?style=for-the-badge&logo=microsoft)](https://github.com/microsoft/hack-together)