# HackTheCalendar - Update your Out-of-Office-Auto-Replies automatically

This is a simple console application built using .NET v7.0 that connects to your calendar via Microsoft Graph, to determine if their are any out-of-office events scheduled. If so, the app will automatically update the AutomaticReplySettings of the user's mailbox to schedule an Out-Of-Office notice.

## Creators 🚀

SomebodyToLove1337
maxhe87

### 1. Register an Azure Active Directory app

Every app that uses Azure AD for authentication must be registered with Azure AD. You can register app through Azure Portal or by using Azure CLI. Please follow one of the options to register your app:

<details>
  <summary>Bla bla bla</summary>

* [Install Azure CLI](https://learn.microsoft.com/cli/azure/install-azure-cli?view=azure-cli-latest) if you haven't already.
* Register your app on Microsoft Azure, by creating a new Azure AD app registration:
  * <details>
      <summary>On macOS/Linux/in Bash</summary>

    * Open terminal and change the working directory to the root of this project
    * To make the setup script executable, run `chmod +x ./setup.sh`
    * To register the app, run `./setup.sh`
    * When prompted, sign in with your **Microsoft 365 developer sandbox account**

    </details>
  * <details>
      <summary>On Windows/in PowerShell</summary>

    * Open PowerShell and change the working directory to the root of this project
    * To register the app, run `.\setup.ps1`
    * When prompted, sign in with your **Microsoft 365 developer sandbox account**

    </details>

</details>


