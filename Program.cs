// The client credentials flow requires that you request the
// /.default scope, and preconfigure your permissions on the
// app registration in Azure. An administrator must grant consent
// to those permissions beforehand.
using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;
using Microsoft.Kiota.Abstractions;

var scopes = new[] { "https://graph.microsoft.com/.default" };

// Multi-tenant apps can use "common",
// single-tenant apps must use the tenant ID from the Azure portal
var tenantId = "93e5635d-6391-453c-afb7-1776f501135d";

// Values from app registration
var clientId = "535ea419-74af-4047-9407-cff30fbb9e3e";
var clientSecret = "zjp8Q~gwQAYuHOXdfbE~TKm4N2ePYq4r5hOjgcj1";

// using Azure.Identity;
var options = new TokenCredentialOptions
{
    AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
};

// https://learn.microsoft.com/dotnet/api/azure.identity.clientsecretcredential
var clientSecretCredential = new ClientSecretCredential(
    tenantId, clientId, clientSecret, options);

var graphClient = new GraphServiceClient(clientSecretCredential, scopes);

//
//---------------------------------------------------------------------------------------------------------------------------------
//Alles drüber ist für die Authentifizierung zu Microsoft Graph

//Variablen welche noch in eine config Datei ausgelagert werden sollten
var apointmentSubject = "test2";
var o365UserID = "64a018d3-7aaa-45fa-a63b-3d6528cbfe09";

//Hier müsste noch eine Funktion gebaut werden bei welcher die E-Mail ausgelesen wird und die ID zurückgegeben wird

// Hole dir Daten von der GraphAPI
var o365CalRequest = await graphClient.Users[$"{o365UserID}"].Events.GetAsync((requestConfiguration) =>
{
    //requestConfiguration.QueryParameters.Select = new string[] { "start/dateTime", "end/dateTime", "subject"};
    requestConfiguration.QueryParameters.Filter = $"startsWith(subject,'{apointmentSubject}')";
    requestConfiguration.QueryParameters.Orderby = new string[] { "start/dateTime asc" };

});
try
{   //Ausgelesene Kalender Daten in Variable speichern
    var subject = o365CalRequest.Value[0].Subject;
    var start = o365CalRequest.Value[0].Start.DateTime;
    var end = o365CalRequest.Value[0].End.DateTime;
    var timezone = o365CalRequest.Value[0].Start.TimeZone;

    // DateTimeTimeZone Werte erstellen
    var startDTTZ = new DateTimeTimeZone();
    var endDTTZ = new DateTimeTimeZone();
    startDTTZ.DateTime = start.ToString();
    startDTTZ.TimeZone = timezone;
    endDTTZ.DateTime = end.ToString();
    endDTTZ.TimeZone = timezone;

    //Mailbox Settings sind unterhalb des "Users" kontext, es werden die kompletten MailboxSettings ausgelesen
    var o365CalRequest2 = await graphClient.Users[$"{o365UserID}"].GetAsync((requestConfiguration) =>
     {
         requestConfiguration.QueryParameters.Select = new string[] { "mailboxSettings" };
         //requestConfiguration.QueryParameters.Filter = $"startsWith(subject,'{apointmentSubject}')";
         //requestConfiguration.QueryParameters.Orderby = new string[] { "start/dateTime asc" };

     });

    //Ausgelesene Werte in Variablen speichern
    var mailboxSettings = o365CalRequest2?.MailboxSettings;
    var outOfOfficeActive = mailboxSettings.AutomaticRepliesSetting.Status;
    var outOfOfficeStart = mailboxSettings.AutomaticRepliesSetting.ScheduledStartDateTime.DateTime;
    var outOfOfficeEnd = mailboxSettings.AutomaticRepliesSetting.ScheduledEndDateTime.DateTime;

    //Die Ausgelesenen Werte überprüfen
    Console.WriteLine("OoO:" + outOfOfficeActive + " - " + outOfOfficeStart + " - " + outOfOfficeEnd);
    Console.WriteLine("Event:" + subject + " - " + start + " - " + end);

    //Convertieren des GraphAPI Rückgabewert in DateTime Format
    var parsedStartDate = DateTime.Parse(start);
    var parsedEndDate = DateTime.Parse(end);
    var parsedStartOoODate = DateTime.Parse(outOfOfficeStart);
    var parsedEndOoODate = DateTime.Parse(outOfOfficeEnd);
    Console.WriteLine("Start Datum:" + parsedStartDate);
    Console.WriteLine("End Datum:" + parsedEndDate);
    Console.WriteLine("Start Datum OoO:" + parsedStartOoODate);
    Console.WriteLine("End Datum OoO:" + parsedEndOoODate);

    //erste IF abfrage rein zum testen
    if (parsedEndOoODate > parsedEndDate)
    {
        Console.WriteLine("There is alreay an earlier OoO Message active.");
    }
    else
    {
                    mailboxSettings = new MailboxSettings
            {
                AutomaticRepliesSetting = new AutomaticRepliesSetting
                {
                    // ScheduledStartDateTime = new DateTimeTimeZone(),
                    // ScheduledEndDateTime = new DateTimeTimeZone()
                    Status = AutomaticRepliesStatus.Scheduled,
                                                  
                    ScheduledStartDateTime = startDTTZ,
                    ScheduledEndDateTime = endDTTZ
                }
            };
            var requestInformation = graphClient.Users[$"{o365UserID}"].ToGetRequestInformation();
            requestInformation.HttpMethod = Method.PATCH;
            requestInformation.UrlTemplate = "{+baseurl}/users/{user%2Did}/mailboxSettings";//update the template to include /mailBoxSettings
            requestInformation.SetContentFromParsable<MailboxSettings>(graphClient.RequestAdapter, "application/json", mailboxSettings);

            await graphClient.RequestAdapter.SendNoContentAsync(requestInformation);

 /*        //Set OoO message to the absence event
        var patchUser = new User
        {
            MailboxSettings = new MailboxSettings
            {
                AutomaticRepliesSetting = new AutomaticRepliesSetting
                {
                    ScheduledStartDateTime = startDTTZ,
                    ScheduledEndDateTime = endDTTZ
                }
            }
        };


        //PATCH the changed configuration to the User/Mailbox Settings - @Benedikt: hier bekomme ich einen Berechtigungsfehler.

        try
{
    await graphClient.Users[$"{o365UserID}"].PatchAsync(patchUser);
}
catch (ODataError odataError)
{
    Console.WriteLine(odataError.Error.Code);
    Console.WriteLine(odataError.Error.Message);
    throw;
} */
        
    }

    //Mail verschicken
    var requestBody = new Microsoft.Graph.Users.Item.SendMail.SendMailPostRequestBody
    {
        Message = new Message
        {
            Subject = "Out of Office Auto Response has been configured",
            Body = new ItemBody
            {
                ContentType = BodyType.Text,
                Content = $"I configured the Auto-Response from {outOfOfficeStart} to {outOfOfficeEnd}",
            },
            ToRecipients = new List<Recipient>
        {
            new Recipient
            {
                EmailAddress = new EmailAddress
                {
                    Address = "AdminBS@xby1p.onmicrosoft.com",
                },
            },
        },
        },
        SaveToSentItems = true,
    };
    await graphClient.Users["64a018d3-7aaa-45fa-a63b-3d6528cbfe09"].SendMail.PostAsync(requestBody);
}
catch (ArgumentOutOfRangeException ex)
{

    Console.WriteLine($"Did not find an Apointment with the subject: {apointmentSubject}");
    Console.WriteLine(ex);
}

/* var MailboxSettingsDiv =>
{
    "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#users('username')/mailboxSettings",
    
    "timeZone": "W. Europe Standard Time",
    "delegateMeetingMessageDeliveryOptions": "sendToDelegateOnly",
    "dateFormat": "dd.MM.yyyy",
    "timeFormat": "HH:mm",
    "userPurpose": "user",
    "automaticRepliesSetting": {
        "status": "disabled",
        "externalAudience": "none",
        "internalReplyMessage": "",
        "externalReplyMessage": "",
        "scheduledStartDateTime": {
            "dateTime": "2022-08-27T08:00:00.0000000",
            "timeZone": "UTC"
        },
        "scheduledEndDateTime": {
            "dateTime": "2022-08-28T08:00:00.0000000",
            "timeZone": "UTC"
        }
    }, */

