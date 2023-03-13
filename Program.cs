// The client credentials flow requires that you request the
// /.default scope, and preconfigure your permissions on the
// app registration in Azure. An administrator must grant consent
// to those permissions beforehand.
using Azure.Identity;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions;
using System.Configuration;
using System.Collections.Specialized;
using Microsoft.Extensions.Configuration.Json;

var builder = new ConfigurationBuilder()
            .SetBasePath(Directory.GetCurrentDirectory())
            .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);

IConfigurationRoot configuration = builder.Build();

var scopes = new[] { "https://graph.microsoft.com/.default" };

// Multi-tenant apps can use "common",
// single-tenant apps must use the tenant ID from the Azure portal
var tenantId = configuration.GetSection("Connect:AzureTenantID").Value;

// Values from app registration
var clientId = configuration.GetSection("Connect:AzureClientID").Value;
var clientSecret = configuration.GetSection("Connect:AzureClientSecret").Value;

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
var apointmentSubject = configuration.GetSection("UserConf:MailSubject").Value;
var o365UserID = configuration.GetSection("UserConf:UserID").Value;
var externalMessage = configuration.GetSection("UserConf:ExternalMessage").Value;
var internalMessage = configuration.GetSection("UserConf:InternalMessage").Value;

// Get data from GraphAPI
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

    //Die ausgelesenen Werte überprüfen
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

    //erste IF abfrage rein zum testen  && (outOfOfficeActive != "scheduled" or outOfOfficeActive == "AlwaysEnabled"))
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
                ScheduledEndDateTime = endDTTZ,
                //OoF Message aktiv auch für extern (none, ContactsOnly, All)
                ExternalAudience = ExternalAudienceScope.ContactsOnly,
                ExternalReplyMessage = externalMessage,
                InternalReplyMessage = internalMessage
            }
        };

        var requestInformation = graphClient.Users[$"{o365UserID}"].ToGetRequestInformation();
        requestInformation.HttpMethod = Method.PATCH;
        requestInformation.UrlTemplate = "{+baseurl}/users/{user%2Did}/mailboxSettings"; //update the template to include /mailBoxSettings
        requestInformation.SetContentFromParsable<MailboxSettings>(graphClient.RequestAdapter, "application/json", mailboxSettings);

        await graphClient.RequestAdapter.SendNoContentAsync(requestInformation);
    }

    //Send Mail
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
                    Address = $"{o365UserID}",
                },
            },
        },
        },
        SaveToSentItems = true,
    };
    await graphClient.Users[$"{o365UserID}"].SendMail.PostAsync(requestBody);
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

