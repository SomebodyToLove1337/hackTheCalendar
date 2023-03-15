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

// Define scopes
var scopes = new[] { "https://graph.microsoft.com/.default" };

// Read values from configuration file
var tenantId = configuration.GetSection("Connect:AzureTenantID").Value;
var clientId = configuration.GetSection("Connect:AzureClientID").Value;
var clientSecret = configuration.GetSection("Connect:AzureClientSecret").Value;
var appointmentSubject = configuration.GetSection("UserConf:EventSubject").Value;
var o365UserIDs = configuration.GetSection("UserConf:UserID").Get<string[]>();
var externalMessage = configuration.GetSection("UserConf:ExternalMessage").Value;
var internalMessage = configuration.GetSection("UserConf:InternalMessage").Value;

// using Azure.Identity;
var options = new TokenCredentialOptions
{
    AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
};

// https://learn.microsoft.com/dotnet/api/azure.identity.clientsecretcredential
var clientSecretCredential = new ClientSecretCredential(tenantId, clientId, clientSecret, options);
var graphClient = new GraphServiceClient(clientSecretCredential, scopes);

//
//---------------------------------------------------------------------------------------------------------------------------------


foreach (var o365UserID in o365UserIDs)
{
    Console.ForegroundColor = ConsoleColor.Yellow;
    Console.WriteLine("--------------------------------------------");
    Console.WriteLine("App starting check appointments for: " + o365UserID.ToString());
    Console.ResetColor();


    // Get data from GraphAPI
    var o365CalRequest = await graphClient.Users[$"{o365UserID}"].Events.GetAsync((requestConfiguration) =>
    {
        //requestConfiguration.QueryParameters.Select = new string[] { "start/dateTime", "end/dateTime", "subject"};
        requestConfiguration.QueryParameters.Filter = $"startsWith(subject,'{appointmentSubject}')";
        requestConfiguration.QueryParameters.Orderby = new string[] { "start/dateTime asc" };

    });
    try
    {   // Store data
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

        // Replace Placeholder from config file
        externalMessage = externalMessage.Replace("[start]", parsedStartDate.ToShortDateString());
        externalMessage = externalMessage.Replace("[end]", parsedEndDate.ToShortDateString());
        internalMessage = internalMessage.Replace("[start]", parsedStartDate.ToShortDateString());
        internalMessage = internalMessage.Replace("[end]", parsedEndDate.ToShortDateString());

        // Set OoO when there is no active earlier message
        if (parsedEndOoODate <= parsedEndDate && (outOfOfficeActive.ToString() == "scheduled" || outOfOfficeActive.ToString() == "alwaysEnabled"))
        {
            Console.WriteLine("There is already an earlier OoO Message active.");
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

                    //Set external OoO Message active (OPTIONS: none, ContactsOnly, All)
                    ExternalAudience = ExternalAudienceScope.ContactsOnly,
                    ExternalReplyMessage = $"{externalMessage}",
                    InternalReplyMessage = $"{internalMessage}"
                }
            };

            var requestInformation = graphClient.Users[$"{o365UserID}"].ToGetRequestInformation();
            requestInformation.HttpMethod = Method.PATCH;
            requestInformation.UrlTemplate = "{+baseurl}/users/{user%2Did}/mailboxSettings"; //update the template to include /mailBoxSettings
            requestInformation.SetContentFromParsable<MailboxSettings>(graphClient.RequestAdapter, "application/json", mailboxSettings);

            await graphClient.RequestAdapter.SendNoContentAsync(requestInformation);
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"Found an appointment with subject: {appointmentSubject}");
            Console.WriteLine($"For user: {o365UserID}");
            Console.WriteLine($"Set auto response from: {parsedStartDate.ToShortDateString()} until {parsedEndDate.ToShortDateString()}");
            Console.WriteLine();
            Console.WriteLine("With this internal Message:");
            Console.WriteLine($"{internalMessage}");
            Console.WriteLine();
            Console.WriteLine("And this external Message:");
            Console.WriteLine($"{externalMessage}");
            Console.ResetColor();
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
                    Content = $"I configured the Auto-Reply from {parsedStartDate.ToShortDateString()} to {parsedEndDate.ToShortDateString()}",
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
        Console.ForegroundColor = ConsoleColor.Red;
        Console.WriteLine($"Did not find any apointment with the subject: {appointmentSubject}");
        Console.WriteLine($"For user: {o365UserID}");
        Console.ResetColor();

    }
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

