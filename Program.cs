// The client credentials flow requires that you request the
// /.default scope, and preconfigure your permissions on the
// app registration in Azure. An administrator must grant consent
// to those permissions beforehand.
using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;

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

var apointmentSubject = "test2";

// Hole dir Daten von der GraphAPI
var o365CalRequest = await graphClient.Users["64a018d3-7aaa-45fa-a63b-3d6528cbfe09"].Events.GetAsync((requestConfiguration) =>
{
    //requestConfiguration.QueryParameters.Select = new string[] { "start/dateTime", "end/dateTime", "subject"};
    requestConfiguration.QueryParameters.Filter = $"startsWith(subject,'{apointmentSubject}')";
    requestConfiguration.QueryParameters.Orderby = new string[] { "start/dateTime asc" };

});
try
{
    var subject = o365CalRequest.Value[0].Subject;
    var start = o365CalRequest.Value[0].Start.DateTime;
    var end = o365CalRequest.Value[0].End.DateTime;

var o365CalRequest2 = await graphClient.Users["64a018d3-7aaa-45fa-a63b-3d6528cbfe09"].GetAsync((requestConfiguration) =>
 {
    requestConfiguration.QueryParameters.Select = new string[] { "mailboxSettings" };
    //requestConfiguration.QueryParameters.Filter = $"startsWith(subject,'{apointmentSubject}')";
    //requestConfiguration.QueryParameters.Orderby = new string[] { "start/dateTime asc" };

}); 
var mailboxSettings = o365CalRequest2?.MailboxSettings;
var outOfOfficeActive = mailboxSettings.AutomaticRepliesSetting.Status;
var outOfOfficeStart = mailboxSettings.AutomaticRepliesSetting.ScheduledStartDateTime.DateTime;
var outOfOfficeEnd = mailboxSettings.AutomaticRepliesSetting.ScheduledEndDateTime.DateTime;

Console.WriteLine(outOfOfficeActive + " - " + outOfOfficeStart + " - " + outOfOfficeEnd);

    Console.WriteLine(subject + " - " + start + " - " + end);

    var parsedStartDate = DateTime.Parse(start);
    var parsedEndDate = DateTime.Parse(end);
    var parsedStartOoODate = DateTime.Parse(outOfOfficeStart);
    var parsedEndOoODate = DateTime.Parse(outOfOfficeEnd);
    Console.WriteLine("Start Datum:" + parsedStartDate);
    Console.WriteLine("End Datum:" + parsedEndDate);
    Console.WriteLine("Start Datum OoO:" + parsedStartOoODate);
    Console.WriteLine("End Datum OoO:" + parsedEndOoODate);


    if (parsedEndOoODate <= parsedEndDate)
    {
        Console.WriteLine("Aktuell ist noch eine OoO Message aktiv!");
    }
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

var requestBody = new Microsoft.Graph.Users.Item.SendMail.SendMailPostRequestBody
{
	Message = new Message
	{
		Subject = "Meet for lunch?",
		Body = new ItemBody
		{
			ContentType = BodyType.Text,
			Content = "The new cafeteria is open.",
		},
		ToRecipients = new List<Recipient>
		{
			new Recipient
			{
				EmailAddress = new EmailAddress
				{
					Address = "DiegoS@xby1p.onmicrosoft.com",
				},
			},
		},
		CcRecipients = new List<Recipient>
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
await graphClient.Users["AdminBS@xby1p.onmicrosoft.com"].SendMail.PostAsync(requestBody);
