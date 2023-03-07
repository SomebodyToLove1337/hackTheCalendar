using Azure.Identity;
using Microsoft.Graph;

var scopes = new[] { ".default" };
var interactiveBrowserCredentialOptions = new InteractiveBrowserCredentialOptions
{
  ClientId = "535ea419-74af-4047-9407-cff30fbb9e3e"
};

var tokenCredential = new InteractiveBrowserCredential(interactiveBrowserCredentialOptions);

var graphClient = new GraphServiceClient(tokenCredential, scopes);

//var MeineKlassenInstanz = new MyClass();
//bool test = MeineKlassenInstanz.PrüfeAufUrlaub("xyz");

// Hole dir Daten von der GraphAPI
var me = await graphClient.Me.Calendar.Events.GetAsync((requestConfiguration) =>
{
  //requestConfiguration.QueryParameters.Select = new string[] { "start/dateTime", "end/dateTime", "subject"};
	requestConfiguration.QueryParameters.Filter = "startsWith(subject,'test2')";
  requestConfiguration.QueryParameters.Orderby = new string []{ "start/dateTime asc" };
  
});

var subject = me.Value[0].Subject;
var start = me.Value[0].Start.DateTime;
var end = me.Value[0].End.DateTime;

Console.WriteLine($"Hello {me?.AdditionalData}!" + " " + subject + " - " + start + " - " + end);

