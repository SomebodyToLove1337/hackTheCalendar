using Azure.Identity;
using Microsoft.Graph;

var scopes = new[] { ".default" };
var interactiveBrowserCredentialOptions = new InteractiveBrowserCredentialOptions
{
    ClientId = "535ea419-74af-4047-9407-cff30fbb9e3e"
};

var apointmentFilter = "startsWith(subject,'test2')";

var tokenCredential = new InteractiveBrowserCredential(interactiveBrowserCredentialOptions);

var graphClient = new GraphServiceClient(tokenCredential, scopes);

//var MeineKlassenInstanz = new MyClass();
//bool test = MeineKlassenInstanz.PrüfeAufUrlaub("xyz");

// Hole dir Daten von der GraphAPI
var o365CalRequest = await graphClient.Me.Calendar.Events.GetAsync((requestConfiguration) =>
{
    //requestConfiguration.QueryParameters.Select = new string[] { "start/dateTime", "end/dateTime", "subject"};
    requestConfiguration.QueryParameters.Filter = apointmentFilter;
    requestConfiguration.QueryParameters.Orderby = new string[] { "start/dateTime asc" };

});
try
{
    var subject = o365CalRequest.Value[0].Subject;
    var start = o365CalRequest.Value[0].Start.DateTime;
    var end = o365CalRequest.Value[0].End.DateTime;



    Console.WriteLine(subject + " - " + start + " - " + end);

    var parsedStartDate = DateTime.Parse(start);
    var parsedEndDate = DateTime.Parse(end);
    Console.WriteLine("Start Datum:" + parsedStartDate);
    Console.WriteLine("End Datum:" + parsedEndDate);

    if (parsedStartDate < parsedEndDate)
    {
        Console.WriteLine("Start Datum kleiner als Enddatum");
    }
}
catch (ArgumentOutOfRangeException ex)
{

    Console.WriteLine($"Did not find an Apointment with");
    Console.WriteLine(ex);
}

