using AE01.AmazonOnboarding;
using AE01.AMC;
using AE01.Assign_Roles;
using AE01.Call_Masking;
using AE01.Finished_Good_Replacement;
using AE01.Inventory;
using AE01.Miscellaneous;
using AE01.Miscellaneous.Production_Support;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using NLog;
using NLog.Config;
using NLog.Layouts;
using NLog.Targets;
using System.Drawing;
using System.Runtime.CompilerServices;
using System.Security.Cryptography.Xml;


internal class Program
{
    private static Logger logger = LogManager.GetCurrentClassLogger();


    private static async Task Main(string[] args)
    {
        var config = new LoggingConfiguration();

        string fileName = "CallMaskingLogs " + DateTime.Today.ToString("dd-MM-yyyy") + ".log";

        var target = new FileTarget { FileName = fileName };
        target.Layout = new JsonLayout
        {
            Attributes =
            {
                    new JsonAttribute("timestamp", "${date:format=yyyy-MM-ddTHH:mm:ssZ}"),
                    new JsonAttribute("level", "${level}"),
                    new JsonAttribute("message", "${message}"),
                    new JsonAttribute("exception", "${exception:format=Message}")
            }
        };

        config.AddTarget("file", target);

        var consoleTarget = new ConsoleTarget();
        config.AddTarget("console", consoleTarget);

        var fileRule = new LoggingRule("*", LogLevel.Info, target);
        config.LoggingRules.Add(fileRule);

        var consoleRule = new LoggingRule("*", LogLevel.Debug, consoleTarget);
        config.LoggingRules.Add(consoleRule);

        LogManager.Configuration = config;

        logger.Info("Hello from Nlog");
        logger.Error("Error from Nlog");
        logger.Fatal("Fatal from Nlog");
        logger.Warn("Warn from Nlog");
        logger.Debug("Debug from Nlog");
        logger.Trace("Trace from Nlog");

        Console.WriteLine("Program Started");
        var _service = D365.CreateDataverseConnection();
        //var _azureContainer = D365.CreateBlobConnection("images");

        IOrganizationService service = await _service;

        SystemAnalytics systemAnalyticsObj = new SystemAnalytics(service);
        systemAnalyticsObj.GetSystemAnalytics();


        MetaData metaDataObj = new MetaData(service);
        metaDataObj.ExportMetaData("C:\\Users\\39054\\OneDrive - Havells\\Desktop\\Havells Projects\\AE01\\MetaData.json");
        Report reportObj = new Report(service);
        reportObj.Caluclate();

        AssignRoles assignRolesObj = new AssignRoles(service);
        assignRolesObj.PrintAllRoles(); 
        assignRolesObj.RemoveRoleFromUser(new Guid("6e2ccd4d-da97-ed11-aad1-6045bdac5778"), "System Customizer");
        assignRolesObj.GetUsersForCallMasking();        

        AmazonOnbaording amazonOnbaordingObj = new AmazonOnbaording(service);
        Console.WriteLine("==========================================================================");
        for (int i = 0; i < 500; i++)
        {
            Console.WriteLine();
            Console.WriteLine($"Sending {i + 1} request"); 
            TimeSpan TotalResponseTime = await amazonOnbaordingObj.MeasureApiCall("https://middlewaredev.havells.com:50001/RESTAdapter/amazoncrm/getunackrequest");
            await Task.Delay(3000);            
            Console.WriteLine(); 
        }
        
        await amazonOnbaordingObj.FetchUnacknowledgedIds();            
        ProductReplacement productReplacementObj = new ProductReplacement(service);
        productReplacementObj.Program();

        TestingInventory testingInventoryObj = new TestingInventory(service);
        await testingInventoryObj.ToCall();

        //string mess = await _azureContainer; 
        TokenValidation tokenValidationObj = new TokenValidation(service);
        tokenValidationObj.Main();

        IotServiceCall iotServiceCallObj = new IotServiceCall(service);
        iotServiceCallObj.GetIoTServiceCalls(new Guid("6f7071e8-0644-ee11-be6f-6045bdac5e8e"));

        //AsyncTesting obj = new AsyncTesting();

        //Console.WriteLine(await obj.API(service));
    }

    private async void CreateConnection()
    {
        // to Access Dataverse using WebAPI
        var client = new HttpClient();
        var request = new HttpRequestMessage(HttpMethod.Post, "https://login.microsoftonline.com/7b7dc2f5-4e6a-4004-96dd-6c7923625b25/oauth2/token");
        request.Headers.Add("Accept", "application/json");
        request.Headers.Add("Cookie", "fpc=AgQDRWDayw9DkGgHra-eTaClIQdRAQAAAMEhZd0OAAAA; stsservicecookie=estsfd; x-ms-gateway-slice=estsfd");
        var content = new MultipartFormDataContent();
        content.Add(new StringContent("41623af4-f2a7-400a-ad3a-ae87462ae44e"), "client_id");
        content.Add(new StringContent("r/qy6jgTisxeKZ7T3nYTVbhswXc5pYSVoSp5XFKJtQ8="), "client_secret");
        content.Add(new StringContent("https://havellscrmdev1.crm8.dynamics.com"), "resource");
        content.Add(new StringContent("client_credentials"), "grant_type");
        request.Content = content;
        var response = await client.SendAsync(request);
        response.EnsureSuccessStatusCode();
        Console.WriteLine(await response.Content.ReadAsStringAsync());
    }
}