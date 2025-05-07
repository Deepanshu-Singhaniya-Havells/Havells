using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

internal class AsyncTesting
{
    internal async Task Tocall()
    {
        await MakeTea();
        Console.ReadLine();
    }

    public async Task<string> MakeTea()
    {
        var boilingwater = BoilWater();
        Console.WriteLine("Take the cups out");
        Console.WriteLine("Put tea in cups");
        var water = await boilingwater;
        var tea = $"Pout {water} in cups";
        Console.WriteLine(tea);
        return tea;
    }
    public async Task<string> BoilWater()
    {
        Console.WriteLine("Start the kettle");
        Console.WriteLine("Waiting for the kettle");
        await Task.Delay(20000);
        Console.WriteLine("Kettle Finished Boiling");
        return "Water";
    }

    public async Task<string> CreateRecord(IOrganizationService service)
    {
        await Task.Delay(10000);
        Console.WriteLine("Creating a phone call ");
        Entity phoneCall = new Entity("phonecall");
        service.Create(phoneCall);
        Console.WriteLine("Phone call created");
        //Console.WriteLine("Delay finished");
        return "Return phone call";
    }

    public async Task FetchPhoneCalls(IOrganizationService service)
    {
        Console.WriteLine("Fetching Phone Calls");
        QueryExpression query = new QueryExpression("phonecall");
        EntityCollection phonecalls = service.RetrieveMultiple(query);
        await Task.Delay(10000);
        Console.WriteLine("Fetched Phone Calls");

    }
    public async Task<string> API(IOrganizationService service)
    {
        var create = CreateRecord(service);
        Console.WriteLine("Record Created");
        Console.WriteLine(await create);
        return "Return Record Created";
    }
}
