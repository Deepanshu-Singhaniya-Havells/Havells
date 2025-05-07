// This code is used to get the maximum amount for a product( on which product can be sell in the market) based on the model number of the product. 
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using Newtonsoft.Json;
using System.Drawing.Drawing2D;
using System.Net.Http.Headers;
using System.Numerics;
using System.Text;

namespace HavellsNewPlugin.Actions
{

    public class GetAmountFromSAP : IPlugin
    {
        public static ITracingService tracingService = null;

        public static void productExceptSelf(List<int> temp)
        {
            List<int> forward = new List<int>();
            List<int> backward = new List<int>(temp.Count);

            int product = 1;
            foreach (var item in temp)
            {
                product *= item;
                forward.Add(product);
            }
            product = 1;
            for (int i = temp.Count - 1; i >= 0; i--)
            {
                product *= temp[i];
                backward[i] = product;
            }

            foreach (var item in forward) { Console.Write(item + " "); }
            Console.WriteLine();
            foreach (var item in backward) { Console.Write(item + " "); }
        }

        public async void Execute(IServiceProvider serviceProvider)
        {
            #region PluginConfig
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion
            try
            {
                if (context.InputParameters.Contains("ModelNumber") && context.InputParameters["ModelNumber"] is string
                        && context.InputParameters.Contains("Amount") && context.InputParameters["Amount"] is string
                        && context.Depth == 1)
                {
                    string ModelNumber = (string)context.InputParameters["ModelNumber"];
                    float Amount = (float)context.InputParameters["Amount"];
                    float ValidAmount = await GetAmount(ModelNumber);
                    if (ValidAmount == -1F)
                    {
                        context.OutputParameters["Status"] = false;
                        context.OutputParameters["Message"] = "API not working. Please try Again";
                    }
                    else if (ValidAmount == -2F)
                    {
                        context.OutputParameters["Status"] = false;
                        context.OutputParameters["Message"] = "There is no record found for corresponding Model Number";
                    }
                    else
                    {
                        context.OutputParameters["Amount"] = ValidAmount;
                        if (Amount >= 0 && Amount <= ValidAmount)
                        {
                            context.OutputParameters["Message"] = "Valid Amount";
                            context.OutputParameters["Status"] = true;
                        }
                        else
                        {
                            context.OutputParameters["Message"] = "Invalid Amount";
                            context.OutputParameters["Status"] = false;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                context.OutputParameters["Message"] = "D365 Internal Server Error : " + ex.Message;
            }
        }

        internal void toCall(ProductAmountAPIRequest request)
        {
            Task.Run(async () =>
            {
                await ProductAmountAPI(request);
            });
        }
        internal async Task<ProductAmountAPIResponse> ProductAmountAPI(ProductAmountAPIRequest request)
        {
            ProductAmountAPIResponse obj = new ProductAmountAPIResponse();
            float amount = await GetAmount(request.ModelNumber);
            if (amount == -1F)
            {
                obj.Message = "Not Able to call the API";
                obj.Amount = -1F;
            }
            else if (amount == -2)
            {
                obj.Message = "Model Number does not exist";
                obj.Amount = -2F;
            }
            else
            {
                obj.Message = "Success";
                obj.Amount = amount;
            }

            return obj;
        }
        internal void BpfTesting(IOrganizationService service)
        {
            Entity Case = service.Retrieve("incident", new Guid("f701340d-b234-44cd-b810-ff172feb2b84"), new ColumnSet(true));
            string logicalNameOfBPF = "phonetocaseprocess";
            Entity activeProcessInstance = GetActiveBPFDetails(Case, service);
            if (activeProcessInstance != null)
            {
                Guid activeBPFId = activeProcessInstance.Id; // Id of the active process instance, which will be used
                                                             // Retrieve the active stage ID of in the active process instance
                Guid activeStageId = new Guid(activeProcessInstance.Attributes["processstageid"].ToString());
                int currentStagePosition = -1;
                RetrieveActivePathResponse pathResp = GetAllStagesOfSelectedBPF(activeBPFId, activeStageId, ref currentStagePosition, service);
                if (currentStagePosition > -1 && pathResp.ProcessStages != null && pathResp.ProcessStages.Entities != null && currentStagePosition + 1 < pathResp.ProcessStages.Entities.Count)
                {
                    // Retrieve the stage ID of the next stage that you want to set as active
                    Guid nextStageId = (Guid)pathResp.ProcessStages.Entities[pathResp.ProcessStages.Entities.Count - 1].Attributes["processstageid"];
                    // Set the next stage as the active stage
                    Entity entBPF = new Entity(logicalNameOfBPF)
                    {
                        Id = activeBPFId
                    };
                    entBPF["activestageid"] = new EntityReference("processstage", nextStageId);
                    service.Update(entBPF);
                    var stateRequest = new SetStateRequest
                    {
                        EntityMoniker = new EntityReference("phonetocaseprocess", activeProcessInstance.Id),
                        State = new OptionSetValue(1), // Inactive.
                        Status = new OptionSetValue(2) // Finished.
                    };
                    service.Execute(stateRequest);
                }
            }
        }

        private Entity GetActiveBPFDetails(Entity entity, IOrganizationService service)
        {
            Entity activeProcessInstance = null;
            RetrieveProcessInstancesRequest entityBPFsRequest = new RetrieveProcessInstancesRequest
            {
                EntityId = entity.Id,
                EntityLogicalName = entity.LogicalName
            };
            RetrieveProcessInstancesResponse entityBPFsResponse = (RetrieveProcessInstancesResponse)service.Execute(entityBPFsRequest);
            if (entityBPFsResponse.Processes != null && entityBPFsResponse.Processes.Entities != null)
            {
                activeProcessInstance = entityBPFsResponse.Processes.Entities[0];
            }
            return activeProcessInstance;
        }
        private RetrieveActivePathResponse GetAllStagesOfSelectedBPF(Guid activeBPFId, Guid activeStageId, ref int currentStagePosition, IOrganizationService service)
        {
            // Retrieve the process stages in the active path of the current process instance
            RetrieveActivePathRequest pathReq = new RetrieveActivePathRequest
            {
                ProcessInstanceId = activeBPFId
            };
            RetrieveActivePathResponse pathResp = (RetrieveActivePathResponse)service.Execute(pathReq);
            for (int i = 0; i < pathResp.ProcessStages.Entities.Count; i++)
            {
                // Retrieve the active stage name and active stage position based on the activeStageId for the process instance
                if (pathResp.ProcessStages.Entities[i].Attributes["processstageid"].ToString() == activeStageId.ToString())
                {
                    currentStagePosition = i;
                }
            }
            return pathResp;
        }

        internal static async Task<float> GetAmount(string modelNumber)
        {


            float ans = -1F;
            string amount = "";

            var client = new HttpClient();
            var request = new HttpRequestMessage(HttpMethod.Post, "https://middlewaredev.havells.com:50001/RESTAdapter/ecom_priceinfo?IM_FLAG=&IM_PROJECT=D365");
            request.Headers.Add("Authorization", "Basic RDM2NV9IQVZFTExTOkRFVkQzNjVAMTIzNA==");
            request.Headers.Add("Cookie", "JSESSIONID=S-8sbup2i2P7m-XQfu34jiWrea1PjgGenHkA_SAPS48sCbzvJ7oRjokDRLt8Pe02; saplb_*=(J2EE7969920)7969950");


            Request req = new Request();
            LTTABLE reqObj = new LTTABLE();
            reqObj.MATNR = modelNumber;

            req.LT_TABLE = reqObj;

            var json = JsonConvert.SerializeObject(req);
            var content = new StringContent(json, null, "application/json");
            request.Content = content;
            var response = await client.SendAsync(request);

            if (response.IsSuccessStatusCode)
            {
                Response obj = JsonConvert.DeserializeObject<Response>(await response.Content.ReadAsStringAsync());
                if (obj != null)
                {
                    foreach (var item in obj.LT_TABLE)
                    {
                        if (item.KSCHL != null && item.KSCHL == "ZWEB")
                        {
                            amount = item.KBETR;
                        }
                    }
                }
                else
                {
                    ans = -2F;
                }
            }
            if (!string.IsNullOrEmpty(amount)) ans = float.Parse(amount);


            return ans;
        }

    }
}

public class ProductAmountAPIResponse
{
    public string? Message { get; set; }
    public float? Amount { get; set; }
}
public class ProductAmountAPIRequest
{
    public string? ModelNumber { get; set; }
}

internal class Request
{
    public LTTABLE LT_TABLE { get; set; }
}

public class LTTABLE
{
    public string? MATNR { get; set; }
    public string? KSCHL { get; set; }
    public string? DATAB { get; set; }
    public string? KBETR { get; set; }
    public string? KONWA { get; set; }
    public string? DATBI { get; set; }
    public string? DELETE_FLAG { get; set; }
    public string? CREATEDBY { get; set; }
    public string? CTIMESTAMP { get; set; }
    public string? MODIFYBY { get; set; }
    public string? MTIMESTAMP { get; set; }
}

public class Response
{
    public List<LTTABLE>? LT_TABLE { get; set; }
}
