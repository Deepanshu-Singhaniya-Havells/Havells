using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Extensions.Hosting;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net.Sockets;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace AE01.AmazonOnboarding
{
    internal class AmazonOnbaording(IOrganizationService _service)
    {
        private IOrganizationService service = _service;

        public class Data
        {
            
        }

        internal async Task FetchUnacknowledgedIds()
        {

            Entity _InventoryRMA = service.Retrieve("hil_inventoryrma", new Guid("450aaba0-da0e-f011-9989-7c1e523d5ba0"), new ColumnSet("hil_franchise", "hil_warehouse", "hil_returntype"));
            if (_InventoryRMA != null)
            {
                EntityReference _franchise = _InventoryRMA.GetAttributeValue<EntityReference>("hil_franchise");
                EntityReference _warehouse = _InventoryRMA.GetAttributeValue<EntityReference>("hil_warehouse");
                EntityReference _returntype = _InventoryRMA.GetAttributeValue<EntityReference>("hil_returntype");

                string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                              <entity name='hil_inventoryproductjournal'>
                                                <attribute name='hil_inventoryproductjournalid' />
                                                <attribute name='hil_name' />
                                                <attribute name='createdon' />
                                                <order attribute='hil_name' descending='false' />
                                                <filter type='and'>
                                                  <condition attribute='statecode' operator='eq' value='0' />
                                                  <condition attribute='hil_transactiontype' operator='eq' value='3' />
                                                  <condition attribute='hil_franchise' operator='eq' value='{_franchise.Id}' />
                                                  <condition attribute='hil_warehouse' operator='eq' value='{_warehouse.Id}' />
                                                  <condition attribute='hil_rmatype' operator='eq' value='{_returntype.Id}' />
                                                  <condition attribute='hil_rma' operator='null' />
                                                </filter>
                                              </entity>
                                            </fetch>";
                EntityCollection entCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                if (entCol.Entities.Count > 0)
                {
                    int batchSize = 1000;

                    for (int i = 0; i < entCol.Entities.Count; i += batchSize)
                    {
                        var batch = entCol.Entities.Skip(i).Take(batchSize).ToList();

                        ExecuteMultipleRequest requestWithResults = new ExecuteMultipleRequest()
                        {
                            // Assign settings that define execution behavior: continue on error, return responses. 
                            Settings = new ExecuteMultipleSettings()
                            {
                                ContinueOnError = false,
                                ReturnResponses = true
                            },
                            // Create an empty organization request collection.
                            Requests = new OrganizationRequestCollection()
                        };
                        Entity InventoryJournal = null;
                        foreach (var entity in batch)
                        {
                            InventoryJournal = new Entity("hil_inventoryproductjournal", entity.Id);
                            InventoryJournal["hil_rma"] = new EntityReference("hil_inventoryrma", new Guid("450aaba0-da0e-f011-9989-7c1e523d5ba0"));
                            UpdateRequest updateRequest = new UpdateRequest() { Target = InventoryJournal };
                            requestWithResults.Requests.Add(updateRequest);
                        }
                        try
                        {
                            ExecuteMultipleResponse responseWithResults = (ExecuteMultipleResponse)service.Execute(requestWithResults);
                            //_statusCode = true;
                            //context.OutputParameters["StatusMessage"] = "Success - Inventory Journal RMA updated";
                        }
                        catch (Exception ex)
                        {
                            //context.OutputParameters["StatusMessage"] = $"ERROR! - {ex.Message}";
                        }
                    }
                }
                else
                {
                    //context.OutputParameters["StatusMessage"] = "NO Active Inventory Journal Lines Found";
                }
            }


            int maxRetries = 3;
            int retryCount = 0;
            int delayMilliseconds = 50;

            while (retryCount < maxRetries)
            {
                try
                {
                    var client = new HttpClient();
                    var request = new HttpRequestMessage(HttpMethod.Get, "https://middlewaredev.havells.com:50001/RESTAdapter/amazoncrm/getunackrequest");
                    request.Headers.Add("Authorization", "Basic RDM2NV9IQVZFTExTOkRFVkQzNjVAMTIzNA==");
                    request.Headers.Add("Cookie", "incap_ses_737_2920498=SPmnGqoZ6GTsmv7l21k6Cq8ryWcAAAAAzuZRc0KDFoSVqGfd3MzO6g==; visid_incap_2920498=gW0JJYrZQVSVq6G3pbQINoBXxWcAAAAAQUIPAAAAAAB1daFtI921b9lvbtfvYaQQ; JSESSIONID=5Q7DDy96MM-NbkKXLGNl1kkUptJplQGfnHkA_SAPJyCLOXqQA7l4ezq0NqppglEf; saplb_*=(J2EE7969920)7969950");

                    var response = await client.SendAsync(request);
                    response.EnsureSuccessStatusCode();

                    var responseBody = await response.Content.ReadAsStringAsync();
                    var jsonResponse = JObject.Parse(responseBody);

                    if (jsonResponse.ContainsKey("message"))
                    {
                        string message = jsonResponse.SelectToken("message").ToString();
                        if (message == "Rate Exceeded")
                        {
                            retryCount++;
                            if (retryCount < maxRetries)
                            {
                                Console.WriteLine($"Rate exceeded. Retrying in {delayMilliseconds}ms. Retry count: {retryCount}");
                                await Task.Delay(delayMilliseconds);
                                delayMilliseconds *= 3;
                                continue;
                            }
                            else
                            {
                                Console.WriteLine("Maximum retries reached. Rate exceeded. Aborting.");
                                return;
                            }
                        }
                    }
                    List<string> unacknowledgedIds = jsonResponse.SelectToken("ids").ToObject<List<string>>();
                    for (int i = 0; i < unacknowledgedIds.Count; i++)
                    {
                        await GetInstallationRequestData(unacknowledgedIds[i]);
                    }
                    break;
                }
                catch (HttpRequestException ex)
                {
                    retryCount++;
                    if (retryCount < maxRetries)
                    {
                        Console.WriteLine($"HttpRequestException: {ex.Message}. Retrying in {delayMilliseconds}ms. Retry count: {retryCount}");
                        await Task.Delay(delayMilliseconds);
                        delayMilliseconds *= 3;
                    }
                    else
                    {
                        Console.WriteLine($"Maximum retries reached. HttpRequestException: {ex.Message}. Aborting.");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"An unexpected error occurred: {ex.Message}");

                }
            }

            //var client = new HttpClient();
            //var request = new HttpRequestMessage(HttpMethod.Get, "https://middlewaredev.havells.com:50001/RESTAdapter/amazoncrm/getunackrequest");
            //request.Headers.Add("Authorization", "Basic RDM2NV9IQVZFTExTOkRFVkQzNjVAMTIzNA==");
            //request.Headers.Add("Cookie", "incap_ses_737_2920498=SPmnGqoZ6GTsmv7l21k6Cq8ryWcAAAAAzuZRc0KDFoSVqGfd3MzO6g==; visid_incap_2920498=gW0JJYrZQVSVq6G3pbQINoBXxWcAAAAAQUIPAAAAAAB1daFtI921b9lvbtfvYaQQ; JSESSIONID=5Q7DDy96MM-NbkKXLGNl1kkUptJplQGfnHkA_SAPJyCLOXqQA7l4ezq0NqppglEf; saplb_*=(J2EE7969920)7969950");
            //var response = await client.SendAsync(request);
            //response.EnsureSuccessStatusCode();
            //var responseBody = await response.Content.ReadAsStringAsync();
            //var jsonResponse = JObject.Parse(responseBody);

            //if (jsonResponse.ContainsKey("message"))
            //{
            //    string message = jsonResponse.SelectToken("message").ToString();
            //    if (message == "Rate Exceeded")
            //    {

            //    }
            //}
            //var ids = jsonResponse["ids"];
            // List<string> unacknowledgedIds = jsonResponse.SelectToken("ids").ToObject<List<string>>();

            //for (int i = 0; i < unacknowledgedIds.Count; i++)
            //{
            //    await GetInstallationRequestData(unacknowledgedIds[i]);
            //}

        }

        public async Task<TimeSpan> MeasureApiCall(string url)
        {
            var stopwatch = Stopwatch.StartNew();
            var client = new HttpClient()
            {
                Timeout = TimeSpan.FromSeconds(300)
            };
            var dnsLookupTime = TimeSpan.Zero;
            var tcpHandshakeTime = TimeSpan.Zero;
            var ttfbTime = TimeSpan.Zero;
            var requestStartTime = DateTime.UtcNow;
            var requestEndTime = DateTime.UtcNow;
            var responseStartTime = DateTime.UtcNow;
            var responseEndTime = DateTime.UtcNow;
            var responseBodyReadTime = TimeSpan.Zero;

            try
            {
                var uri = new Uri(url);
                var host = uri.Host;
                var port = uri.Port;

                Console.WriteLine($"Starting API call to: {url}");
                Console.WriteLine($"Request Time (UTC): {requestStartTime:O}");

                // DNS Lookup
                var dnsLookupStart = Stopwatch.StartNew();
                Console.WriteLine("Starting DNS Lookup...");
                var addresses = await System.Net.Dns.GetHostAddressesAsync(host);
                dnsLookupTime = dnsLookupStart.Elapsed;
                Console.WriteLine($"DNS Lookup completed in {dnsLookupTime.TotalMilliseconds} ms");

                // TCP Handshake
                var tcpHandshakeStart = Stopwatch.StartNew();
                Console.WriteLine("Starting TCP Handshake...");
                using (var tcpClient = new TcpClient())
                {
                    await tcpClient.ConnectAsync(addresses[0], port);
                    tcpHandshakeTime = tcpHandshakeStart.Elapsed;
                    Console.WriteLine($"TCP Handshake completed in {tcpHandshakeTime.TotalMilliseconds} ms");
                }

                // Time to First Byte (TTFB) and API response time.
                var ttfbStopwatch = Stopwatch.StartNew();
                Console.WriteLine("Creating HttpRequestMessage...");
                var request = new HttpRequestMessage(HttpMethod.Get, url);
                request.Headers.Add("Authorization", "Basic RDM2NV9IQVZFTExTOkRFVkQzNjVAMTIzNA==");
                request.Headers.Add("Cookie", "incap_ses_737_2920498=SPmnGqoZ6GTsmv7l21k6Cq8ryWcAAAAAzuZRc0KDFoSVqGfd3MzO6g==; visid_incap_2920498=gW0JJYrZQVSVq6G3pbQINoBXxWcAAAAAQUIPAAAAAAB1daFtI921b9lvbtfvYaQQ; JSESSIONID=5Q7DDy96MM-NbkKXLGNl1kkUptJplQGfnHkA_SAPJyCLOXqQA7l4ezq0NqppglEf; saplb_*=(J2EE7969920)7969950");
                Console.WriteLine("Sending HttpRequest...");
                responseStartTime = DateTime.UtcNow;
                var response = await client.SendAsync(request);
                responseEndTime = DateTime.UtcNow;
                ttfbTime = ttfbStopwatch.Elapsed;
                Console.WriteLine($"TTFB: {ttfbTime.TotalMilliseconds} ms");

                Console.WriteLine("Checking response status code...");
                response.EnsureSuccessStatusCode(); // throw on error status codes
                Console.WriteLine("Response status code OK");

                Console.WriteLine("Reading response body...");
                var bodyReadStart = Stopwatch.StartNew();
                await response.Content.ReadAsStringAsync(); // force the entire response to be read.
                responseBodyReadTime = bodyReadStart.Elapsed;
                Console.WriteLine($"Response body read in {responseBodyReadTime.TotalMilliseconds} ms");

            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
            finally
            {
                stopwatch.Stop();
                requestEndTime = DateTime.UtcNow;
                Console.WriteLine($"Request End Time (UTC): {requestEndTime:O}");
            }

            Console.WriteLine($"DNS Lookup: {dnsLookupTime.TotalMilliseconds} ms");
            Console.WriteLine($"TCP Handshake: {tcpHandshakeTime.TotalMilliseconds} ms");
            Console.WriteLine($"TTFB: {ttfbTime.TotalMilliseconds} ms");
            Console.WriteLine($"Response Body Read Time: {responseBodyReadTime.TotalMilliseconds} ms");
            if (stopwatch.Elapsed.TotalMilliseconds > 3000) Console.ForegroundColor = ConsoleColor.DarkRed;
            Console.WriteLine($"Total Time: {stopwatch.Elapsed.TotalMilliseconds} ms");
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine($"Response Start Time (UTC): {responseStartTime:O}");
            Console.WriteLine($"Response End Time (UTC): {responseEndTime:O}");

            return stopwatch.Elapsed;
        }

        internal async Task GetInstallationRequestData(string id)
        {
            int maxRetries = 3;
            int retryCount = 0;
            int delayMilliseconds = 50;

            
            while (retryCount < maxRetries)
            {
                try
                {
                    var client = new HttpClient();
                    var request = new HttpRequestMessage(HttpMethod.Get, $"https://middlewaredev.havells.com:50001/RESTAdapter/amazoncrm/getinstallationreqdata?request-ids={id}");
                    request.Headers.Add("Authorization", "Basic RDM2NV9IQVZFTExTOkRFVkQzNjVAMTIzNA==");
                    request.Headers.Add("Cookie", "incap_ses_737_2920498=BtaGZbX4tREC/R7m21k6CsxYyWcAAAAAjLcXKnlHxhUNhfWJ78/1lA==; visid_incap_2920498=gW0JJYrZQVSVq6G3pbQINoBXxWcAAAAAQUIPAAAAAAB1daFtI921b9lvbtfvYaQQ; JSESSIONID=5Q7DDy96MM-NbkKXLGNl1kkUptJplQGfnHkA_SAPJyCLOXqQA7l4ezq0NqppglEf; saplb_*=(J2EE7969920)7969951");
                    var response = await client.SendAsync(request);
                    response.EnsureSuccessStatusCode();
                    var responseBody = await response.Content.ReadAsStringAsync();
                    var jsonResponse = JObject.Parse(responseBody);

                    if (jsonResponse.ContainsKey("message"))
                    {
                        string message = jsonResponse.SelectToken("message").ToString();
                        if (message == "Rate Exceeded")
                        {
                            retryCount++;
                            if (retryCount < maxRetries)
                            {
                                Console.WriteLine($"Rate exceeded. Retrying in {delayMilliseconds}ms. Retry count: {retryCount}");
                                await Task.Delay(delayMilliseconds);
                                delayMilliseconds *= 3;
                                continue;
                            }
                            else
                            {
                                Console.WriteLine("Maximum retries reached. Rate exceeded. Aborting.");
                                return;
                            }
                        }
                    }

                    ProcessInstallationRequestResponse(jsonResponse);
                    break;

                }
                catch (HttpRequestException ex)
                {

                    retryCount++;
                    if (retryCount < maxRetries)
                    {
                        Console.WriteLine($"HttpRequestException: {ex.Message}. Retrying in {delayMilliseconds}ms. Retry count: {retryCount}");
                        await Task.Delay(delayMilliseconds);
                        delayMilliseconds *= 3;
                    }
                    else
                    {
                        Console.WriteLine($"Maximum retries reached. HttpRequestException: {ex.Message}. Aborting.");

                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"An unexpected error occurred: {ex.Message}");

                }
            }

        }

        private void ProcessInstallationRequestResponse(JObject jsonResponse)
        {

            // Check for failed request IDs
            if (jsonResponse.ContainsKey("failedRequestIds") && jsonResponse["failedRequestIds"].Any())
            {
                Console.WriteLine("Failed Request IDs:");
                foreach (var failedId in jsonResponse["failedRequestIds"])
                {
                    Console.WriteLine($"- {failedId}");
                }
            }
            else
            {
                Console.WriteLine("No failed request2 IDs.");
            }

            // Process line items
            if (jsonResponse.ContainsKey("lineItems"))
            {
                Console.WriteLine("\nLine Items:");
                var lineItems = jsonResponse["lineItems"].ToObject<Dictionary<string, JObject>>();
                foreach (var lineItemEntry in lineItems)
                {
                    var lineItemId = lineItemEntry.Key;
                    var lineItemData = lineItemEntry.Value;

                    Console.WriteLine($"\n  Line Item ID: {lineItemId}");

                    // Access customer information
                    if (lineItemData.ContainsKey("customer"))
                    {
                        Console.WriteLine("    Customer:");
                        Console.WriteLine($"      Name: {lineItemData["customer"]["name"]}");
                        Console.WriteLine($"      Email: {lineItemData["customer"]["email"]}");
                        Console.WriteLine($"      Phone: {lineItemData["customer"]["phoneNumber"]}");
                    }
                    // Access estimated delivery date
                    if (lineItemData.ContainsKey("estimatedDeliveryDate"))
                    {
                        long estimatedDeliveryDateMilliseconds = (long)lineItemData["estimatedDeliveryDate"];
                        DateTimeOffset dateTimeOffset = DateTimeOffset.FromUnixTimeMilliseconds(estimatedDeliveryDateMilliseconds);
                        Console.WriteLine($"    Estimated Delivery Date: {dateTimeOffset.UtcDateTime}");
                    }

                    // Access item information
                    if (lineItemData.ContainsKey("item"))
                    {
                        Console.WriteLine("    Item:");
                        Console.WriteLine($"      Brand: {lineItemData["item"]["brand"]}");
                        Console.WriteLine($"      Category: {lineItemData["item"]["category"]}");
                        Console.WriteLine($"      Model Number: {lineItemData["item"]["modelNumber"]}");
                        Console.WriteLine($"      Title: {lineItemData["item"]["title"]}");
                    }
                    // Access mailing address
                    if (lineItemData.ContainsKey("mailingAddress"))
                    {
                        Console.WriteLine("    Mailing Address:");
                        Console.WriteLine($"      Address: {lineItemData["mailingAddress"]["address"]}");
                        Console.WriteLine($"      City: {lineItemData["mailingAddress"]["city"]}");
                        Console.WriteLine($"   Postal Code: {lineItemData["mailingAddress"]["postalCode"]}");
                    }

                    // Access order ID
                    if (lineItemData.ContainsKey("orderId"))
                    {
                        Console.WriteLine($"    Order ID: {lineItemData["orderId"]}");
                    }
                }
            }
            else
            {
                Console.WriteLine("No line items found in the response.");
            }

            // Process requests
            if (jsonResponse.ContainsKey("requests"))
            {
                Console.WriteLine("\nRequests:");
                var requests = jsonResponse["requests"].ToObject<Dictionary<string, JObject>>();
                foreach (var requestEntry in requests)
                {
                    var requestId = requestEntry.Key;
                    var requestData = requestEntry.Value;

                    Console.WriteLine($"\n  Request ID: {requestId}");

                    // Access CRM ticket ID
                    if (requestData.ContainsKey("crmTicketId") && requestData["crmTicketId"] != null)
                    {
                        Console.WriteLine($"    CRM Ticket ID: {requestData["crmTicketId"]}");
                    }
                    else
                    {
                        Console.WriteLine("    CRM Ticket ID: Not available");
                    }

                    // Access ID
                    if (requestData.ContainsKey("id"))
                    {
                        Console.WriteLine($"    ID: {requestData["id"]}");
                    }

                    // Access line item ID
                    if (requestData.ContainsKey("lineItemId"))
                    {
                        Console.WriteLine($"    Line Item ID: {requestData["lineItemId"]}");
                    }

                    // Access status
                    if (requestData.ContainsKey("status"))
                    {
                        Console.WriteLine($"    Status: {requestData["status"]}");
                    }
                }
            }
            else
            {
                Console.WriteLine("No requests found in the response.");
            }

            Console.WriteLine("");
        }

        public class _403079933671507191
        {
            public object crmTicketId { get; set; }
            public string id { get; set; }
            public string lineItemId { get; set; }
            public string status { get; set; }
        }

        public class Customer
        {
            public string email { get; set; }
            public string name { get; set; }
            public string phoneNumber { get; set; }
        }

        public class Item
        {
            public string brand { get; set; }
            public string category { get; set; }
            public string modelNumber { get; set; }
            public string title { get; set; }
        }

        public class LineItemId1604896032
        {
            public Customer customer { get; set; }
            public long estimatedDeliveryDate { get; set; }
            public Item item { get; set; }
            public MailingAddress mailingAddress { get; set; }
            public string orderId { get; set; }
        }

        public class LineItems
        {
            [JsonProperty("LineItemId-1604896032")]
            public LineItemId1604896032 LineItemId1604896032 { get; set; }
        }

        public class MailingAddress
        {
            public string address { get; set; }
            public string city { get; set; }
            public string postalCode { get; set; }
        }

        public class Requests
        {
            [JsonProperty("403-0799336-7150719-1")]
            public _403079933671507191 _403079933671507191 { get; set; }
        }

        public class InstallationRequestData
        {
            public List<object> failedRequestIds { get; set; }
            public LineItems lineItems { get; set; }
            public Requests requests { get; set; }
        }
    }
}




//Fetch Unacknowledged installation request ids
//Method: GET

//URL: https://middlewaredev.havells.com:50001/RESTAdapter/amazoncrm/getunackrequest

//Payload: Refer attached Amazon doc

//Get Installation request data:
//Method: GET

//URL: https://middlewaredev.havells.com:50001/RESTAdapter/amazoncrm/getinstallationreqdata?request-ids=402-1497528-0337926-1

//Payload: Refer attached Amazon doc

//Acknowledge Requests:
//Method: POST

//URL: https://middlewaredev.havells.com:50001/RESTAdapter/amazoncrm/ackrequest

//Payload: Refer attached Amazon doc

//Update Requests:
//Method: POST

//URL: https://middlewaredev.havells.com:50001/RESTAdapter/amazoncrm/
//

