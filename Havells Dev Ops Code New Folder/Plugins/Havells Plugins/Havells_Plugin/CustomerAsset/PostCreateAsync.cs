using System;
using System.IO;
using System.Net;
using System.Linq;
using System.Text;
using Newtonsoft.Json;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Havells_Plugin.HelperIntegration;

namespace Havells_Plugin.CustomerAsset
{
    public class PostCreateAsync : IPlugin
    {
        public static ITracingService tracingService = null;

        public void Execute(IServiceProvider serviceProvider)
        {
            #region PluginConfig
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion
            try
            {
                if (context.InputParameters.Contains("Target")
                    && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == "msdyn_customerasset"
                    && context.MessageName.ToUpper() == "CREATE")
                {
                    OrganizationServiceContext orgContext = new OrganizationServiceContext(service);
                    Entity entity = (Entity)context.InputParameters["Target"];
                    //Create child asset for assets having serial number
                    if (!entity.Contains("msdyn_parentasset"))
                    {
                       //GetChildAsset(entity, service);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.CustomerAsset.PostCreate.PostCreateAsync" + ex.Message);
            }
        }
        public static void GetChildAsset(Entity entity, IOrganizationService service)
        {
            try
            {
                msdyn_customerasset ParentCustomerAsset = entity.ToEntity<msdyn_customerasset>();//(msdyn_customerasset)entity;
                if (ParentCustomerAsset.msdyn_name != null && ParentCustomerAsset.statuscode.Value == 1)//active
                {
                    String sSerialNumber = ParentCustomerAsset.msdyn_name;
                    using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
                    {
                        #region Credentials
                        String sUserName = String.Empty;
                        String sPassword = String.Empty;
                        var obj2 = from _IConfig in orgContext.CreateQuery<hil_integrationconfiguration>()
                                   where _IConfig.hil_name == "Credentials"
                                   select new { _IConfig };
                        foreach (var iobj2 in obj2)
                        {
                            if (iobj2._IConfig.hil_Username != String.Empty)
                                sUserName = iobj2._IConfig.hil_Username;
                            if (iobj2._IConfig.hil_Password != String.Empty)
                                sPassword = iobj2._IConfig.hil_Password;
                        }
                        #endregion

                        var obj = from _IConfig in orgContext.CreateQuery<hil_integrationconfiguration>()
                                  where _IConfig.hil_name == "SerialNumberGetChildAssets"
                                  select new { _IConfig.hil_Url };
                        foreach (var iobj in obj)
                        {
                            if (iobj.hil_Url != null)
                            {
                                // Create a request using a URL that can receive a post.   
                                WebRequest request = WebRequest.Create(iobj.hil_Url);
                                string credentials = Convert.ToBase64String(Encoding.ASCII.GetBytes(sUserName + ":" + sPassword));
                                request.Headers[HttpRequestHeader.Authorization] = "Basic " + credentials;
                                //  request.Credentials = new NetworkCredential("D365_Havells", "QAD365@1234");
                                // Set the Method property of the request to POST.  
                                request.Method = "POST";
                                // Create POST data and convert it to a byte array.
                                string postData = "{\"IM_SERIAL_NUMBER\" : \"" + sSerialNumber + "\"}";
                                byte[] byteArray = Encoding.UTF8.GetBytes(postData);
                                // Set the ContentType property of the WebRequest.  
                                request.ContentType = "application/x-www-form-urlencoded";
                                // Set the ContentLength property of the WebRequest.  
                                request.ContentLength = byteArray.Length;
                                // Get the request stream.  
                                Stream dataStream = request.GetRequestStream();
                                // Write the data to the request stream.  
                                dataStream.Write(byteArray, 0, byteArray.Length);
                                // Close the Stream object.  
                                dataStream.Close();
                                // Get the response.  
                                WebResponse response = request.GetResponse();
                                // Display the status.  
                                Console.WriteLine(((HttpWebResponse)response).StatusDescription);
                                // Get the stream containing content returned by the server.  
                                dataStream = response.GetResponseStream();
                                // Open the stream using a StreamReader for easy access.  
                                StreamReader reader = new StreamReader(dataStream);
                                // Read the content.  
                                string responseFromServer = reader.ReadToEnd();

                                if (responseFromServer.Contains("Serial Number is Wrong"))
                                {
                                    //serial number is wrong
                                    //throw new InvalidPluginExecutionException("serial number is wrong");
                                }
                                else
                                {
                                    ChildAsset empObj = JsonConvert.DeserializeObject<ChildAsset>(responseFromServer);
                                    foreach (ET_SERIAL_DETAIL ChildAsset in empObj.ET_SERIAL_DETAIL)
                                    {
                                        String PartSerialNo = ChildAsset.PART_LABEL_ID;
                                        String PartMaterialCode = ChildAsset.PART_MATNR;
                                        Guid fsPorductCode = Guid.Empty;
                                        Guid fsMatGroup = Guid.Empty;
                                        Guid fsMatGroup1 = Guid.Empty;

                                        #region FindPartMaterialCode
                                        var obj1 = (from _Product in orgContext.CreateQuery<Product>()
                                                    where _Product.ProductNumber == PartSerialNo
                                                    select new
                                                    {
                                                        _Product.ProductId
                                                        ,
                                                        _Product.hil_MaterialGroup
                                                        ,
                                                        _Product.hil_MaterialGroup1
                                                    }).Take(1);
                                        foreach (var iobj1 in obj1)
                                        {
                                            fsPorductCode = iobj1.ProductId.Value;
                                            if (iobj1.hil_MaterialGroup != null)
                                                fsMatGroup = iobj1.hil_MaterialGroup.Id;
                                            if (iobj1.hil_MaterialGroup1 != null)
                                                fsMatGroup = iobj1.hil_MaterialGroup1.Id;
                                        }
                                        #endregion
                                        msdyn_customerasset crCreateChildAsset = new msdyn_customerasset();
                                        crCreateChildAsset.msdyn_ParentAsset = new EntityReference(msdyn_customerasset.EntityLogicalName, ParentCustomerAsset.msdyn_customerassetId.Value);
                                        if (ParentCustomerAsset.hil_Customer != null)
                                            crCreateChildAsset.hil_Customer = ParentCustomerAsset.hil_Customer;
                                        crCreateChildAsset.msdyn_name = PartSerialNo;
                                        if (fsPorductCode != Guid.Empty)
                                            crCreateChildAsset.msdyn_Product = new EntityReference(Product.EntityLogicalName, fsPorductCode);
                                        if (fsMatGroup != Guid.Empty)
                                            crCreateChildAsset.hil_ProductCategory = new EntityReference(Product.EntityLogicalName, fsMatGroup);
                                        if (fsMatGroup1 != Guid.Empty)
                                            crCreateChildAsset.hil_ProductSubcategory = new EntityReference(Product.EntityLogicalName, fsMatGroup1);
                                        service.Create(crCreateChildAsset);
                                    }
                                }
                                Console.WriteLine(responseFromServer);
                                // Clean up the streams.  
                                reader.Close();
                                dataStream.Close();
                                response.Close();

                            }
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.CustomerAsset.PostCreate.PostCreateAsync" + ex.Message);
            }
        }
    }
}
