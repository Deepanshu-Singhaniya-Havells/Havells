using System;
using System.Collections.Generic;
using System.Net;
using System.Text;
using System.Web.Script.Serialization;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using RestSharp;


namespace HavellsNewPlugin.TenderModule.OrderCheckList
{
    public class ExtractMarginContribution : IPlugin
    {
        public static ITracingService tracingService = null;
        public static Integration IntegrationConfiguration(IOrganizationService service, string Param)
        {
            Integration output = new Integration();
            try
            {
                QueryExpression qsCType = new QueryExpression("hil_integrationconfiguration");
                qsCType.ColumnSet = new ColumnSet("hil_url", "hil_username", "hil_password");
                qsCType.NoLock = true;
                qsCType.Criteria = new FilterExpression(LogicalOperator.And);
                qsCType.Criteria.AddCondition("hil_name", ConditionOperator.Equal, Param);
                Entity integrationConfiguration = service.RetrieveMultiple(qsCType)[0];
                output.uri = integrationConfiguration.GetAttributeValue<string>("hil_url");
                if (integrationConfiguration.Contains("hil_username") && integrationConfiguration.Contains("hil_password"))
                {
                    output.Auth = integrationConfiguration.GetAttributeValue<string>("hil_username") + ":" + integrationConfiguration.GetAttributeValue<string>("hil_password");
                }

            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error:- " + ex.Message);
            }
            return output;
        }
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
                if (context.InputParameters.Contains("Ids") && context.InputParameters["Ids"] is string && context.Depth == 1)
                {
                    //MarginContributionRes _objMargin = null;
                    Entity ent = new Entity();
                    Entity entOCL = new Entity();
                    string MATNR = "";
                    Decimal QUANTITY = 0;
                    string WERKS = "";
                    string KUNNR = "";
                    Decimal NET_VAL_PC = 0;
                    EntityReference ocl = null;
                    MarginContribution Reqdata = new MarginContribution();
                    List<REQ_INPUT> resInput = new List<REQ_INPUT>();
                    var OCLlineids = context.InputParameters["Ids"].ToString();
                    string ids = OCLlineids;
                    //string[] idsarray = ids.Split(';');
                    try
                    {

                        QueryExpression query1 = new QueryExpression("hil_orderchecklistproduct");
                        query1.ColumnSet = new ColumnSet("hil_product", "hil_porate", "hil_quantity", "hil_poqty", "hil_orderchecklistid", "hil_poamount", "hil_totalvalueinrs", "hil_tenderproductid");
                        query1.Criteria = new FilterExpression(LogicalOperator.And);
                        query1.Criteria.AddCondition("hil_orderchecklistid", ConditionOperator.Equal, new Guid(ids));
                        EntityCollection entOCLPROD = service.RetrieveMultiple(query1);

                        if (entOCLPROD.Entities.Count > 0)
                        {
                            foreach (Entity entprod in entOCLPROD.Entities)
                            {
                                ocl = entprod.GetAttributeValue<EntityReference>("hil_orderchecklistid");
                                entOCL = service.Retrieve(ocl.LogicalName, ocl.Id, new ColumnSet("hil_typeoforder", "hil_nameofclientcustomercode", "hil_despatchpoint", "hil_tenderno","hil_name"));
                                tracingService.Trace("1");

                                if (entprod.Attributes.Contains("hil_product"))
                                {
                                    tracingService.Trace("2");
                                    MATNR = entprod.GetAttributeValue<EntityReference>("hil_product").Name;
                                }
                                else
                                {
                                    throw new InvalidPluginExecutionException("Product Code Required");
                                }
                                if (entprod.Contains("hil_tenderproductid"))
                                {
                                    if (entprod.Contains("hil_porate"))
                                    {
                                        tracingService.Trace("2");
                                        NET_VAL_PC = Convert.ToDecimal(entprod.GetAttributeValue<Money>("hil_porate").Value);
                                    }
                                    else
                                    {
                                        throw new InvalidPluginExecutionException("PO Rate Required");
                                    }
                                }
                                else
                                {
                                    if (entprod.Attributes.Contains("hil_totalvalueinrs"))
                                    {
                                        tracingService.Trace("22");
                                        NET_VAL_PC = Convert.ToDecimal(entprod.GetAttributeValue<Money>("hil_totalvalueinrs").Value);
                                    }
                                    else
                                    {
                                        throw new InvalidPluginExecutionException("Total Value Required");
                                    }
                                }
                                if (entprod.Contains("hil_tenderproductid"))
                                {
                                    if (entprod.Contains("hil_poqty"))
                                    {
                                        tracingService.Trace("3");
                                        tracingService.Trace("ent.getattributeValue<decimal>(hil_poqty)  " + entprod.GetAttributeValue<Decimal>("hil_poqty"));
                                        QUANTITY = entprod.GetAttributeValue<Decimal>("hil_poqty");
                                        tracingService.Trace("QUANTITY " + QUANTITY.ToString());
                                    }
                                    else
                                    {
                                        throw new InvalidPluginExecutionException("Quantity Required");
                                    }
                                }
                                else
                                {
                                    if (entprod.Contains("hil_quantity"))
                                    {
                                        tracingService.Trace("33");
                                        QUANTITY = entprod.GetAttributeValue<Decimal>("hil_quantity");
                                        tracingService.Trace("QUANTITYy " + QUANTITY.ToString());
                                    }
                                    else
                                    {
                                        throw new InvalidPluginExecutionException("Quantity Required");
                                    }
                                }

                                if (entOCL.Attributes.Contains("hil_despatchpoint"))
                                {
                                    tracingService.Trace("4");
                                    WERKS = entOCL.GetAttributeValue<EntityReference>("hil_despatchpoint").Name;
                                }
                                else
                                {
                                    throw new InvalidPluginExecutionException("Plant Code Required");
                                }
                                if (entOCL.Attributes.Contains("hil_nameofclientcustomercode"))
                                {

                                    Entity customer = service.Retrieve(entOCL.GetAttributeValue<EntityReference>("hil_nameofclientcustomercode").LogicalName,
                                        entOCL.GetAttributeValue<EntityReference>("hil_nameofclientcustomercode").Id, new ColumnSet("accountnumber"));
                                    KUNNR = customer.GetAttributeValue<string>("accountnumber");
                                    tracingService.Trace("KUNNR" + KUNNR);
                                }
                                else
                                {
                                    throw new InvalidPluginExecutionException("Customer Code Required");
                                }
                                string ORDERTYPE = string.Empty;
                                if (entOCL.Attributes.Contains("hil_typeoforder"))
                                {
                                    ORDERTYPE = entOCL.FormattedValues["hil_typeoforder"];
                                    string[] ordertypearr = ORDERTYPE.Split('-');
                                    ORDERTYPE = ordertypearr[1];
                                }

                                resInput.Add(new REQ_INPUT
                                {
                                    MATNR = MATNR,
                                    QUANTITY = QUANTITY.ToString(),
                                    WERKS = WERKS,
                                    KUNNR = KUNNR,
                                    NET_VAL_PC = NET_VAL_PC.ToString(),
                                    ZORDERTYPE = ORDERTYPE
                                });

                            }
                        }
                        Reqdata.T_INPUT = resInput;
                        var Reqjson = JsonConvert.SerializeObject(resInput);
                        //tracingService.Trace("Reqdata " + Reqjson);


                        Integration integration = ExtractMarginContribution.IntegrationConfiguration(service, "MarginContribution");

                        //var client = new RestClient("https://middlewareqa.havells.com:50001/RESTAdapter/dynamics/MarginPercentage");
                        var client = new RestClient(integration.uri);
                        //client.Timeout = -1;
                        var request = new RestRequest();
                        #region logrequest
                        Entity intigrationTrace = new Entity("hil_integrationtrace");
                        intigrationTrace["hil_entityname"] = entOCL.LogicalName;
                        intigrationTrace["hil_entityid"] = entOCL.Id.ToString();
                        intigrationTrace["hil_request"] = JsonConvert.SerializeObject(Reqdata);
                        intigrationTrace["hil_name"] = entOCL.GetAttributeValue<string>("hil_name");
                        Guid intigrationTraceID = service.Create(intigrationTrace);
                        #endregion logrequest
                        tracingService.Trace("request " + JsonConvert.SerializeObject(Reqdata));

                        string _authInfo = Convert.ToBase64String(Encoding.ASCII.GetBytes(integration.Auth));//"D365_HAVELLS:QAD365@1234"
                        request.AddHeader("Authorization", "Basic " + _authInfo);
                        request.AddHeader("Content-Type", "application/json");
                        request.AddParameter("application/json", JsonConvert.SerializeObject(Reqdata), ParameterType.RequestBody);
                        RestResponse response = client.Execute(request, Method.Post);

                        #region logresponse
                        Entity intigrationTraceUp = new Entity("hil_integrationtrace");
                        intigrationTraceUp["hil_response"] = response.Content == "" ? response.ErrorMessage : response.Content;
                        intigrationTraceUp.Id = intigrationTraceID;
                        service.Update(intigrationTraceUp);
                        #endregion logresponse

                        tracingService.Trace("Response " + response.Content.ToString());
                        if (response.StatusCode == HttpStatusCode.OK)
                        {
                            MarginContributionRes MarRes = (new JavaScriptSerializer()).Deserialize<MarginContributionRes>(response.Content.ToString());

                            //dynamic MarRes = JsonConvert.DeserializeObject<MarginContributionRes>(response.Content);

                            if (MarRes.T_OUTPUT.Count > 0)
                            {

                                Entity OCLUPDATE = new Entity(ocl.LogicalName);
                                OCLUPDATE.Id = ocl.Id;
                                OCLUPDATE["hil_tot_contribution"] = Convert.ToDecimal(MarRes.EV_TOT_CONTRIBUTION);
                                OCLUPDATE["hil_tot_margin"] = Convert.ToDecimal(MarRes.EV_TOT_MARGIN);
                                service.Update(OCLUPDATE);
                                context.OutputParameters["TotMargin"] = Convert.ToDecimal(MarRes.EV_TOT_MARGIN);
                                context.OutputParameters["TotContribution"] = Convert.ToDecimal(MarRes.EV_TOT_CONTRIBUTION);
                                tracingService.Trace("TotMargin");
                                foreach (var item in MarRes.T_OUTPUT)
                                {
                                    string Material = item.MATNR;
                                    decimal MARGIN = Convert.ToDecimal(item.MARGIN);
                                    decimal CONTRIBUTION = Convert.ToDecimal(item.CONTRIBUTION);
                                    if (!string.IsNullOrEmpty(Material))
                                    {
                                        QueryExpression entProd = new QueryExpression("product");
                                        entProd.ColumnSet = new ColumnSet(false);
                                        entProd.Criteria = new FilterExpression(LogicalOperator.And);
                                        entProd.Criteria.AddCondition("name", ConditionOperator.Equal, Material);
                                        EntityCollection entityColl = service.RetrieveMultiple(entProd);

                                        QueryExpression query = new QueryExpression("hil_orderchecklistproduct");
                                        query.ColumnSet = new ColumnSet(false);
                                        query.Criteria = new FilterExpression(LogicalOperator.And);
                                        query.Criteria.AddCondition("hil_product", ConditionOperator.Equal, entityColl.Entities[0].Id);
                                        query.Criteria.AddCondition("hil_orderchecklistid", ConditionOperator.Equal, ocl.Id);
                                        EntityCollection entityCollProd = service.RetrieveMultiple(query);

                                        if (entityCollProd.Entities.Count > 0)
                                        {
                                            foreach (Entity enti in entityCollProd.Entities)
                                            {
                                                enti["hil_contribution"] = CONTRIBUTION;
                                                enti["hil_margin"] = MARGIN;
                                                service.Update(enti);
                                                tracingService.Trace("complete");
                                            }
                                        }
                                    }

                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        throw new InvalidPluginExecutionException("D365 Internal Server Error : " + ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Extract Margin Contribution Error :- " + ex.Message);
            }
        }
    }

    public class MarginContributionRes
    {
        public string EV_TOT_MARGIN { get; set; }
        public string RETURN { get; set; }
        public string EV_TOT_CONTRIBUTION { get; set; }
        public List<REQ_INPUT> T_OUTPUT { get; set; }
    }
    public class MarginContribution
    {
        public List<REQ_INPUT> T_INPUT { get; set; }
    }
    public class REQ_INPUT
    {
        public string MATNR { get; set; }
        public string QUANTITY { get; set; }
        public string WERKS { get; set; }
        public string KUNNR { get; set; }
        public string NET_VAL_PC { get; set; }
        public string MARGIN { get; set; }
        public string CONTRIBUTION { get; set; }
        public string ZORDERTYPE { get; set; }
    }
    public class Integration
    {
        public string uri { get; set; }
        public string Auth { get; set; }
    }

}
