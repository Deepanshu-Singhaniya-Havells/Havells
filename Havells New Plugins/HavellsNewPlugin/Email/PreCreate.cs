using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace HavellsNewPlugin.Email
{
    public class PreCreate : IPlugin
    {
        public static ITracingService tracingService = null;
        string TenderNo = string.Empty;
        string CustomerName = string.Empty;
        string ProjectName = string.Empty;
        string _subject = string.Empty;
        string _emailSubject = string.Empty;
        Guid id = new Guid();
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
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity
                      && context.PrimaryEntityName.ToLower() == "email" && context.MessageName.ToUpper() == "CREATE")
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    if (entity.Contains("regardingobjectid"))
                    {
                        EntityReference Regarding = entity.GetAttributeValue<EntityReference>("regardingobjectid");

                        string message = entity.GetAttributeValue<String>("description");
                        if (message.Contains("{#clickHere#}"))
                        {
                            String URL = @"https://havells.crm8.dynamics.com/main.aspx?appid=675ffa44-b6d0-ea11-a813-000d3af05d7b&forceUCI=1&pagetype=entityrecord&etn=" + Regarding.LogicalName + "&id=" + Regarding.Id;
                            URL = "<a href="+URL+">Click Here</a>";
                            message.Replace("{#clickHere#}", URL);
                            entity["description"] = message;
                        }
                        
                        if (Regarding.LogicalName == "hil_tender")
                        {

                            if (entity.Contains("subject") && entity["subject"] != null)
                            {
                                _subject = (string)entity["subject"];
                                tracingService.Trace("subject:" + _subject);
                            }
                            if (entity.Contains("regardingobjectid") && entity["regardingobjectid"] != null)
                            {
                                Regarding = (EntityReference)entity["regardingobjectid"];
                                id = Regarding.Id;
                            }
                            Entity tenderEntity = service.Retrieve(Regarding.LogicalName, Regarding.Id, new ColumnSet("hil_customerprojectname", "hil_salesoffice",
                                "hil_name", "hil_customername"));
                            if (tenderEntity.Contains("hil_customerprojectname") && tenderEntity["hil_customerprojectname"] != null)
                            {
                                ProjectName = (string)tenderEntity["hil_customerprojectname"];
                            }
                            if (tenderEntity.Contains("hil_name") && tenderEntity["hil_name"] != null)
                            {
                                TenderNo = (string)tenderEntity["hil_name"];
                            }
                            if (tenderEntity.Contains("hil_customername") && tenderEntity["hil_customername"] != null)
                            {
                                EntityReference customer = (EntityReference)tenderEntity["hil_customername"];
                                Guid id = customer.Id;
                                CustomerName = customer.Name;
                            }
                            EntityReference SO = null;
                            string SOname = string.Empty;
                            if (tenderEntity.Contains("hil_salesoffice") && tenderEntity["hil_salesoffice"] != null)
                            {
                                SO = (EntityReference)tenderEntity["hil_salesoffice"];
                                SOname = SO.Name;
                            }

                            _emailSubject = TenderNo + "' SO: '" + SOname + "'" + " Client '" + CustomerName + "' Project: '" + ProjectName + "' SO: '" + SOname + "'";
                            tracingService.Trace("Email full subject:" + _emailSubject);
                            if (_subject.IndexOf(TenderNo) < 0)
                            {
                                entity["subject"] = _emailSubject + "-" + _subject;
                            }
                        }
                        else if (Regarding.LogicalName == "hil_orderchecklistproduct")
                        {
                            //Tender No<Tender No. if any> <OCLNo.> <Project Name if any><Customer Name> - In case if the same belongs to OMS
                            if (entity.Contains("subject") && entity["subject"] != null)
                            {
                                _subject = (string)entity["subject"];
                                tracingService.Trace("subject:" + _subject);
                            }
                            if (entity.Contains("regardingobjectid") && entity["regardingobjectid"] != null)
                            {
                                Regarding = (EntityReference)entity["regardingobjectid"];
                                id = Regarding.Id;
                            }
                            Entity tenderEntity = service.Retrieve(Regarding.LogicalName, Regarding.Id,
                                new ColumnSet("hil_tenderno", "hil_projectname", "hil_name", "hil_nameofclientcustomercode"));
                            if (tenderEntity.Contains("hil_tenderno") && tenderEntity["hil_tenderno"] != null)
                            {
                                _emailSubject = "Tender No " + tenderEntity.GetAttributeValue<EntityReference>("hil_tenderno").Name + " ";
                            }
                            if (tenderEntity.Contains("hil_name") && tenderEntity["hil_name"] != null)
                            {
                                _emailSubject = _emailSubject + tenderEntity.GetAttributeValue<string>("hil_name") + " ";
                            }
                            if (tenderEntity.Contains("hil_projectname") && tenderEntity["hil_projectname"] != null)
                            {
                                _emailSubject = _emailSubject + "Project " + tenderEntity.GetAttributeValue<string>("hil_projectname") + " ";
                            }
                            if (tenderEntity.Contains("hil_nameofclientcustomercode") && tenderEntity["hil_nameofclientcustomercode"] != null)
                            {
                                _emailSubject = _emailSubject + "customer " + tenderEntity.GetAttributeValue<EntityReference>("hil_nameofclientcustomercode").Name;
                            }
                            tracingService.Trace("Email full subject:" + _emailSubject);
                            entity["subject"] = _emailSubject + "-" + _subject;

                        }
                        else if (Regarding.LogicalName == "hil_oaheader")
                        {

                            //Tender No<Tender No. if any> <OCLNo.> <OA No.><Project Name if any><Customer Name> - In case if the same belongs to SMS
                            if (entity.Contains("subject") && entity["subject"] != null)
                            {
                                _subject = (string)entity["subject"];
                                tracingService.Trace("subject:" + _subject);
                            }
                            if (_subject.Contains("APPROVAL(Cable BU) FOR LIMIT ON SELF EXPOSURE"))
                            {
                                return;
                            }
                            if (entity.Contains("regardingobjectid") && entity["regardingobjectid"] != null)
                            {
                                Regarding = (EntityReference)entity["regardingobjectid"];
                                id = Regarding.Id;
                            }
                            Entity tenderEntity = service.Retrieve(Regarding.LogicalName, Regarding.Id,
                                new ColumnSet("hil_orderchecklistid", "hil_customername", "hil_name"));
                            if (tenderEntity.Contains("hil_orderchecklistid") && tenderEntity["hil_orderchecklistid"] != null)
                            {
                                EntityReference oclRef = tenderEntity.GetAttributeValue<EntityReference>("hil_orderchecklistid");
                                Entity OCL = service.Retrieve(oclRef.LogicalName, oclRef.Id, new ColumnSet("hil_tenderno", "hil_projectname"));
                                if (OCL.Contains("hil_tenderno") && OCL["hil_tenderno"] != null)
                                {
                                    _emailSubject = "Tender No " + OCL.GetAttributeValue<EntityReference>("hil_tenderno").Name + " ";
                                }
                                _emailSubject = oclRef.Name + " ";
                                if (tenderEntity.Contains("hil_name") && tenderEntity["hil_name"] != null)
                                {
                                    _emailSubject = "OA No. " + _emailSubject + tenderEntity.GetAttributeValue<string>("hil_name") + " ";
                                }
                                if (OCL.Contains("hil_projectname") && OCL["hil_projectname"] != null)
                                {
                                    _emailSubject = _emailSubject + "Project " + OCL.GetAttributeValue<string>("hil_projectname") + " ";
                                }
                            }
                            else if (tenderEntity.Contains("hil_name") && tenderEntity["hil_name"] != null)
                            {
                                _emailSubject = "OA No. " + _emailSubject + tenderEntity.GetAttributeValue<string>("hil_name") + " ";
                            }
                            if (tenderEntity.Contains("hil_customername") && tenderEntity["hil_customername"] != null)
                            {
                                _emailSubject = _emailSubject + "customer " + tenderEntity.GetAttributeValue<EntityReference>("hil_customername").Name;
                            }
                            tracingService.Trace("Email full subject:" + _emailSubject);
                            entity["subject"] = _emailSubject + "-" + _subject;
                            //throw new InvalidPluginExecutionException("test Message");
                            //if (_subject.Contains("Saurabh Tripathi"))
                            //{
                            //    throw new InvalidPluginExecutionException("test Message");
                            //}
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
    }
}
