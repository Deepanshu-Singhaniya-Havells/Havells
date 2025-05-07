using Havells_Plugin.HelperIntegration;
using HavellsConnection;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using RestSharp;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Globalization;
using System.Linq;
using System.Runtime.Remoting.Contexts;
using System.Text;
using System.Threading.Tasks;

namespace Havells.Crm.WebJob
{

    class Program
    {
        static void Main(string[] args)
        {

            string logicalNameOfBPF = "msdyn_bpf_2c5fe86acc8b414b8322ae571000c799";

            var connStr = "AuthType=ClientSecret;url={0};ClientId=41623af4-f2a7-400a-ad3a-ae87462ae44e;ClientSecret=r/qy6jgTisxeKZ7T3nYTVbhswXc5pYSVoSp5XFKJtQ8=";
            var CrmURL = "https://havellsfsm.crm8.dynamics.com";
            string finalString = string.Format(connStr, CrmURL);

            IOrganizationService service = HavellsConnection.CreateConnection.createConnection(finalString);


            Entity entity = service.Retrieve("msdyn_purchaseorderreceiptproduct", new Guid("b00cf0c1-aa98-ed11-aad1-6045bdac5c01"), new ColumnSet(true));
            ////tracingService.Trace("0");

            //entity = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet(true));
            decimal damageQty = 0;
            if (entity.Contains("hil_damagedreceivequantity"))
            {
                damageQty = entity.GetAttributeValue<decimal>("hil_damagedreceivequantity");
            }
            if (damageQty == 0)
            {
                return;
            }
            if (entity.Contains("msdyn_associatetowarehouse"))
            {
                string fetch = $@"<fetch version=""1.0"" output-format=""xml-platform"" mapping=""logical"" distinct=""true"">
                                  <entity name=""msdyn_warehouse"">
                                    <attribute name=""msdyn_name"" />
                                    <attribute name=""createdon"" />
                                    <attribute name=""msdyn_description"" />
                                    <attribute name=""msdyn_warehouseid"" />
                                    <order attribute=""msdyn_name"" descending=""false"" />
                                    <filter type=""and"">
                                      <condition attribute=""msdyn_warehouseid"" operator=""eq"" value=""{entity.GetAttributeValue<EntityReference>("msdyn_associatetowarehouse").Id}"" />
                                    </filter>
                                    <link-entity name=""bookableresource"" from=""msdyn_warehouse"" to=""msdyn_warehouseid"" link-type=""inner"" alias=""aa"">
                                      <attribute name=""hil_damagegoodswarehouse"" />
                                    </link-entity>
                                  </entity>
                                </fetch>";
                EntityCollection wareHouse = service.RetrieveMultiple(new FetchExpression(fetch));
                if (wareHouse.Entities.Count == 1)
                {
                    EntityReference warehouseDamage = (EntityReference)wareHouse[0].GetAttributeValue<AliasedValue>("aa.hil_damagegoodswarehouse").Value;
                    Entity entity1 = new Entity(entity.LogicalName);
                    entity1["msdyn_associatetowarehouse"] = warehouseDamage;
                    entity1["msdyn_quantity"] = (double)(damageQty);
                    //entity1["msdyn_name"] = entity["msdyn_name"];
                    if (entity.Contains("msdyn_purchaseorderreceipt"))
                        entity1["msdyn_purchaseorderreceipt"] = entity.GetAttributeValue<EntityReference>("msdyn_purchaseorderreceipt");

                    if (entity.Contains("msdyn_purchaseorder"))
                        entity1["msdyn_purchaseorder"] = entity.GetAttributeValue<EntityReference>("msdyn_purchaseorder");

                    if (entity.Contains("msdyn_purchaseorderproduct"))
                        entity1["msdyn_purchaseorderproduct"] = entity.GetAttributeValue<EntityReference>("msdyn_purchaseorderproduct");

                    entity1["transactioncurrencyid"] = entity.GetAttributeValue<EntityReference>("transactioncurrencyid");
                    service.Create(entity1);
                }
                else
                {
                    throw new InvalidPluginExecutionException("Damaged Warehouse is not configred " + wareHouse.Entities.Count);
                }
            }


            //EntityCollection collection = service.RetrieveMultiple(new FetchExpression(fetchXML));

            //Entity entity = service.Retrieve("msdyn_purchaseorder", new Guid("4ff90676-c905-45ac-b8d7-0a54dfee6de1"), new ColumnSet(true));

            Entity activeProcessInstance = GetActiveBPFDetails(entity, service);
            if (activeProcessInstance != null)
            {
                Guid activeBPFId = activeProcessInstance.Id; // Id of the active process instance, which will be used
                                                             // Retrieve the active stage ID of in the active process instance
                Guid activeStageId = new Guid(activeProcessInstance.Attributes["processstageid"].ToString());
                int currentStagePosition = -1;
                RetrieveActivePathResponse pathResp = GetAllStagesOfSelectedBPF(activeBPFId, activeStageId, ref currentStagePosition, service);

                string ActiveStageName = pathResp.ProcessStages.Entities[currentStagePosition].GetAttributeValue<string>("stagename");
                //CreateRecipt(entity.ToEntityReference(), service);
                //if (ActiveStageName == "Purchase Order Draft")
                {
                    if (currentStagePosition > -1 && pathResp.ProcessStages != null && pathResp.ProcessStages.Entities != null &&
                        currentStagePosition + 1 < pathResp.ProcessStages.Entities.Count)
                    {
                        // Retrieve the stage ID of the next stage that you want to set as active
                        Guid nextStageId = (Guid)pathResp.ProcessStages.Entities[currentStagePosition + 1].Attributes["processstageid"];
                        // Set the next stage as the active stage
                        Entity entBPF = new Entity(logicalNameOfBPF)
                        {
                            Id = activeBPFId
                        };
                        entBPF["activestageid"] = new EntityReference("processstage", nextStageId);
                        service.Update(entBPF);
                    }
                }
            }

        }

        static void CreateRecipt(EntityReference entityReference, IOrganizationService service)
        {
            Entity entity = new Entity("msdyn_purchaseorderreceipt");
            entity["msdyn_purchaseorder"] = entityReference;
            Guid guid = service.Create(entity);

            QueryExpression query = new QueryExpression("msdyn_purchaseorderproduct");
            query.ColumnSet = new ColumnSet(true);
            query.Criteria.AddCondition(new ConditionExpression("msdyn_purchaseorder", ConditionOperator.In, entityReference.Id));
            EntityCollection entityCollection = service.RetrieveMultiple(query);
            foreach (Entity entity1 in entityCollection.Entities)
            {
                Entity entity2 = new Entity("msdyn_purchaseorderreceiptproduct");
                entity2["msdyn_purchaseorder"] = entityReference;
                entity2["msdyn_purchaseorderreceipt"] = new EntityReference(entity.LogicalName, guid);
                entity2["msdyn_purchaseorderproduct"] = entity1.ToEntityReference();
                entity2["msdyn_associatetowarehouse"] = entity1.GetAttributeValue<EntityReference>("msdyn_associatetowarehouse");
                entity2["msdyn_quantity"] = entity1.GetAttributeValue<double>("msdyn_quantity");
                Guid dd = service.Create(entity2);
            }

        }
        static public Entity GetActiveBPFDetails(Entity entity, IOrganizationService crmService)
        {
            Entity activeProcessInstance = null;
            RetrieveProcessInstancesRequest entityBPFsRequest = new RetrieveProcessInstancesRequest
            {
                EntityId = entity.Id,
                EntityLogicalName = entity.LogicalName
            };
            RetrieveProcessInstancesResponse entityBPFsResponse = (RetrieveProcessInstancesResponse)crmService.Execute(entityBPFsRequest);
            if (entityBPFsResponse.Processes != null && entityBPFsResponse.Processes.Entities != null)
            {
                activeProcessInstance = entityBPFsResponse.Processes.Entities[0];
            }
            return activeProcessInstance;
        }
        static public RetrieveActivePathResponse GetAllStagesOfSelectedBPF(Guid activeBPFId, Guid activeStageId, ref int currentStagePosition, IOrganizationService crmService)
        {
            // Retrieve the process stages in the active path of the current process instance
            RetrieveActivePathRequest pathReq = new RetrieveActivePathRequest
            {
                ProcessInstanceId = activeBPFId
            };
            RetrieveActivePathResponse pathResp = (RetrieveActivePathResponse)crmService.Execute(pathReq);
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
        public static IOrganizationService CreateCRMConnection()
        {
            var CrmURL = "https://havellsService.crm8.dynamics.com";// "https://orga23838be.crm11.dynamics.com/";
            var ClientId = "41623af4-f2a7-400a-ad3a-ae87462ae44e";// "bc6676d6-1387-4dc4-be89-ba13b08ceb4e";
            var ClientSecret = "r/qy6jgTisxeKZ7T3nYTVbhswXc5pYSVoSp5XFKJtQ8=";// "73P7Q~sWxupzl4j8-B55y5g3QNosxhkjkV6Q2";
            string finalString = string.Format("AuthType=ClientSecret;url={0};ClientId={1};ClientSecret={2}", CrmURL, ClientId, ClientSecret);
            IOrganizationService service = HavellsConnection.CreateConnection.createConnection(finalString);
            return service;
        }
        static void IsJobLate(IOrganizationService service, Entity job)
        {
            string query = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                <entity name='ogre_jobtypestatus'>
                    <attribute name='ogre_jobtypestatusid' />
                    <attribute name='ogre_name' />
                    <order attribute='ogre_name' descending='false' />
                    <filter type='and'>
                        <condition attribute='ogre_jobtype' operator='eq' uiname='Delivery - Field' uitype='ogre_jobtype' value='{job.GetAttributeValue<EntityReference>("ogre_jobtype").Id}' />
                        < condition attribute='ogre_jobstatus' operator='eq' uiname='Open' uitype='ogre_status' value='{job.GetAttributeValue<EntityReference>("ogre_statusid").Id}' />
                    </filter>
                    <link-entity name='ogre_jobrunningstatusrule' from='ogre_jobrunningstatusruleid' to='ogre_jobrule' link-type='inner' alias='ab'>
                        <attribute name='ogre_startdatetimeparameter' />
                        <attribute name='ogre_operator' />
                        <attribute name='ogre_enddatetimeparameter' />
                        <attribute name='ogre_enddatetimegracemin' />
                        <filter type='and'>
                        <condition attribute='ogre_ruletype' operator='eq' value='1' />
                        </filter>
                        <link-entity name='ogre_jobrunningstatusruleparameter' from='ogre_jobrunningstatusruleparameterid' to='ogre_startdatetimeparameter' link-type='inner' alias='ag'>
                        <attribute name='ogre_parameterschemaname' />
                        </link-entity>
                        <link-entity name='ogre_jobrunningstatusruleparameter' from='ogre_jobrunningstatusruleparameterid' to='ogre_enddatetimeparameter' link-type='inner' alias='ah'>
                        <attribute name='ogre_parameterschemaname' />
                        </link-entity>
                    </link-entity>
                </entity>
            </fetch>";
            EntityCollection entityCollection = service.RetrieveMultiple((QueryBase)new FetchExpression(query));
            if (entityCollection.Entities.Count <= 0)
                return;
            Entity entity = entityCollection[0];
            string attributeLogicalName1 = entity.GetAttributeValue<AliasedValue>("ah.ogre_parameterschemaname").Value.ToString();
            string attributeLogicalName2 = entity.GetAttributeValue<AliasedValue>("ag.ogre_parameterschemaname").Value.ToString();
            int num = ((OptionSetValue)entity.GetAttributeValue<AliasedValue>("ab.ogre_operator").Value).Value;
            string s = entity.GetAttributeValue<AliasedValue>("ab.ogre_enddatetimegracemin").Value.ToString();
            DateTime dateTime1 = attributeLogicalName2 == "Now()" ? DateTime.Now : job.GetAttributeValue<DateTime>(attributeLogicalName2);
            DateTime dateTime2 = attributeLogicalName1 == "Now()" ? DateTime.Now : job.GetAttributeValue<DateTime>(attributeLogicalName1).AddMinutes((double)int.Parse(s));
            bool flag = false;
            switch (num)
            {
                case 1:
                    if (dateTime1 == dateTime2)
                    {
                        flag = true;
                        break;
                    }
                    break;
                case 2:
                    if (dateTime1 > dateTime2)
                    {
                        flag = true;
                        break;
                    }
                    break;
                case 3:
                    if (dateTime1 < dateTime2)
                    {
                        flag = true;
                        break;
                    }
                    break;
                case 4:
                    if (dateTime1 >= dateTime2)
                    {
                        flag = true;
                        break;
                    }
                    break;
                case 5:
                    if (dateTime1 <= dateTime2)
                    {
                        flag = true;
                        break;
                    }
                    break;
            }

            Entity job1 = new Entity(job.LogicalName, job.Id);
            job1["ogre_isjoblate"] = flag;
            service.Update(job1);
        }


        static void cancleJob(IOrganizationService service, string[] jobId)
        {
            try
            {
                QueryExpression query = new QueryExpression("msdyn_workorder");
                query.ColumnSet = new ColumnSet("msdyn_name");
                query.Criteria.AddCondition(new ConditionExpression("msdyn_name", ConditionOperator.In, jobId));
                EntityCollection entityCollection = service.RetrieveMultiple(query);
                foreach (Entity entity in entityCollection.Entities)
                {
                    try
                    {
                        Console.WriteLine("Job Updating " + entity["msdyn_name"]);
                        entity["msdyn_substatus"] = new EntityReference("msdyn_workordersubstatus", new Guid("1527fa6c-fa0f-e911-a94e-000d3af060a1"));
                        entity["hil_closureremarks"] = "Cancelled as per approval - FW: IER_Havells_KKG_Dashboard May'22";
                        entity["hil_jobcancelreason"] = new OptionSetValue(4);
                        service.Update(entity);
                        Console.WriteLine("Job Updated");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Job Updating " + entity["msdyn_name"] + " Error " + ex.Message);
                    }

                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Job Updating Error " + ex.Message);
            }
        }
        public static void createPriceListItem(IOrganizationService service)
        {
            try
            {
                Entity pricelistItem = new Entity("productpricelevel");
                pricelistItem["productid"] = new EntityReference("product", new Guid("55454d05-2ffc-eb11-94ef-6045bd72ead7"));
                pricelistItem["amount"] = new Money(9130);
                pricelistItem["pricelevelid"] = new EntityReference("pricelevel", new Guid("06CBCC30-19AC-EC11-9840-6045BDAD4704"));
                pricelistItem["uomid"] = new EntityReference("uom", new Guid("41179f41-080c-e911-a94e-000d3af06091"));
                pricelistItem["transactioncurrencyid"] = new EntityReference("transactioncurrency", new Guid("68a6a9ca-6beb-e811-a96c-000d3af05828"));
                pricelistItem["quantitysellingcode"] = new OptionSetValue(1);
                service.Create(pricelistItem);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error " + ex.Message);
            }
        }
        public static ITracingService tracingService = null;
        public string _primaryField = string.Empty;
        protected static void Execute(IOrganizationService service)
        {
            try
            {
                EntityCollection entCCList = new EntityCollection();
                EntityCollection entToList = new EntityCollection();
                EntityReferenceCollection materialGroup = new EntityReferenceCollection();

                EntityCollection toTeamMembers = new EntityCollection();
                EntityCollection ccTeamMembers = new EntityCollection();
                EntityReference department = null;
                EntityReference ocl = null;

                EntityReference fromRef = new EntityReference("queue", new Guid("b6f0037d-3e93-ec11-b400-6045bdaad0b5"));
                EntityReference regardingRef = new EntityReference("hil_oaheader", new Guid("b0e962b5-1aa4-ec11-9840-6045bdac66ea"));
                string to = "t.CMT;";
                string cc = "p.BU Head;p.zonal Head;p.zonal representor;p.Branch Product Head;";
                string mailBody = "Limit requested by cable Team is here by approved";
                string mailsubject = @"<div data-wrapper='true' style='font - size:9pt; font - family:'Segoe UI','Helvetica Neue',sans - serif; '><div>Dear CMT Team,<br>&nbsp;<br>Limit requested by Cable Team is hereby approved for ₹&nbsp; {Limit Requested(OA Header)}  , in addition to the existing limit of Customer Name  {Customer Name(OA Header)} , Customer Code {Customer Code(OA Header)} .<br>\nPlz,&nbsp;<a href='#URL' target='_blank'>Click Here</a> for more details.&nbsp;<br>\n&nbsp;<br>\nRegards<br>\nAnil Rai Gupta</div></div>";


                Entity _regardingEntity = service.Retrieve(regardingRef.LogicalName, regardingRef.Id,
                    new ColumnSet("hil_department", "ownerid", "hil_orderchecklistid", "hil_zdc", "hil_scm", "hil_zonerepresentor", "hil_salesoffice"));
                if (_regardingEntity.Contains("hil_department"))
                {
                    department = _regardingEntity.GetAttributeValue<EntityReference>("hil_department");
                    Console.WriteLine("Department is " + department.Name);
                }
                if (_regardingEntity.Contains("hil_orderchecklistid"))
                {
                    ocl = _regardingEntity.GetAttributeValue<EntityReference>("hil_orderchecklistid");
                    Console.WriteLine("ocl is " + ocl.Name);
                }

                Entity _oclEntity = service.Retrieve(ocl.LogicalName, ocl.Id, new ColumnSet("hil_department", "hil_despatchpoint", "hil_buhead", "ownerid", "hil_rm", "hil_zonalhead", "hil_zdc", "hil_scm", "hil_zonerepresentor", "hil_salesoffice"));

                EntityReference plant = _oclEntity.GetAttributeValue<EntityReference>("hil_despatchpoint");

                QueryExpression queryOA = new QueryExpression("hil_oaproduct");
                queryOA.ColumnSet = new ColumnSet("hil_materialgroup");
                queryOA.Criteria = new FilterExpression(LogicalOperator.And);
                queryOA.Criteria.AddCondition(new ConditionExpression("hil_oaheader", ConditionOperator.Equal, _regardingEntity.Id));
                EntityCollection oaproductColl = service.RetrieveMultiple(queryOA);
                foreach (Entity oaPrd in oaproductColl.Entities)
                {
                    string material = oaPrd.GetAttributeValue<string>("hil_materialgroup");

                    QueryExpression queryMat = new QueryExpression("hil_materialgroup");
                    queryMat.ColumnSet = new ColumnSet(false);
                    queryMat.Criteria = new FilterExpression(LogicalOperator.And);
                    queryMat.Criteria.AddCondition(new ConditionExpression("hil_code", ConditionOperator.Equal, material));
                    EntityCollection matColl = service.RetrieveMultiple(queryMat);
                    if (matColl.Entities.Count > 0)
                    {
                        materialGroup.Add(matColl[0].ToEntityReference());
                    }
                }

                string[] toTeam = to.Split(';');
                string[] ccTeam = cc.Split(';');

                Console.WriteLine("toTeam count" + toTeam.Length);

                foreach (string totype in toTeam)
                {
                    if (totype.Contains("t.") || totype.Contains("T."))
                    {
                        string teamName = (totype.Replace("t.", "")).Replace("T.", "");
                        Console.WriteLine("team name" + teamName);
                        toTeamMembers = retriveTeamMembers(service, teamName, materialGroup, department, plant, toTeamMembers);
                    }
                    else if (totype.Contains("p.") || totype.Contains("P."))
                    {
                        string position = (totype.Replace("p.", "")).Replace("P.", "");
                        Console.WriteLine("position name" + position);
                        EntityReference positionRef = null;
                        if (position.ToLower() == "Zonal Head".ToLower())
                        {
                            positionRef = _oclEntity.GetAttributeValue<EntityReference>("hil_zonalhead");
                        }
                        else if (position.ToLower() == "scm".ToLower())
                        {
                            positionRef = _oclEntity.GetAttributeValue<EntityReference>("hil_scm");
                        }
                        else if (position.ToLower() == "BU Head".ToLower())
                        {
                            positionRef = _oclEntity.GetAttributeValue<EntityReference>("hil_buhead");
                        }
                        else if (position.ToLower() == "Branch Product Head".ToLower())
                        {
                            positionRef = _oclEntity.GetAttributeValue<EntityReference>("hil_rm");
                        }
                        else if (position.ToLower() == "Enquiry Creator".ToLower())
                        {
                            positionRef = _oclEntity.GetAttributeValue<EntityReference>("ownerid");
                        }
                        else if (position.ToLower() == "zonal representor".ToLower())
                        {
                            positionRef = _oclEntity.GetAttributeValue<EntityReference>("hil_zonerepresentor");
                        }
                        else if (position.ToLower() == "Design Team".ToLower())
                        {
                            EntityReference salsesOffice = _oclEntity.GetAttributeValue<EntityReference>("hil_salesoffice");
                            QueryExpression query = new QueryExpression("hil_designteambranchmapping");
                            query.ColumnSet = new ColumnSet("hil_user");
                            query.Criteria = new FilterExpression(LogicalOperator.And);
                            query.Criteria.AddCondition(new ConditionExpression("hil_salesoffice", ConditionOperator.Equal, salsesOffice.Id));
                            // query.Criteria.AddCondition(new ConditionExpression("statuscode", ConditionOperator.Equal, 0));
                            EntityCollection design = service.RetrieveMultiple(query);
                            if (design.Entities.Count > 0)
                                positionRef = design[0].GetAttributeValue<EntityReference>("hil_user");
                        }
                        Entity entity = service.Retrieve(positionRef.LogicalName, positionRef.Id, new ColumnSet(false));
                        toTeamMembers.Entities.Add(entity);
                    }
                }
                Console.WriteLine("ccTeam count" + ccTeam.Length);
                foreach (string totype in ccTeam)
                {
                    if (totype.Contains("t.") || totype.Contains("T."))
                    {
                        string teamName = (totype.Replace("t.", "")).Replace("T.", "");
                        Console.WriteLine("team name" + teamName);
                        ccTeamMembers = retriveTeamMembers(service, teamName, materialGroup, department, plant, toTeamMembers);
                    }
                    else if (totype.Contains("p.") || totype.Contains("P."))
                    {
                        string position = (totype.Replace("p.", "")).Replace("P.", "");
                        Console.WriteLine("position name" + position);
                        EntityReference positionRef = null;
                        if (position.ToLower() == "Zonal Head".ToLower())
                        {
                            positionRef = _oclEntity.GetAttributeValue<EntityReference>("hil_zonalhead");
                        }
                        else if (position.ToLower() == "scm".ToLower())
                        {
                            positionRef = _oclEntity.GetAttributeValue<EntityReference>("hil_scm");
                        }
                        else if (position.ToLower() == "BU Head".ToLower())
                        {
                            positionRef = _oclEntity.GetAttributeValue<EntityReference>("hil_buhead");
                        }
                        else if (position.ToLower() == "Branch Product Head".ToLower())
                        {
                            positionRef = _oclEntity.GetAttributeValue<EntityReference>("hil_rm");
                        }
                        else if (position.ToLower() == "Enquiry Creator".ToLower())
                        {
                            positionRef = _oclEntity.GetAttributeValue<EntityReference>("ownerid");
                        }
                        else if (position.ToLower() == "zonal representor".ToLower())
                        {
                            positionRef = _oclEntity.GetAttributeValue<EntityReference>("hil_zonerepresentor");
                        }
                        else if (position.ToLower() == "Design Team".ToLower())
                        {
                            EntityReference salsesOffice = _oclEntity.GetAttributeValue<EntityReference>("hil_salesoffice");
                            QueryExpression query = new QueryExpression("hil_designteambranchmapping");
                            query.ColumnSet = new ColumnSet("hil_user");
                            query.Criteria = new FilterExpression(LogicalOperator.And);
                            query.Criteria.AddCondition(new ConditionExpression("hil_salesoffice", ConditionOperator.Equal, salsesOffice.Id));
                            EntityCollection design = service.RetrieveMultiple(query);
                            if (design.Entities.Count > 0)
                                positionRef = design[0].GetAttributeValue<EntityReference>("hil_user");
                        }

                        Entity entity = service.Retrieve(positionRef.LogicalName, positionRef.Id, new ColumnSet(false));
                        ccTeamMembers.Entities.Add(entity);
                    }
                }
                String URL = @"https://havells.crm8.dynamics.com/main.aspx?appid=675ffa44-b6d0-ea11-a813-000d3af05d7b&forceUCI=1&pagetype=entityrecord&etn=";
                string recordURL = URL + _regardingEntity.LogicalName + "&id=" + _regardingEntity.Id;
                if (mailBody.Contains("#URL"))
                {
                    mailBody = mailBody.Replace("#URL", recordURL);
                }
                Console.WriteLine("toTeamMembers count" + toTeamMembers.Entities.Count);

                foreach (Entity ccEntity in toTeamMembers.Entities)
                {
                    Entity entCC = new Entity("activityparty");
                    entCC["partyid"] = ccEntity.ToEntityReference();
                    entToList.Entities.Add(entCC);
                }
                Console.WriteLine("entCCList count" + entCCList.Entities.Count);
                foreach (Entity ccEntity in ccTeamMembers.Entities)
                {
                    Entity entTo = new Entity("activityparty");
                    entTo["partyid"] = ccEntity.ToEntityReference();
                    entCCList.Entities.Add(entTo);
                }
                Entity entFrom = new Entity("activityparty");
                entFrom["partyid"] = new EntityReference(fromRef.LogicalName, fromRef.Id);
                Entity[] entFromList = { entFrom };

                Console.WriteLine("1");

                Entity email = new Entity("email");
                email["from"] = entFromList;
                email["to"] = entToList;
                email["cc"] = entCCList;
                email["description"] = mailBody;
                email["subject"] = mailsubject;
                email["regardingobjectid"] = _regardingEntity.ToEntityReference();
                Guid emailId = service.Create(email);

                SendEmailRequest sendEmailReq = new SendEmailRequest()
                {
                    EmailId = emailId,
                    IssueSend = true
                };
                SendEmailResponse sendEmailRes = (SendEmailResponse)service.Execute(sendEmailReq);

            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error :- " + ex.Message);// "HavellsNewPlugin.TenderModule.OrderCheckList.OrderCheckListPostCreate.Execute Error " + ex.Message);
            }
        }
        static public EntityCollection retriveTeamMembers(IOrganizationService service, string _teamName, EntityReferenceCollection _materialGroup, EntityReference _department,
              EntityReference _plant, EntityCollection extTeamMembers)
        {

            try
            {
                List<Guid> materialGuids = new List<Guid>();
                foreach (EntityReference entityReference in _materialGroup)
                {
                    materialGuids.Add(entityReference.Id);
                }
                QueryExpression _query = new QueryExpression("hil_bdteam");
                _query.ColumnSet = new ColumnSet("hil_name", "hil_materialgroup", "hil_department", "hil_plant");
                _query.Criteria = new FilterExpression(LogicalOperator.And);
                if (_teamName != null)
                    _query.Criteria.AddCondition(new ConditionExpression("hil_name", ConditionOperator.Equal, _teamName));
                if (_department != null)
                    _query.Criteria.AddCondition(new ConditionExpression("hil_department", ConditionOperator.Equal, _department.Id));
                EntityCollection bdteamCol = service.RetrieveMultiple(_query);
                if (bdteamCol.Entities.Count > 0)
                {
                    _query = new QueryExpression("hil_bdteam");
                    _query.ColumnSet = new ColumnSet(false);
                    _query.Criteria = new FilterExpression(LogicalOperator.And);
                    if (_teamName != null && bdteamCol[0].Contains("hil_name"))
                        _query.Criteria.AddCondition(new ConditionExpression("hil_name", ConditionOperator.Equal, _teamName));
                    if (materialGuids.Count > 0 && bdteamCol[0].Contains("hil_materialgroup"))
                        _query.Criteria.AddCondition(new ConditionExpression("hil_materialgroup", ConditionOperator.In, materialGuids.ToArray()));
                    if (_department != null && bdteamCol[0].Contains("hil_department"))
                        _query.Criteria.AddCondition(new ConditionExpression("hil_department", ConditionOperator.Equal, _department.Id));
                    if (_plant != null && bdteamCol[0].Contains("hil_plant"))
                        _query.Criteria.AddCondition(new ConditionExpression("hil_plant", ConditionOperator.Equal, _plant.Id));
                    //_query.Criteria.AddCondition(new ConditionExpression("statuscode", ConditionOperator.Equal, 0));
                    bdteamCol = service.RetrieveMultiple(_query);
                    if (bdteamCol.Entities.Count > 0)
                    {
                        Console.WriteLine("bdteamCol count " + bdteamCol.Entities.Count);
                        QueryExpression _querymem = new QueryExpression("hil_bdteammember");
                        _querymem.ColumnSet = new ColumnSet("emailaddress");
                        _querymem.Criteria = new FilterExpression(LogicalOperator.And);
                        _querymem.Criteria.AddCondition(new ConditionExpression("hil_team", ConditionOperator.Equal, bdteamCol.Entities[0].Id));
                        //_querymem.Criteria.AddCondition(new ConditionExpression("statuscode", ConditionOperator.Equal, 0));
                        EntityCollection bdteammemCol = service.RetrieveMultiple(_querymem);
                        EntityCollection entTOList = new EntityCollection();
                        Console.WriteLine("Team Members count" + entTOList.Entities.Count);
                        if (bdteammemCol.Entities.Count > 0)
                        {
                            foreach (Entity entity in bdteammemCol.Entities)
                            {
                                extTeamMembers.Entities.Add(entity);
                            }
                        }
                    }
                }



            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error in Retriving Team Members : " + ex.Message);
            }
            return extTeamMembers;
        }
        public string emailSubjectForOCL(IOrganizationService service, EntityReference regarding)
        {
            string _emailSubject = null;
            Entity tenderEntity = service.Retrieve(regarding.LogicalName, regarding.Id,
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
            Console.WriteLine("Email full subject:" + _emailSubject);
            return _emailSubject;

        }

        static void CloseExistingTatLine(int fromStage, int toStage, EntityReference regardingRef, IOrganizationService service)
        {
            try
            {
                QueryExpression query = new QueryExpression("hil_salestatline");
                query.ColumnSet = new ColumnSet("actualstart");
                query.Criteria = new FilterExpression(LogicalOperator.And);
                query.Criteria.AddCondition("regardingobjectid", ConditionOperator.Equal, regardingRef.Id);
                //query.Criteria.AddCondition("hil_fromstage", ConditionOperator.Equal, new OptionSetValue(fromStage));
                //query.Criteria.AddCondition("hil_tostage", ConditionOperator.Equal, new OptionSetValue(toStage));
                query.Criteria.AddCondition("statuscode", ConditionOperator.NotIn, new object[] { 2, 3 });
                EntityCollection Found = service.RetrieveMultiple(query);
                if (Found.Entities.Count > 0)
                {
                    DateTime start = Found[0].GetAttributeValue<DateTime>("actualstart");
                    DateTime now = DateTime.Now;
                    TimeSpan ts = start - now;
                    int duration = int.Parse(ts.TotalMinutes.ToString());
                    Entity entity = new Entity(Found[0].LogicalName, Found[0].Id);
                    entity["actualend"] = now;
                    entity["actualdurationminutes"] = duration;
                    service.Update(entity);
                    SetStateRequest setStateRequest = new SetStateRequest()
                    {
                        EntityMoniker = new EntityReference
                        {
                            Id = entity.Id,
                            LogicalName = entity.LogicalName,
                        },
                        State = new OptionSetValue(1),
                        Status = new OptionSetValue(2)
                    };
                    service.Execute(setStateRequest);
                }
            }
            catch (Exception ex)
            {

                throw new InvalidPluginExecutionException("Error on Close Existing TAT Line " + ex.Message);
            }

        }
        static void createTATLine(int fromStage, int toStage, EntityReference regardingRef, EntityReference department, IOrganizationService service)
        {
            try
            {
                QueryExpression query = new QueryExpression("hil_salestatmaster");
                query.ColumnSet = new ColumnSet("hil_durationmin", "hil_name");
                query.Criteria = new FilterExpression(LogicalOperator.And);
                query.Criteria.AddCondition("hil_department", ConditionOperator.Equal, department.Id);
                query.Criteria.AddCondition("hil_fromstage", ConditionOperator.Equal, new OptionSetValue(fromStage));
                query.Criteria.AddCondition("hil_tostage", ConditionOperator.Equal, new OptionSetValue(toStage));
                query.Criteria.AddCondition("statuscode", ConditionOperator.Equal, new OptionSetValue(0));
                EntityCollection Found = service.RetrieveMultiple(query);
                if (Found.Entities.Count > 0)
                {
                    int durationmin = Found[0].Contains("hil_durationmin") ? Found[0].GetAttributeValue<int>("hil_durationmin") : throw new InvalidPluginExecutionException("***** Duration Not Defin for this type *****"); ;
                    DateTime startTime = DateTime.Now;
                    DateTime endTime = DateTime.Now.AddMinutes(durationmin);
                    Entity entity = new Entity("hil_salestatline");
                    entity["subject"] = Found[0].Contains("hil_name") ? Found[0].GetAttributeValue<string>("hil_name") : throw new InvalidPluginExecutionException("***** name Not Defin for this type *****"); ;
                    entity["hil_fromstage"] = new OptionSetValue(fromStage);
                    entity["hil_tostage"] = new OptionSetValue(toStage);
                    entity["scheduleddurationminutes"] = durationmin;
                    entity["scheduledstart"] = startTime;
                    entity["scheduledend"] = endTime;
                    entity["actualstart"] = DateTime.Now;
                    entity["regardingobjectid"] = regardingRef;
                    service.Create(entity);
                }
                else
                {
                    throw new InvalidPluginExecutionException("***** Master Not Defin for this type *****");
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error on Create TAT Line " + ex.Message);
            }
        }

        static void upadateMubileNumber(IOrganizationService service)
        {

            string guids = "F97828ED-480D-E911-A94F-000D3AF063A5;86F31BC6-4B0D-E911-A94F-000D3AF063A5;58A507A5-490D-E911-A94F-000D3AF063A5;F70F9B43-490D-E911-A94D-000D3AF03089;1426F8F7-490D-E911-A950-000D3AF06A16;D5F41BC6-4B0D-E911-A94F-000D3AF063A5;706B8585-4B0D-E911-A94D-000D3AF03089;C631AD80-4B0D-E911-A94F-000D3AF06923;B902ADAF-480D-E911-A94E-000D3AF060A1;BA02ADAF-480D-E911-A94E-000D3AF060A1;C8B4109F-490D-E911-A94F-000D3AF063A5;28F874C8-4A0D-E911-A94E-000D3AF060A1;BD78AC1A-4A0D-E911-A94F-000D3AF06923;E0B19BF4-490D-E911-A94F-000D3AF00F43;793E3CB5-480D-E911-A94F-000D3AF06923;46E6DC61-490D-E911-A94F-000D3AF06923;2B3E3CB5-480D-E911-A94F-000D3AF06923;73088E7F-4B0D-E911-A94D-000D3AF03089;34B39BF4-490D-E911-A94F-000D3AF00F43;0C839E16-4B0D-E911-A94F-000D3AF063A5;D725D460-4B0D-E911-A950-000D3AF06A16;A7674278-4B0D-E911-A94F-000D3AF00F43;55694278-4B0D-E911-A94F-000D3AF00F43;9B07DBB0-4A0D-E911-A950-000D3AF06A16;F326D460-4B0D-E911-A950-000D3AF06A16;78A507A5-490D-E911-A94F-000D3AF063A5;30A4E750-490D-E911-A94E-000D3AF060A1;29698585-4B0D-E911-A94D-000D3AF03089;2FC98157-4A0D-E911-A94F-000D3AF063A5;92C8B114-4A0D-E911-A94F-000D3AF06923;43F51BC6-4B0D-E911-A94F-000D3AF063A5;62244B1F-4C0D-E911-A950-000D3AF06A16;B3E4DC61-490D-E911-A94F-000D3AF06923;3E04ADAF-480D-E911-A94E-000D3AF060A1;DEC8B114-4A0D-E911-A94F-000D3AF06923;F8F5CAAA-4A0D-E911-A950-000D3AF06A16;C3B59BF4-490D-E911-A94F-000D3AF00F43;5D121BCC-4B0D-E911-A94F-000D3AF063A5;EC842336-4C0D-E911-A94F-000D3AF06923;DCE4DC61-490D-E911-A94F-000D3AF06923;6007DBB0-4A0D-E911-A950-000D3AF06A16;2E60F14A-490D-E911-A94E-000D3AF060A1;6F58A23D-490D-E911-A94D-000D3AF03089;70D63BD1-4A0D-E911-A94F-000D3AF06923;800694FA-490D-E911-A94F-000D3AF00F43;DB59A23D-490D-E911-A94D-000D3AF03089;73B49BF4-490D-E911-A94F-000D3AF00F43;5704ADAF-480D-E911-A94E-000D3AF060A1;B2E6DC61-490D-E911-A94F-000D3AF06923;BE9F2C30-4C0D-E911-A94F-000D3AF06923;2E59A23D-490D-E911-A94D-000D3AF03089;9526D460-4B0D-E911-A950-000D3AF06A16;8D746BA2-4B0D-E911-A94E-000D3AF060A1;CE1F785D-4A0D-E911-A94F-000D3AF063A5;4A60E832-490D-E911-A94F-000D3AF00F43;8A73F4FE-490D-E911-A94D-000D3AF03089;8B7D28ED-480D-E911-A94F-000D3AF063A5;15BDE7F1-490D-E911-A950-000D3AF06A16;040794FA-490D-E911-A94F-000D3AF00F43;6417EFFD-490D-E911-A950-000D3AF06A16;EEF874C8-4A0D-E911-A94E-000D3AF060A1;932EAD80-4B0D-E911-A94F-000D3AF06923;8CA260B6-4A0D-E911-A94F-000D3AF00F43;552BF8F7-490D-E911-A950-000D3AF06A16;B2214B1F-4C0D-E911-A950-000D3AF06A16;C3684278-4B0D-E911-A94F-000D3AF00F43;54654278-4B0D-E911-A94F-000D3AF00F43;83FA74C8-4A0D-E911-A94E-000D3AF060A1;FA7F140A-4A0D-E911-A94E-000D3AF060A1;C6234B1F-4C0D-E911-A950-000D3AF06A16;F5234B1F-4C0D-E911-A950-000D3AF06A16;F20AD567-490D-E911-A94F-000D3AF06923;9FF41BC6-4B0D-E911-A94F-000D3AF063A5;B26A4278-4B0D-E911-A94F-000D3AF00F43;53C98157-4A0D-E911-A94F-000D3AF063A5;CA5AA23D-490D-E911-A94D-000D3AF03089;54F51BC6-4B0D-E911-A94F-000D3AF063A5;F423D460-4B0D-E911-A950-000D3AF06A16;6CF974C8-4A0D-E911-A94E-000D3AF060A1;DE05ADAF-480D-E911-A94E-000D3AF060A1;E5829E16-4B0D-E911-A94F-000D3AF063A5;8B736BA2-4B0D-E911-A94E-000D3AF060A1;1075F4FE-490D-E911-A94D-000D3AF03089;A26A8585-4B0D-E911-A94D-000D3AF03089;A7F4CAAA-4A0D-E911-A950-000D3AF06A16;267BAC1A-4A0D-E911-A94F-000D3AF06923;94577DC2-4A0D-E911-A94E-000D3AF060A1;2B62E832-490D-E911-A94F-000D3AF00F43;86C78157-4A0D-E911-A94F-000D3AF063A5;CC5EF14A-490D-E911-A94E-000D3AF060A1;1678F4FE-490D-E911-A94D-000D3AF03089;CEA160B6-4A0D-E911-A94F-000D3AF00F43;68D4961C-4B0D-E911-A94F-000D3AF063A5;0E7B28ED-480D-E911-A94F-000D3AF063A5;3E736BA2-4B0D-E911-A94E-000D3AF060A1;DA7F140A-4A0D-E911-A94E-000D3AF060A1;1CE5DC61-490D-E911-A94F-000D3AF06923;5FB43A7E-4B0D-E911-A94F-000D3AF00F43;AF839E16-4B0D-E911-A94F-000D3AF063A5;13F4A4C1-4A0D-E911-A94D-000D3AF03089;3F24D460-4B0D-E911-A950-000D3AF06A16;D1C21C04-4A0D-E911-A94E-000D3AF060A1;90131BCC-4B0D-E911-A94F-000D3AF063A5;1878AC1A-4A0D-E911-A94F-000D3AF06923;A0A160B6-4A0D-E911-A94F-000D3AF00F43;EA746BA2-4B0D-E911-A94E-000D3AF060A1;40C98157-4A0D-E911-A94F-000D3AF063A5;C60BD567-490D-E911-A94F-000D3AF06923;C85AA23D-490D-E911-A94D-000D3AF03089;D20AD567-490D-E911-A94F-000D3AF06923;2E0C45AF-480D-E911-A94F-000D3AF06923;5B5C43CB-4A0D-E911-A94F-000D3AF06923;1D098E7F-4B0D-E911-A94D-000D3AF03089;C26FEB04-4A0D-E911-A94D-000D3AF03089;26654278-4B0D-E911-A94F-000D3AF00F43;36A02C30-4C0D-E911-A94F-000D3AF06923;9C0F9B43-490D-E911-A94D-000D3AF03089;31F41BC6-4B0D-E911-A94F-000D3AF063A5;431F785D-4A0D-E911-A94F-000D3AF063A5;8A7BAC1A-4A0D-E911-A94F-000D3AF06923;FB6B8585-4B0D-E911-A94D-000D3AF03089;F2B6109F-490D-E911-A94F-000D3AF063A5;39F2A4C1-4A0D-E911-A94D-000D3AF03089;E4CAB114-4A0D-E911-A94F-000D3AF06923;6377F4FE-490D-E911-A94D-000D3AF03089;A3F31BC6-4B0D-E911-A94F-000D3AF063A5;89F3A4C1-4A0D-E911-A94D-000D3AF03089;655EE832-490D-E911-A94F-000D3AF00F43;9EA3E750-490D-E911-A94E-000D3AF060A1;B220D460-4B0D-E911-A950-000D3AF06A16;9E839E16-4B0D-E911-A94F-000D3AF063A5;9AA660B6-4A0D-E911-A94F-000D3AF00F43;9620D460-4B0D-E911-A950-000D3AF06A16;3B20785D-4A0D-E911-A94F-000D3AF063A5;6E842336-4C0D-E911-A94F-000D3AF06923;8D6A4278-4B0D-E911-A94F-000D3AF00F43;EA098E7F-4B0D-E911-A94D-000D3AF03089;FD34B67A-4B0D-E911-A94F-000D3AF06923;55BDE7F1-490D-E911-A950-000D3AF06A16;C0F4A4C1-4A0D-E911-A94D-000D3AF03089;986A4278-4B0D-E911-A94F-000D3AF00F43;6A5A43CB-4A0D-E911-A94F-000D3AF06923;6D9F2C30-4C0D-E911-A94F-000D3AF06923;B05EF14A-490D-E911-A94E-000D3AF060A1;EA6EEB04-4A0D-E911-A94D-000D3AF03089;9561E832-490D-E911-A94F-000D3AF00F43;5D25D460-4B0D-E911-A950-000D3AF06A16;EC9F2C30-4C0D-E911-A94F-000D3AF06923;EFF31BC6-4B0D-E911-A94F-000D3AF063A5;3460E832-490D-E911-A94F-000D3AF00F43;C26C8585-4B0D-E911-A94D-000D3AF03089;1B80140A-4A0D-E911-A94E-000D3AF060A1;3E0CD567-490D-E911-A94F-000D3AF06923;5E61E832-490D-E911-A94F-000D3AF00F43;A2B39BF4-490D-E911-A94F-000D3AF00F43;6A31AD80-4B0D-E911-A94F-000D3AF06923;33A760B6-4A0D-E911-A94F-000D3AF00F43;7E73F4FE-490D-E911-A94D-000D3AF03089;9AC8B114-4A0D-E911-A94F-000D3AF06923;4E59A23D-490D-E911-A94D-000D3AF03089;6C3E3CB5-480D-E911-A94F-000D3AF06923;216A8585-4B0D-E911-A94D-000D3AF03089;9B6EEB04-4A0D-E911-A94D-000D3AF03089;C96E6BA2-4B0D-E911-A94E-000D3AF060A1;0F08DBB0-4A0D-E911-A950-000D3AF06A16;31C88157-4A0D-E911-A94F-000D3AF063A5;5130AD80-4B0D-E911-A94F-000D3AF06923;A6706BA2-4B0D-E911-A94E-000D3AF060A1;45F974C8-4A0D-E911-A94E-000D3AF060A1;5A7C28ED-480D-E911-A94F-000D3AF063A5;B1B6109F-490D-E911-A94F-000D3AF063A5;E2567DC2-4A0D-E911-A94E-000D3AF060A1;9D1E785D-4A0D-E911-A94F-000D3AF063A5;DC33AD80-4B0D-E911-A94F-000D3AF06923;B46C8585-4B0D-E911-A94D-000D3AF03089;22BDE7F1-490D-E911-A950-000D3AF06A16;5679AC1A-4A0D-E911-A94F-000D3AF06923;20E5DC61-490D-E911-A94F-000D3AF06923;DEF41BC6-4B0D-E911-A94F-000D3AF063A5;75BCE7F1-490D-E911-A950-000D3AF06A16;9B6FEB04-4A0D-E911-A94D-000D3AF03089;C208DBB0-4A0D-E911-A950-000D3AF06A16;F70694FA-490D-E911-A94F-000D3AF00F43;076A4278-4B0D-E911-A94F-000D3AF00F43;68F4CAAA-4A0D-E911-A950-000D3AF06A16;E0A360B6-4A0D-E911-A94F-000D3AF00F43;FA684278-4B0D-E911-A94F-000D3AF00F43;8B7A28ED-480D-E911-A94F-000D3AF063A5;FCC65219-4C0D-E911-A950-000D3AF06A16;36F4CAAA-4A0D-E911-A950-000D3AF06A16;7407DBB0-4A0D-E911-A950-000D3AF06A16;1EC75219-4C0D-E911-A950-000D3AF06A16;EDFA74C8-4A0D-E911-A94E-000D3AF060A1;B10794FA-490D-E911-A94F-000D3AF00F43;0C80140A-4A0D-E911-A94E-000D3AF060A1;B5C8B114-4A0D-E911-A94F-000D3AF06923;5C5B43CB-4A0D-E911-A94F-000D3AF06923;216C8585-4B0D-E911-A94D-000D3AF03089;2BC41C04-4A0D-E911-A94E-000D3AF060A1;6EC78157-4A0D-E911-A94F-000D3AF063A5;FCCAB114-4A0D-E911-A94F-000D3AF06923;4D04ADAF-480D-E911-A94E-000D3AF060A1;035A43CB-4A0D-E911-A94F-000D3AF06923;E85AA23D-490D-E911-A94D-000D3AF03089;9D694278-4B0D-E911-A94F-000D3AF00F43;5F684278-4B0D-E911-A94F-000D3AF00F43;3A23D460-4B0D-E911-A950-000D3AF06A16;1B859E16-4B0D-E911-A94F-000D3AF063A5;9DA507A5-490D-E911-A94F-000D3AF063A5;52C8B114-4A0D-E911-A94F-000D3AF06923;00587DC2-4A0D-E911-A94E-000D3AF060A1;B22047C5-4A0D-E911-A94F-000D3AF06923;9ED63BD1-4A0D-E911-A94F-000D3AF06923;A5A360B6-4A0D-E911-A94F-000D3AF00F43;1127D460-4B0D-E911-A950-000D3AF06A16;60B43A7E-4B0D-E911-A94F-000D3AF00F43;FFA360B6-4A0D-E911-A94F-000D3AF00F43;3BF5A4C1-4A0D-E911-A94D-000D3AF03089;0925D460-4B0D-E911-A950-000D3AF06A16;47746BA2-4B0D-E911-A94E-000D3AF060A1;262DAD80-4B0D-E911-A94F-000D3AF06923;C459A23D-490D-E911-A94D-000D3AF03089;FB77F4FE-490D-E911-A94D-000D3AF03089;153F3CB5-480D-E911-A94F-000D3AF06923;D4098E7F-4B0D-E911-A94D-000D3AF03089;DF736BA2-4B0D-E911-A94E-000D3AF060A1;6DA360B6-4A0D-E911-A94F-000D3AF00F43;F7736BA2-4B0D-E911-A94E-000D3AF060A1;741E785D-4A0D-E911-A94F-000D3AF063A5;9AE6DC61-490D-E911-A94F-000D3AF06923;3CB43A7E-4B0D-E911-A94F-000D3AF00F43;E0C85219-4C0D-E911-A950-000D3AF06A16;5DF41BC6-4B0D-E911-A94F-000D3AF063A5;B2587DC2-4A0D-E911-A94E-000D3AF060A1;1233AD80-4B0D-E911-A94F-000D3AF06923;7031AD80-4B0D-E911-A94F-000D3AF06923;446A4278-4B0D-E911-A94F-000D3AF00F43;3F684278-4B0D-E911-A94F-000D3AF00F43;5D109B43-490D-E911-A94D-000D3AF03089;440894FA-490D-E911-A94F-000D3AF00F43;C680140A-4A0D-E911-A94E-000D3AF060A1;05869E16-4B0D-E911-A94F-000D3AF063A5;B9F3CAAA-4A0D-E911-A950-000D3AF06A16;48684278-4B0D-E911-A94F-000D3AF00F43;5B73F4FE-490D-E911-A94D-000D3AF03089;CF5A43CB-4A0D-E911-A94F-000D3AF06923;F9F2A4C1-4A0D-E911-A94D-000D3AF03089;96F4A4C1-4A0D-E911-A94D-000D3AF03089;C65FE832-490D-E911-A94F-000D3AF00F43;B4B6109F-490D-E911-A94F-000D3AF063A5;645C43CB-4A0D-E911-A94F-000D3AF06923;14698585-4B0D-E911-A94D-000D3AF03089;28A460B6-4A0D-E911-A94F-000D3AF00F43;A05FF14A-490D-E911-A94E-000D3AF060A1;63C88157-4A0D-E911-A94F-000D3AF063A5;DCC75219-4C0D-E911-A950-000D3AF06A16;F658A23D-490D-E911-A94D-000D3AF03089;30849E16-4B0D-E911-A94F-000D3AF063A5;2E63E832-490D-E911-A94F-000D3AF00F43;DDF2A4C1-4A0D-E911-A94D-000D3AF03089;145C43CB-4A0D-E911-A94F-000D3AF06923;C06E6BA2-4B0D-E911-A94E-000D3AF060A1;C9694278-4B0D-E911-A94F-000D3AF00F43;0480140A-4A0D-E911-A94E-000D3AF060A1;CC0E45AF-480D-E911-A94F-000D3AF06923;B8694278-4B0D-E911-A94F-000D3AF00F43;3618EFFD-490D-E911-A950-000D3AF06A16;4A5BA23D-490D-E911-A94D-000D3AF03089;4CC78157-4A0D-E911-A94F-000D3AF063A5;EA254B1F-4C0D-E911-A950-000D3AF06A16;4AC21C04-4A0D-E911-A94E-000D3AF060A1;1EF6A4C1-4A0D-E911-A94D-000D3AF03089;2FF51BC6-4B0D-E911-A94F-000D3AF063A5;B303ADAF-480D-E911-A94E-000D3AF060A1;70244B1F-4C0D-E911-A950-000D3AF06A16;58F51BC6-4B0D-E911-A94F-000D3AF063A5;6877F4FE-490D-E911-A94D-000D3AF03089;5C0CD567-490D-E911-A94F-000D3AF06923;D073F4FE-490D-E911-A94D-000D3AF03089;CB644278-4B0D-E911-A94F-000D3AF00F43;477A28ED-480D-E911-A94F-000D3AF063A5;19F6A4C1-4A0D-E911-A94D-000D3AF03089;B8B1109F-490D-E911-A94F-000D3AF063A5;D80794FA-490D-E911-A94F-000D3AF00F43;FBA6E750-490D-E911-A94E-000D3AF060A1;2D716BA2-4B0D-E911-A94E-000D3AF060A1;0C3C3CB5-480D-E911-A94F-000D3AF06923;0A839E16-4B0D-E911-A94F-000D3AF063A5;880794FA-490D-E911-A94F-000D3AF00F43;E0FA74C8-4A0D-E911-A94E-000D3AF060A1;47A260B6-4A0D-E911-A94F-000D3AF00F43;E9F674C8-4A0D-E911-A94E-000D3AF060A1;9CC8B114-4A0D-E911-A94F-000D3AF06923;267B28ED-480D-E911-A94F-000D3AF063A5;40A560B6-4A0D-E911-A94F-000D3AF00F43;8C7928ED-480D-E911-A94F-000D3AF063A5;0B7A28ED-480D-E911-A94F-000D3AF063A5;0A5943CB-4A0D-E911-A94F-000D3AF06923;4B098E7F-4B0D-E911-A94D-000D3AF03089;43F3A4C1-4A0D-E911-A94D-000D3AF03089;2536B67A-4B0D-E911-A94F-000D3AF06923;A1F5A4C1-4A0D-E911-A94D-000D3AF03089;87A360B6-4A0D-E911-A94F-000D3AF00F43;CA4F9749-490D-E911-A94D-000D3AF03089;FEF3CAAA-4A0D-E911-A950-000D3AF06A16;BE16EFFD-490D-E911-A950-000D3AF06A16;E4C65219-4C0D-E911-A950-000D3AF06A16;AD09DBB0-4A0D-E911-A950-000D3AF06A16;49C95219-4C0D-E911-A950-000D3AF06A16;189F5242-490D-E911-A950-000D3AF06A16;296C8585-4B0D-E911-A94D-000D3AF03089;51F974C8-4A0D-E911-A94E-000D3AF060A1;A63D3CB5-480D-E911-A94F-000D3AF06923;4D09DBB0-4A0D-E911-A950-000D3AF06A16;316FEB04-4A0D-E911-A94D-000D3AF03089;4878AC1A-4A0D-E911-A94F-000D3AF06923;469F5242-490D-E911-A950-000D3AF06A16;0F6A4278-4B0D-E911-A94F-000D3AF00F43;6C6EEB04-4A0D-E911-A94D-000D3AF03089;30C95219-4C0D-E911-A950-000D3AF06A16;5224D460-4B0D-E911-A950-000D3AF06A16;83F3CAAA-4A0D-E911-A950-000D3AF06A16;4F75F4FE-490D-E911-A94D-000D3AF03089;7D6FEB04-4A0D-E911-A94D-000D3AF03089;E82BF8F7-490D-E911-A950-000D3AF06A16;A8B6109F-490D-E911-A94F-000D3AF063A5;A8D63BD1-4A0D-E911-A94F-000D3AF06923;35684278-4B0D-E911-A94F-000D3AF00F43;BC3E3CB5-480D-E911-A94F-000D3AF06923;1FCBB114-4A0D-E911-A94F-000D3AF06923;A76A4278-4B0D-E911-A94F-000D3AF00F43;A920D460-4B0D-E911-A950-000D3AF06A16;7360E832-490D-E911-A94F-000D3AF00F43;96E5DC61-490D-E911-A94F-000D3AF06923;E278AC1A-4A0D-E911-A94F-000D3AF06923;3718EFFD-490D-E911-A950-000D3AF06A16;092EAD80-4B0D-E911-A94F-000D3AF06923;D5B19BF4-490D-E911-A94F-000D3AF00F43;B8C9B114-4A0D-E911-A94F-000D3AF06923;43CAB114-4A0D-E911-A94F-000D3AF06923;7C76F4FE-490D-E911-A94D-000D3AF03089;22A760B6-4A0D-E911-A94F-000D3AF00F43;7B22D460-4B0D-E911-A950-000D3AF06A16;4B21D460-4B0D-E911-A950-000D3AF06A16;33C95219-4C0D-E911-A950-000D3AF06A16;455BA23D-490D-E911-A94D-000D3AF03089;F99F2C30-4C0D-E911-A94F-000D3AF06923;8E80140A-4A0D-E911-A94E-000D3AF060A1;40852336-4C0D-E911-A94F-000D3AF06923;B2C78157-4A0D-E911-A94F-000D3AF063A5;A8684278-4B0D-E911-A94F-000D3AF00F43;AA0BD567-490D-E911-A94F-000D3AF06923;F81E785D-4A0D-E911-A94F-000D3AF063A5;195C43CB-4A0D-E911-A94F-000D3AF06923;27C21C04-4A0D-E911-A94E-000D3AF060A1;BA6EEB04-4A0D-E911-A94D-000D3AF03089;0F22D460-4B0D-E911-A950-000D3AF06A16;0859A23D-490D-E911-A94D-000D3AF03089;90015A3C-490D-E911-A950-000D3AF06A16;4731AD80-4B0D-E911-A94F-000D3AF06923;3AB43A7E-4B0D-E911-A94F-000D3AF00F43;905AA23D-490D-E911-A94D-000D3AF03089;06C75219-4C0D-E911-A950-000D3AF06A16;312DAD80-4B0D-E911-A94F-000D3AF06923;857F140A-4A0D-E911-A94E-000D3AF060A1;4426D460-4B0D-E911-A950-000D3AF06A16;3A684278-4B0D-E911-A94F-000D3AF00F43;F90794FA-490D-E911-A94F-000D3AF00F43;C65B43CB-4A0D-E911-A94F-000D3AF06923;B7746BA2-4B0D-E911-A94E-000D3AF060A1;B3842336-4C0D-E911-A94F-000D3AF06923;5076F4FE-490D-E911-A94D-000D3AF03089;E009DBB0-4A0D-E911-A950-000D3AF06A16;9521D460-4B0D-E911-A950-000D3AF06A16;996A4278-4B0D-E911-A94F-000D3AF00F43;7A36B67A-4B0D-E911-A94F-000D3AF06923;1E5EF14A-490D-E911-A94E-000D3AF060A1;45C95219-4C0D-E911-A950-000D3AF06A16;0717EFFD-490D-E911-A950-000D3AF06A16;B31F785D-4A0D-E911-A94F-000D3AF063A5;AD9F2C30-4C0D-E911-A94F-000D3AF06923;39756BA2-4B0D-E911-A94E-000D3AF060A1;D780140A-4A0D-E911-A94E-000D3AF060A1;18A4E750-490D-E911-A94E-000D3AF060A1;0CF774C8-4A0D-E911-A94E-000D3AF060A1;18F4CAAA-4A0D-E911-A950-000D3AF06A16;EA015A3C-490D-E911-A950-000D3AF06A16;96A507A5-490D-E911-A94F-000D3AF063A5;E0587DC2-4A0D-E911-A94E-000D3AF060A1;745AA23D-490D-E911-A94D-000D3AF03089;1A63E832-490D-E911-A94F-000D3AF00F43;ED34B67A-4B0D-E911-A94F-000D3AF06923;D3A160B6-4A0D-E911-A94F-000D3AF00F43;A14F9749-490D-E911-A94D-000D3AF03089;C9A8E750-490D-E911-A94E-000D3AF060A1;8D08DBB0-4A0D-E911-A950-000D3AF06A16;2180140A-4A0D-E911-A94E-000D3AF060A1;670894FA-490D-E911-A94F-000D3AF00F43;76F31BC6-4B0D-E911-A94F-000D3AF063A5;EC859E16-4B0D-E911-A94F-000D3AF063A5;4D1F785D-4A0D-E911-A94F-000D3AF063A5;8B20785D-4A0D-E911-A94F-000D3AF063A5;720E9B43-490D-E911-A94D-000D3AF03089;41859E16-4B0D-E911-A94F-000D3AF063A5;4AC78157-4A0D-E911-A94F-000D3AF063A5;987F140A-4A0D-E911-A94E-000D3AF060A1;140DD567-490D-E911-A94F-000D3AF06923;BBC9B114-4A0D-E911-A94F-000D3AF06923;CDF5A4C1-4A0D-E911-A94D-000D3AF03089;FE746BA2-4B0D-E911-A94E-000D3AF060A1;28BDE7F1-490D-E911-A950-000D3AF06A16;8A035A3C-490D-E911-A950-000D3AF06A16;EE23D460-4B0D-E911-A950-000D3AF06A16;F5CAB114-4A0D-E911-A94F-000D3AF06923;38F41BC6-4B0D-E911-A94F-000D3AF063A5;EE62E832-490D-E911-A94F-000D3AF00F43;8EA3E750-490D-E911-A94E-000D3AF060A1;0C0DD567-490D-E911-A94F-000D3AF06923;F222D460-4B0D-E911-A950-000D3AF06A16;8F842336-4C0D-E911-A94F-000D3AF06923;8BA760B6-4A0D-E911-A94F-000D3AF00F43;1E76F4FE-490D-E911-A94D-000D3AF03089;305943CB-4A0D-E911-A94F-000D3AF06923;70F31BC6-4B0D-E911-A94F-000D3AF063A5;540C45AF-480D-E911-A94F-000D3AF06923;94842336-4C0D-E911-A94F-000D3AF06923;5522D460-4B0D-E911-A950-000D3AF06A16;7EA507A5-490D-E911-A94F-000D3AF063A5;40849E16-4B0D-E911-A94F-000D3AF063A5;54C78157-4A0D-E911-A94F-000D3AF063A5;317BAC1A-4A0D-E911-A94F-000D3AF06923;D85943CB-4A0D-E911-A94F-000D3AF06923;F5C78157-4A0D-E911-A94F-000D3AF063A5;E9025A3C-490D-E911-A950-000D3AF06A16;C660E832-490D-E911-A94F-000D3AF00F43;B02047C5-4A0D-E911-A94F-000D3AF06923;700CD567-490D-E911-A94F-000D3AF06923;566FEB04-4A0D-E911-A94D-000D3AF03089;0BB2109F-490D-E911-A94F-000D3AF063A5;E478AC1A-4A0D-E911-A94F-000D3AF06923;C80894FA-490D-E911-A94F-000D3AF00F43;62F2A4C1-4A0D-E911-A94D-000D3AF03089;06A02C30-4C0D-E911-A94F-000D3AF06923;5A03ADAF-480D-E911-A94E-000D3AF060A1;D1726BA2-4B0D-E911-A94E-000D3AF060A1;8CC95219-4C0D-E911-A950-000D3AF06A16;8BA507A5-490D-E911-A94F-000D3AF063A5;950894FA-490D-E911-A94F-000D3AF00F43;F563E832-490D-E911-A94F-000D3AF00F43;0C79AC1A-4A0D-E911-A94F-000D3AF06923;9025D460-4B0D-E911-A950-000D3AF06A16;EB30AD80-4B0D-E911-A94F-000D3AF06923;F40CD567-490D-E911-A94F-000D3AF06923;7622D460-4B0D-E911-A950-000D3AF06A16;53109B43-490D-E911-A94D-000D3AF03089;0F07DBB0-4A0D-E911-A950-000D3AF06A16;DC0D45AF-480D-E911-A94F-000D3AF06923;7F07DBB0-4A0D-E911-A950-000D3AF06A16;44B6109F-490D-E911-A94F-000D3AF063A5;C50D45AF-480D-E911-A94F-000D3AF06923;FB706BA2-4B0D-E911-A94E-000D3AF060A1;8DF41BC6-4B0D-E911-A94F-000D3AF063A5;B9839E16-4B0D-E911-A94F-000D3AF063A5;873E3CB5-480D-E911-A94F-000D3AF06923;675DF14A-490D-E911-A94E-000D3AF060A1;D3684278-4B0D-E911-A94F-000D3AF00F43;F6BE1C04-4A0D-E911-A94E-000D3AF060A1;97F6A4C1-4A0D-E911-A94D-000D3AF03089;AE6EEB04-4A0D-E911-A94D-000D3AF03089;885BA23D-490D-E911-A94D-000D3AF03089;61C11C04-4A0D-E911-A94E-000D3AF060A1;E56B8585-4B0D-E911-A94D-000D3AF03089;C56F6BA2-4B0D-E911-A94E-000D3AF060A1;F9849E16-4B0D-E911-A94F-000D3AF063A5;7C60E832-490D-E911-A94F-000D3AF00F43;6DB39BF4-490D-E911-A94F-000D3AF00F43;7B05ADAF-480D-E911-A94E-000D3AF060A1;D259A23D-490D-E911-A94D-000D3AF03089;08C95219-4C0D-E911-A950-000D3AF06A16;A4C31C04-4A0D-E911-A94E-000D3AF060A1;A0F4CAAA-4A0D-E911-A950-000D3AF06A16;7C0F9B43-490D-E911-A94D-000D3AF03089;DCA3E750-490D-E911-A94E-000D3AF060A1;9F829E16-4B0D-E911-A94F-000D3AF063A5;E224D460-4B0D-E911-A950-000D3AF06A16;79869E16-4B0D-E911-A94F-000D3AF063A5;35E7DC61-490D-E911-A94F-000D3AF06923;057A28ED-480D-E911-A94F-000D3AF063A5;8FC85219-4C0D-E911-A950-000D3AF06A16;2A63E832-490D-E911-A94F-000D3AF00F43;747728ED-480D-E911-A94F-000D3AF063A5;93131BCC-4B0D-E911-A94F-000D3AF063A5;9363E832-490D-E911-A94F-000D3AF00F43;AA23D460-4B0D-E911-A950-000D3AF06A16;7017EFFD-490D-E911-A950-000D3AF06A16;52A7E750-490D-E911-A94E-000D3AF060A1;5705ADAF-480D-E911-A94E-000D3AF060A1;DC79AC1A-4A0D-E911-A94F-000D3AF06923;70A8E750-490D-E911-A94E-000D3AF060A1;FC09DBB0-4A0D-E911-A950-000D3AF06A16;8C5AA23D-490D-E911-A94D-000D3AF03089;A463E832-490D-E911-A94F-000D3AF00F43;857A28ED-480D-E911-A94F-000D3AF063A5;C858A23D-490D-E911-A94D-000D3AF03089;DD5843CB-4A0D-E911-A94F-000D3AF06923;7733AD80-4B0D-E911-A94F-000D3AF06923;4E07DBB0-4A0D-E911-A950-000D3AF06A16;B3CAB114-4A0D-E911-A94F-000D3AF06923;9D869E16-4B0D-E911-A94F-000D3AF063A5;A074F4FE-490D-E911-A94D-000D3AF03089;D16A8585-4B0D-E911-A94D-000D3AF03089;4521D460-4B0D-E911-A950-000D3AF06A16;350CD567-490D-E911-A94F-000D3AF06923;E0088E7F-4B0D-E911-A94D-000D3AF03089;26A660B6-4A0D-E911-A94F-000D3AF00F43;0324D460-4B0D-E911-A950-000D3AF06A16;355943CB-4A0D-E911-A94F-000D3AF06923;95C98157-4A0D-E911-A94F-000D3AF063A5;0926F8F7-490D-E911-A950-000D3AF06A16;AC73F4FE-490D-E911-A94D-000D3AF03089;F1664278-4B0D-E911-A94F-000D3AF00F43;6B9F2C30-4C0D-E911-A94F-000D3AF06923;4DBDE7F1-490D-E911-A950-000D3AF06A16;4C78AC1A-4A0D-E911-A94F-000D3AF06923;AAF874C8-4A0D-E911-A94E-000D3AF060A1;8AE7DC61-490D-E911-A94F-000D3AF06923;170A8E7F-4B0D-E911-A94D-000D3AF03089;6875F4FE-490D-E911-A94D-000D3AF03089;62F874C8-4A0D-E911-A94E-000D3AF060A1;687F140A-4A0D-E911-A94E-000D3AF060A1;F47A28ED-480D-E911-A94F-000D3AF063A5;B2B59BF4-490D-E911-A94F-000D3AF00F43;C0BBE7F1-490D-E911-A950-000D3AF06A16;80045A3C-490D-E911-A950-000D3AF06A16;16C85219-4C0D-E911-A950-000D3AF06A16;03C21C04-4A0D-E911-A94E-000D3AF060A1;12151BCC-4B0D-E911-A94F-000D3AF063A5;D3078E7F-4B0D-E911-A94D-000D3AF03089;B3698585-4B0D-E911-A94D-000D3AF03089;B4E5DC61-490D-E911-A94F-000D3AF06923;E620D460-4B0D-E911-A950-000D3AF06A16;9CC88157-4A0D-E911-A94F-000D3AF063A5;9FF6A4C1-4A0D-E911-A94D-000D3AF03089;1776F4FE-490D-E911-A94D-000D3AF03089;D15EE832-490D-E911-A94F-000D3AF00F43;093F3CB5-480D-E911-A94F-000D3AF06923;B4B162A8-4B0D-E911-A94E-000D3AF060A1;F575F4FE-490D-E911-A94D-000D3AF03089;057D28ED-480D-E911-A94F-000D3AF063A5;B708DBB0-4A0D-E911-A950-000D3AF06A16;6EE6DC61-490D-E911-A94F-000D3AF06923;706A8585-4B0D-E911-A94D-000D3AF03089;B7577DC2-4A0D-E911-A94E-000D3AF060A1;A2A460B6-4A0D-E911-A94F-000D3AF00F43;CEC75219-4C0D-E911-A950-000D3AF06A16;79A7E750-490D-E911-A94E-000D3AF060A1;12A360B6-4A0D-E911-A94F-000D3AF00F43;4A5FF14A-490D-E911-A94E-000D3AF060A1;55C11C04-4A0D-E911-A94E-000D3AF060A1;0FB39BF4-490D-E911-A94F-000D3AF00F43;2864E832-490D-E911-A94F-000D3AF00F43;AF30AD80-4B0D-E911-A94F-000D3AF06923;83A260B6-4A0D-E911-A94F-000D3AF00F43;82F4A4C1-4A0D-E911-A94D-000D3AF03089;3DB29BF4-490D-E911-A94F-000D3AF00F43;A6C8B114-4A0D-E911-A94F-000D3AF06923;270CD567-490D-E911-A94F-000D3AF06923;AD859E16-4B0D-E911-A94F-000D3AF063A5;780794FA-490D-E911-A94F-000D3AF00F43;CAB49BF4-490D-E911-A94F-000D3AF00F43;2AF4CAAA-4A0D-E911-A950-000D3AF06A16;8DBCE7F1-490D-E911-A950-000D3AF06A16;A15FE832-490D-E911-A94F-000D3AF00F43;6D2AF8F7-490D-E911-A950-000D3AF06A16;5504ADAF-480D-E911-A94E-000D3AF060A1;4B5B43CB-4A0D-E911-A94F-000D3AF06923;6902ADAF-480D-E911-A94E-000D3AF060A1;4201ADAF-480D-E911-A94E-000D3AF060A1;C42BF8F7-490D-E911-A950-000D3AF06A16;C9842336-4C0D-E911-A94F-000D3AF06923;8BA660B6-4A0D-E911-A94F-000D3AF00F43;99F5CAAA-4A0D-E911-A950-000D3AF06A16;3A716BA2-4B0D-E911-A94E-000D3AF060A1;E0F4CAAA-4A0D-E911-A950-000D3AF06A16;ED859E16-4B0D-E911-A94F-000D3AF063A5;7C0D45AF-480D-E911-A94F-000D3AF06923;F6577DC2-4A0D-E911-A94E-000D3AF060A1;437D28ED-480D-E911-A94F-000D3AF063A5;DE577DC2-4A0D-E911-A94E-000D3AF060A1;5E098E7F-4B0D-E911-A94D-000D3AF03089;1F33B67A-4B0D-E911-A94F-000D3AF06923;2778AC1A-4A0D-E911-A94F-000D3AF06923;C2869E16-4B0D-E911-A94F-000D3AF063A5;7B0C45AF-480D-E911-A94F-000D3AF06923;5CE7DC61-490D-E911-A94F-000D3AF06923;CE0794FA-490D-E911-A94F-000D3AF00F43;B86FEB04-4A0D-E911-A94D-000D3AF03089;C0C75219-4C0D-E911-A950-000D3AF06A16;C780140A-4A0D-E911-A94E-000D3AF060A1;7EF5CAAA-4A0D-E911-A950-000D3AF06A16;385D43CB-4A0D-E911-A94F-000D3AF06923;16664278-4B0D-E911-A94F-000D3AF00F43;D2A6E750-490D-E911-A94E-000D3AF060A1;75E7DC61-490D-E911-A94F-000D3AF06923;3304ADAF-480D-E911-A94E-000D3AF060A1;E5F41BC6-4B0D-E911-A94F-000D3AF063A5;D2005A3C-490D-E911-A950-000D3AF06A16;EDF41BC6-4B0D-E911-A94F-000D3AF063A5;BEA407A5-490D-E911-A94F-000D3AF063A5;95A6E750-490D-E911-A94E-000D3AF060A1;F6F4CAAA-4A0D-E911-A950-000D3AF06A16;E5F674C8-4A0D-E911-A94E-000D3AF060A1;835943CB-4A0D-E911-A94F-000D3AF06923;35C41C04-4A0D-E911-A94E-000D3AF060A1;6EF31BC6-4B0D-E911-A94F-000D3AF063A5;4FF0A4C1-4A0D-E911-A94D-000D3AF03089;13C98157-4A0D-E911-A94F-000D3AF063A5;CFA02C30-4C0D-E911-A94F-000D3AF06923;03214B1F-4C0D-E911-A950-000D3AF06A16;3BA260B6-4A0D-E911-A94F-000D3AF00F43;125A43CB-4A0D-E911-A94F-000D3AF06923;636A4278-4B0D-E911-A94F-000D3AF00F43;21B6109F-490D-E911-A94F-000D3AF063A5;327C28ED-480D-E911-A94F-000D3AF063A5;4DF51BC6-4B0D-E911-A94F-000D3AF063A5;5B0694FA-490D-E911-A94F-000D3AF00F43;62F61BC6-4B0D-E911-A94F-000D3AF063A5;7303ADAF-480D-E911-A94E-000D3AF060A1;170E9B43-490D-E911-A94D-000D3AF03089;A2A360B6-4A0D-E911-A94F-000D3AF00F43;13A460B6-4A0D-E911-A94F-000D3AF00F43;DC5C43CB-4A0D-E911-A94F-000D3AF06923;04587DC2-4A0D-E911-A94E-000D3AF060A1;757A28ED-480D-E911-A94F-000D3AF063A5;24716BA2-4B0D-E911-A94E-000D3AF060A1;585C43CB-4A0D-E911-A94F-000D3AF06923;AC0E45AF-480D-E911-A94F-000D3AF06923;DDC65219-4C0D-E911-A950-000D3AF06A16;0261E832-490D-E911-A94F-000D3AF00F43;04C9B114-4A0D-E911-A94F-000D3AF06923;A60BD567-490D-E911-A94F-000D3AF06923;CE05ADAF-480D-E911-A94E-000D3AF060A1;EF60E832-490D-E911-A94F-000D3AF00F43;0863E832-490D-E911-A94F-000D3AF00F43;68BCE7F1-490D-E911-A950-000D3AF06A16;A7D63BD1-4A0D-E911-A94F-000D3AF06923;3A5D43CB-4A0D-E911-A94F-000D3AF06923;F47D28ED-480D-E911-A94F-000D3AF063A5;D8C11C04-4A0D-E911-A94E-000D3AF060A1;DE131BCC-4B0D-E911-A94F-000D3AF063A5;0509DBB0-4A0D-E911-A950-000D3AF06A16;7EF51BC6-4B0D-E911-A94F-000D3AF063A5;1C1F785D-4A0D-E911-A94F-000D3AF063A5;6A254B1F-4C0D-E911-A950-000D3AF06A16;8DA460B6-4A0D-E911-A94F-000D3AF00F43;820C45AF-480D-E911-A94F-000D3AF06923;22C01C04-4A0D-E911-A94E-000D3AF060A1;89224B1F-4C0D-E911-A950-000D3AF06A16;59141BCC-4B0D-E911-A94F-000D3AF063A5;0BD73BD1-4A0D-E911-A94F-000D3AF06923;5FF61BC6-4B0D-E911-A94F-000D3AF063A5;65F51BC6-4B0D-E911-A94F-000D3AF063A5;F0C9B114-4A0D-E911-A94F-000D3AF06923;0C0E45AF-480D-E911-A94F-000D3AF06923;D4C9B114-4A0D-E911-A94F-000D3AF06923;FB577DC2-4A0D-E911-A94E-000D3AF060A1;9FBCE7F1-490D-E911-A950-000D3AF06A16;EAA6E750-490D-E911-A94E-000D3AF060A1;DB07DBB0-4A0D-E911-A950-000D3AF06A16;7024D460-4B0D-E911-A950-000D3AF06A16;36C9B114-4A0D-E911-A94F-000D3AF06923;E9005A3C-490D-E911-A950-000D3AF06A16;CB17EFFD-490D-E911-A950-000D3AF06A16;E3BBE7F1-490D-E911-A950-000D3AF06A16;CA1F785D-4A0D-E911-A94F-000D3AF063A5;5E716BA2-4B0D-E911-A94E-000D3AF060A1;2B70EB04-4A0D-E911-A94D-000D3AF03089;FCC01C04-4A0D-E911-A94E-000D3AF060A1;F332B67A-4B0D-E911-A94F-000D3AF06923;94B2109F-490D-E911-A94F-000D3AF063A5;74F3CAAA-4A0D-E911-A950-000D3AF06A16;4F26D460-4B0D-E911-A950-000D3AF06A16;C72BF8F7-490D-E911-A950-000D3AF06A16;6C0CD567-490D-E911-A94F-000D3AF06923;4EB29BF4-490D-E911-A94F-000D3AF00F43;ED73F4FE-490D-E911-A94D-000D3AF03089;951E785D-4A0D-E911-A94F-000D3AF063A5;3FF774C8-4A0D-E911-A94E-000D3AF060A1;99849E16-4B0D-E911-A94F-000D3AF063A5;B6849E16-4B0D-E911-A94F-000D3AF063A5;B3BF1C04-4A0D-E911-A94E-000D3AF060A1;63E7DC61-490D-E911-A94F-000D3AF06923;26015A3C-490D-E911-A950-000D3AF06A16;0B20785D-4A0D-E911-A94F-000D3AF063A5;CBA8E750-490D-E911-A94E-000D3AF060A1;24141BCC-4B0D-E911-A94F-000D3AF063A5;28694278-4B0D-E911-A94F-000D3AF00F43;D66E6BA2-4B0D-E911-A94E-000D3AF060A1;B1C01C04-4A0D-E911-A94E-000D3AF060A1;325D43CB-4A0D-E911-A94F-000D3AF06923;23CBB114-4A0D-E911-A94F-000D3AF06923;24F974C8-4A0D-E911-A94E-000D3AF060A1;23C31C04-4A0D-E911-A94E-000D3AF060A1;935A43CB-4A0D-E911-A94F-000D3AF06923;BD5A43CB-4A0D-E911-A94F-000D3AF06923;B4C21C04-4A0D-E911-A94E-000D3AF060A1;D2BF1C04-4A0D-E911-A94E-000D3AF060A1;2E7B28ED-480D-E911-A94F-000D3AF063A5;3430AD80-4B0D-E911-A94F-000D3AF06923;2FF2A4C1-4A0D-E911-A94D-000D3AF03089;E316EFFD-490D-E911-A950-000D3AF06A16;E90F9B43-490D-E911-A94D-000D3AF03089;B25C43CB-4A0D-E911-A94F-000D3AF06923;ED0694FA-490D-E911-A94F-000D3AF00F43;C677F4FE-490D-E911-A94D-000D3AF03089;1EA41EF3-480D-E911-A94F-000D3AF063A5;D377F4FE-490D-E911-A94D-000D3AF03089;6EA660B6-4A0D-E911-A94F-000D3AF00F43;4676F4FE-490D-E911-A94D-000D3AF03089;4529F8F7-490D-E911-A950-000D3AF06A16;85716BA2-4B0D-E911-A94E-000D3AF060A1;91F4CAAA-4A0D-E911-A950-000D3AF06A16;3C6A8585-4B0D-E911-A94D-000D3AF03089;BEF31BC6-4B0D-E911-A94F-000D3AF063A5;513D3CB5-480D-E911-A94F-000D3AF06923;CBC31C04-4A0D-E911-A94E-000D3AF060A1;8DE4DC61-490D-E911-A94F-000D3AF06923;65A360B6-4A0D-E911-A94F-000D3AF00F43;025A43CB-4A0D-E911-A94F-000D3AF06923;ED842336-4C0D-E911-A94F-000D3AF06923;690AD567-490D-E911-A94F-000D3AF06923;DE587DC2-4A0D-E911-A94E-000D3AF060A1;41A507A5-490D-E911-A94F-000D3AF063A5;E975F4FE-490D-E911-A94D-000D3AF03089;A4716BA2-4B0D-E911-A94E-000D3AF060A1;C80BD567-490D-E911-A94F-000D3AF06923;E906ADAF-480D-E911-A94E-000D3AF060A1;E6B6109F-490D-E911-A94F-000D3AF063A5;405EE832-490D-E911-A94F-000D3AF00F43;947BAC1A-4A0D-E911-A94F-000D3AF06923;30BDE7F1-490D-E911-A950-000D3AF06A16;2860F14A-490D-E911-A94E-000D3AF060A1;5D60F14A-490D-E911-A94E-000D3AF060A1;00A660B6-4A0D-E911-A94F-000D3AF00F43;6DF4CAAA-4A0D-E911-A950-000D3AF06A16;A076F4FE-490D-E911-A94D-000D3AF03089;1D0DD567-490D-E911-A94F-000D3AF06923;99F0A4C1-4A0D-E911-A94D-000D3AF03089;9F03ADAF-480D-E911-A94E-000D3AF060A1;5B23D460-4B0D-E911-A950-000D3AF06A16;95111BCC-4B0D-E911-A94F-000D3AF063A5;E562E832-490D-E911-A94F-000D3AF00F43;171F785D-4A0D-E911-A94F-000D3AF063A5;10244B1F-4C0D-E911-A950-000D3AF06A16;7636B67A-4B0D-E911-A94F-000D3AF06923;1B70EB04-4A0D-E911-A94D-000D3AF03089;685DF14A-490D-E911-A94E-000D3AF060A1;7E869E16-4B0D-E911-A94F-000D3AF063A5;61698585-4B0D-E911-A94D-000D3AF03089;5F59A23D-490D-E911-A94D-000D3AF03089;A25C43CB-4A0D-E911-A94F-000D3AF06923;B5C75219-4C0D-E911-A950-000D3AF06A16;33B6109F-490D-E911-A94F-000D3AF063A5;B00E9B43-490D-E911-A94D-000D3AF03089;7BA760B6-4A0D-E911-A94F-000D3AF00F43;5560E832-490D-E911-A94F-000D3AF00F43;17706BA2-4B0D-E911-A94E-000D3AF060A1;8D025A3C-490D-E911-A950-000D3AF06A16;48A260B6-4A0D-E911-A94F-000D3AF00F43;32684278-4B0D-E911-A94F-000D3AF00F43;81F41BC6-4B0D-E911-A94F-000D3AF063A5;3780140A-4A0D-E911-A94E-000D3AF060A1;75B59BF4-490D-E911-A94F-000D3AF00F43;950E45AF-480D-E911-A94F-000D3AF06923;963C3CB5-480D-E911-A94F-000D3AF06923;E95843CB-4A0D-E911-A94F-000D3AF06923;176A8585-4B0D-E911-A94D-000D3AF03089;42A507A5-490D-E911-A94F-000D3AF063A5;4029F8F7-490D-E911-A950-000D3AF06A16;2E7D28ED-480D-E911-A94F-000D3AF063A5;B216EFFD-490D-E911-A950-000D3AF06A16;06F41BC6-4B0D-E911-A94F-000D3AF063A5;C5706BA2-4B0D-E911-A94E-000D3AF060A1;7C23D460-4B0D-E911-A950-000D3AF06A16;6260E832-490D-E911-A94F-000D3AF00F43;FBB3109F-490D-E911-A94F-000D3AF063A5;4627D460-4B0D-E911-A950-000D3AF06A16;CC694278-4B0D-E911-A94F-000D3AF00F43;29869E16-4B0D-E911-A94F-000D3AF063A5;63684278-4B0D-E911-A94F-000D3AF00F43;A3F6A4C1-4A0D-E911-A94D-000D3AF03089;F523D460-4B0D-E911-A950-000D3AF06A16;D4849E16-4B0D-E911-A94F-000D3AF063A5;E523D460-4B0D-E911-A950-000D3AF06A16;BCC8B114-4A0D-E911-A94F-000D3AF06923;BB6EEB04-4A0D-E911-A94D-000D3AF03089;D30894FA-490D-E911-A94F-000D3AF00F43;6906ADAF-480D-E911-A94E-000D3AF060A1;4EB5109F-490D-E911-A94F-000D3AF063A5;AC5EF14A-490D-E911-A94E-000D3AF060A1;2D859E16-4B0D-E911-A94F-000D3AF063A5;66E5DC61-490D-E911-A94F-000D3AF06923;D65DF14A-490D-E911-A94E-000D3AF060A1;21CAB114-4A0D-E911-A94F-000D3AF06923;2309DBB0-4A0D-E911-A950-000D3AF06A16;91C78157-4A0D-E911-A94F-000D3AF063A5;491E785D-4A0D-E911-A94F-000D3AF063A5;357C28ED-480D-E911-A94F-000D3AF063A5;3D5C43CB-4A0D-E911-A94F-000D3AF06923;E0E6DC61-490D-E911-A94F-000D3AF06923;257AAC1A-4A0D-E911-A94F-000D3AF06923;5DC41C04-4A0D-E911-A94E-000D3AF060A1;CF0BD567-490D-E911-A94F-000D3AF06923;2F3C3CB5-480D-E911-A94F-000D3AF06923;E3C85219-4C0D-E911-A950-000D3AF06A16;76045A3C-490D-E911-A950-000D3AF06A16;EBCAB114-4A0D-E911-A94F-000D3AF06923;110C45AF-480D-E911-A94F-000D3AF06923;2E5C43CB-4A0D-E911-A94F-000D3AF06923;080E9B43-490D-E911-A94D-000D3AF03089;1C839E16-4B0D-E911-A94F-000D3AF063A5;A461E832-490D-E911-A94F-000D3AF00F43;50C9B114-4A0D-E911-A94F-000D3AF06923;EC07DBB0-4A0D-E911-A950-000D3AF06A16;67849E16-4B0D-E911-A94F-000D3AF063A5;39E7DC61-490D-E911-A94F-000D3AF06923;587828ED-480D-E911-A94F-000D3AF063A5;C1644278-4B0D-E911-A94F-000D3AF00F43;805B43CB-4A0D-E911-A94F-000D3AF06923;329F5242-490D-E911-A950-000D3AF06A16;CB79AC1A-4A0D-E911-A94F-000D3AF06923;53C11C04-4A0D-E911-A94E-000D3AF060A1;2B889E16-4B0D-E911-A94F-000D3AF063A5;2B78AC1A-4A0D-E911-A94F-000D3AF06923;47F2A4C1-4A0D-E911-A94D-000D3AF03089;E05EE832-490D-E911-A94F-000D3AF00F43;9A109B43-490D-E911-A94D-000D3AF03089;7E03ADAF-480D-E911-A94E-000D3AF060A1;C274F4FE-490D-E911-A94D-000D3AF03089;83141BCC-4B0D-E911-A94F-000D3AF063A5;1C09DBB0-4A0D-E911-A950-000D3AF06A16;B80C45AF-480D-E911-A94F-000D3AF06923;EE74F4FE-490D-E911-A94D-000D3AF03089;AB33AD80-4B0D-E911-A94F-000D3AF06923;A5141BCC-4B0D-E911-A94F-000D3AF063A5;337D28ED-480D-E911-A94F-000D3AF063A5;62577DC2-4A0D-E911-A94E-000D3AF060A1;105FE832-490D-E911-A94F-000D3AF00F43;F505ADAF-480D-E911-A94E-000D3AF060A1;EA0AD567-490D-E911-A94F-000D3AF06923;1DB6109F-490D-E911-A94F-000D3AF063A5;39F4CAAA-4A0D-E911-A950-000D3AF06A16;DB01ADAF-480D-E911-A94E-000D3AF060A1;8501ADAF-480D-E911-A94E-000D3AF060A1;AFC88157-4A0D-E911-A94F-000D3AF063A5;305C43CB-4A0D-E911-A94F-000D3AF06923;AA22D460-4B0D-E911-A950-000D3AF06A16;270AD567-490D-E911-A94F-000D3AF06923;C4005A3C-490D-E911-A950-000D3AF06A16;B6A160B6-4A0D-E911-A94F-000D3AF00F43;F4C85219-4C0D-E911-A950-000D3AF06A16;77E7DC61-490D-E911-A94F-000D3AF06923;D80BD567-490D-E911-A94F-000D3AF06923;48577DC2-4A0D-E911-A94E-000D3AF060A1;FD2AF8F7-490D-E911-A950-000D3AF06A16;A35BA23D-490D-E911-A94D-000D3AF03089;E178AC1A-4A0D-E911-A94F-000D3AF06923;2B6FEB04-4A0D-E911-A94D-000D3AF03089;52B5109F-490D-E911-A94F-000D3AF063A5;6B2DAD80-4B0D-E911-A94F-000D3AF06923;8208DBB0-4A0D-E911-A950-000D3AF06A16;0B04ADAF-480D-E911-A94E-000D3AF060A1;DB24D460-4B0D-E911-A950-000D3AF06A16;DF58A23D-490D-E911-A94D-000D3AF03089;38F4CAAA-4A0D-E911-A950-000D3AF06A16;8A088E7F-4B0D-E911-A94D-000D3AF03089;89716BA2-4B0D-E911-A94E-000D3AF060A1;A077F4FE-490D-E911-A94D-000D3AF03089;0C06ADAF-480D-E911-A94E-000D3AF060A1;0A79AC1A-4A0D-E911-A94F-000D3AF06923;BBC75219-4C0D-E911-A950-000D3AF06A16;8B109B43-490D-E911-A94D-000D3AF03089;6A79AC1A-4A0D-E911-A94F-000D3AF06923;AB7728ED-480D-E911-A94F-000D3AF063A5;A45843CB-4A0D-E911-A94";

            string[] guiddd = guids.Split(';');
            foreach (string id in guiddd)
            {
                Entity entity = service.Retrieve("contact", new Guid(id), new ColumnSet("mobilephone"));
                string mob = entity.GetAttributeValue<string>("mobilephone");
                mob = mob.Substring(2, 10);
                try
                {
                    Entity entity1 = new Entity(entity.LogicalName, entity.Id);
                    entity1["mobilephone"] = mob;
                    service.Update(entity1);
                    Console.WriteLine("ddd");
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    continue;
                }

            }
        }
    }

    // 

}
