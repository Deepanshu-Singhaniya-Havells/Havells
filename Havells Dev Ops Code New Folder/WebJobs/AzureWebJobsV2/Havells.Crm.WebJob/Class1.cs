using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Text;

namespace Havells.Crm.WebJob
{
   public static class Class1123
    {
        static public string _primaryField = string.Empty;
        public static  void Execute(IOrganizationService service)
        {


            var InvoiceId = "5435b98b-3c66-ed11-9562-6045bdac5778";
            SendURLD365Request reqParm = new SendURLD365Request();
            reqParm.InvoiceId = InvoiceId;
            SendPaymentUrlResponse sendPaymentUrlResponse = new SendPaymentUrlResponse();
            try
            {
                if (reqParm.InvoiceId == null)
                {
                    Console.WriteLine("Status || Failed");
                   Console.WriteLine("Message || Invalid Invoice GUID");
                }
                else
                {
                    Entity FoundInvoice = service.Retrieve("invoice", new Guid(InvoiceId), new ColumnSet(true));
                    EntityReference customerref = FoundInvoice.GetAttributeValue<EntityReference>("customerid");

                    Entity customer = service.Retrieve(customerref.LogicalName, customerref.Id, new ColumnSet("mobilephone", "emailaddress1"));

                    string mobile = customer.Contains("mobilephone") ? customer.GetAttributeValue<String>("mobilephone").ToString() : null;
                    string email = customer.Contains("emailaddress1") ? customer.GetAttributeValue<String>("emailaddress1").ToString() : null;

                    if (mobile == null)
                    {
                        Console.WriteLine("Status || Failed");
                       Console.WriteLine("Message || Invalid Customer Mobile Number");
                    }
                    else
                    {
                        reqParm.Amount = FoundInvoice.Contains("hil_receiptamount") ? FoundInvoice.GetAttributeValue<Money>("hil_receiptamount").Value.ToString() : "0";
                        if (decimal.Parse(reqParm.Amount) < 1)
                        {
                            Console.WriteLine("Status || Failed");
                           Console.WriteLine("Message || Invalid recipt amount, please enter correct recipt amount.");
                        }
                        else
                        {
                            SendPaymentUrlRequest req = new SendPaymentUrlRequest();
                            String comm = "create_invoice";
                            req.PROJECT = "D365";
                            req.command = comm.Trim();

                            RemotePaymentLinkDetails remotePaymentLinkDetails = new RemotePaymentLinkDetails();

                            QueryExpression Query = new QueryExpression("hil_paymentstatus");
                            Query.ColumnSet = new ColumnSet("hil_url");
                            Query.Criteria = new FilterExpression(LogicalOperator.And);
                            Query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, FoundInvoice.GetAttributeValue<string>("name"));
                            EntityCollection FoundPay = service.RetrieveMultiple(Query);
                            if (FoundPay.Entities.Count > 0)
                            {
                                Console.WriteLine("Status || Failed");
                               Console.WriteLine("Message || payment Link is Allready send to Customer");
                            }
                            else
                            {
                                EntityReference AddressId = FoundInvoice.GetAttributeValue<EntityReference>("hil_address");
                                Entity AddressCol = service.Retrieve(AddressId.LogicalName, AddressId.Id, new ColumnSet(true));

                                String state = AddressCol.Contains("hil_state") ? AddressCol.GetAttributeValue<EntityReference>("hil_state").Name.ToString() : string.Empty;
                                String zip = string.Empty;

                                string address = AddressCol.Contains("hil_businessgeo") ? AddressCol.GetAttributeValue<EntityReference>("hil_businessgeo").Name.ToString() : string.Empty;
                                zip = AddressCol.Contains("hil_pincode") ? AddressCol.GetAttributeValue<EntityReference>("hil_pincode").Name.ToString() : string.Empty;

                                string city = string.Empty;
                                decimal amt = Convert.ToDecimal(reqParm.Amount);
                                remotePaymentLinkDetails.amount = Math.Round(amt, 2).ToString();

                                Entity ent = service.Retrieve("hil_branch", AddressCol.GetAttributeValue<EntityReference>("hil_branch").Id, new ColumnSet("hil_mamorandumcode"));
                                string _txnId = string.Empty;
                                string _mamorandumCode = "";
                                if (ent.Attributes.Contains("hil_mamorandumcode"))
                                {
                                    _mamorandumCode = ent.GetAttributeValue<string>("hil_mamorandumcode");
                                }



                                _txnId = "D365_" + FoundInvoice.GetAttributeValue<string>("name");
                                remotePaymentLinkDetails.txnid = _txnId;
                                remotePaymentLinkDetails.firstname = FoundInvoice.Contains("customerid") ? FoundInvoice.GetAttributeValue<EntityReference>("customerid").Name.ToString() : string.Empty;
                                remotePaymentLinkDetails.email = customer.Contains("emailaddress1") ? customer.GetAttributeValue<String>("emailaddress1").ToString() : "abc@gmail.com";
                                remotePaymentLinkDetails.phone = customer.Contains("mobilephone") ? customer.GetAttributeValue<String>("mobilephone").ToString() : "";

                                remotePaymentLinkDetails.address1 = address.Length > 99 ? address.Substring(0, 99) : address;
                                remotePaymentLinkDetails.state = state;
                                remotePaymentLinkDetails.country = "India";
                                remotePaymentLinkDetails.template_id = "1";
                                remotePaymentLinkDetails.productinfo = _mamorandumCode; //"B2C_PAYUBIZ_TEST_SMS";
                                remotePaymentLinkDetails.validation_period = "24";
                                remotePaymentLinkDetails.send_email_now = "1";
                                remotePaymentLinkDetails.send_sms = "1";
                                remotePaymentLinkDetails.time_unit = "H";
                                remotePaymentLinkDetails.zipcode = zip;
                                req.RemotePaymentLinkDetails = remotePaymentLinkDetails;

                                #region logrequest             
                                Entity intigrationTrace = new Entity("hil_integrationtrace");
                                intigrationTrace["hil_entityname"] = FoundInvoice.LogicalName;
                                intigrationTrace["hil_entityid"] = FoundInvoice.Id.ToString();
                                intigrationTrace["hil_request"] = JsonConvert.SerializeObject(req);
                                intigrationTrace["hil_name"] = FoundInvoice.GetAttributeValue<string>("name");
                                Guid intigrationTraceID = service.Create(intigrationTrace);
                                #endregion logrequest


                                IntegrationConfiguration inconfig = GetIntegrationConfiguration(service, "Send Payment Link");

                                var client = new RestClient(inconfig.url);
                                client.Timeout = -1;
                                var request = new RestRequest(Method.POST);

                                string authInfo = inconfig.userName + ":" + inconfig.password;
                                authInfo = "Basic " + Convert.ToBase64String(Encoding.Default.GetBytes(authInfo));
                                request.AddHeader("Authorization", authInfo);
                                request.AddHeader("Content-Type", "application/json");
                                request.AddParameter("application/json", JsonConvert.SerializeObject(req), ParameterType.RequestBody);
                                IRestResponse response = client.Execute(request);

                                dynamic obj = JsonConvert.DeserializeObject<SendPaymentUrlResponse>(response.Content);
                               
                                #region logresponse
                                Entity intigrationTraceUp = new Entity("hil_integrationtrace");
                                intigrationTraceUp["hil_response"] = response.Content == "" ? response.ErrorMessage : response.Content;
                                intigrationTraceUp.Id = intigrationTraceID;
                                service.Update(intigrationTraceUp);
                                #endregion logresponse

                                if (obj.msg == null)
                                {
                                    string url = obj.URL;
                                    string[] invoicenumber = url.Split('=');
                                    Entity statusPayment = new Entity("hil_paymentstatus");
                                    statusPayment["hil_name"] = obj.Transaction_Id;
                                    statusPayment["hil_url"] = obj.URL;
                                    statusPayment["hil_statussendurl"] = obj.Status;
                                    statusPayment["hil_email_id"] = obj.Email_Id;
                                    statusPayment["hil_phone"] = obj.Phone;
                                    statusPayment["hil_invoiceid"] = invoicenumber[1];
                                    service.Create(statusPayment);

                                    Entity entinvoice = new Entity(FoundInvoice.LogicalName, FoundInvoice.Id);
                                    entinvoice["statuscode"] = new OptionSetValue(4);
                                    service.Update(entinvoice);
                                    Console.WriteLine("Status || Sucess");
                                   Console.WriteLine("Message || Payment link send sucessfully");
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Status || Failed");
               Console.WriteLine("Message || D365 Internal Error " + ex.Message);
            }
        }
        private static IntegrationConfiguration GetIntegrationConfiguration(IOrganizationService _service, string name)
        {
            try
            {
                IntegrationConfiguration inconfig = new IntegrationConfiguration();
                QueryExpression qsCType = new QueryExpression("hil_integrationconfiguration");
                qsCType.ColumnSet = new ColumnSet("hil_url", "hil_username", "hil_password");
                qsCType.NoLock = true;
                qsCType.Criteria = new FilterExpression(LogicalOperator.And);
                qsCType.Criteria.AddCondition("hil_name", ConditionOperator.Equal, name);
                Entity integrationConfiguration = _service.RetrieveMultiple(qsCType)[0];
                inconfig.url = integrationConfiguration.GetAttributeValue<string>("hil_url");
                inconfig.userName = integrationConfiguration.GetAttributeValue<string>("hil_username");
                inconfig.password = integrationConfiguration.GetAttributeValue<string>("hil_password");
                return inconfig;
            }
            catch (Exception ex)
            {
                throw new Exception("Error : " + ex.Message);
            }
        }
    }
    public class ApprovalHelper
    {
        const String URL = @"https://havells.crm8.dynamics.com/main.aspx?appid=675ffa44-b6d0-ea11-a813-000d3af05d7b&forceUCI=1&pagetype=entityrecord&etn=";
        public static void createApproval(string entityName, string entityId, string purpose, string _primaryField, IOrganizationService service)
        {
            try
            {

                Console.WriteLine("createApproval method of Approval Helper is Started");

                QueryExpression _query = new QueryExpression("hil_approvalmatrix");
                _query.ColumnSet = new ColumnSet("hil_approvalmatrixid", "hil_name", "createdon", "hil_kpi",
                    "statecode", "hil_purpose", "hil_mailbody", "hil_level", "hil_entity", "hil_approver", "hil_discountmin", "hil_discountmax",
                    "hil_approverposition", "hil_copytoposition", "hil_duehrs", "hil_optional");
                _query.Criteria = new FilterExpression(LogicalOperator.And);
                _query.Criteria.AddCondition(new ConditionExpression("hil_entity", ConditionOperator.Equal, entityName));
                _query.Criteria.AddCondition(new ConditionExpression("hil_purpose", ConditionOperator.Equal, purpose));
                _query.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
                _query.AddOrder("hil_level", OrderType.Ascending);
                EntityCollection ApprovalsMatrixLines = service.RetrieveMultiple(_query);

                Guid lastApproval = Guid.Empty;
                Guid currentApproval = Guid.Empty;
                if (entityName == "hil_orderchecklist")
                {

                    Entity ocl = service.Retrieve(entityName, new Guid(entityId), new ColumnSet(true));
                    Decimal maxDiscoount = getMaxDiscount(ocl, service);
                   Console.WriteLine("getMaxDiscount ended value is " + maxDiscoount);


                    for (int i = 0; i < ApprovalsMatrixLines.Entities.Count; i++) // Entity entColl[i] in entColl.Entities)
                    {
                        decimal max = ApprovalsMatrixLines[i].GetAttributeValue<decimal>("hil_discountmax");
                       Console.WriteLine("getMaxDiscount ended value is " + max);

                        decimal min = ApprovalsMatrixLines[i].GetAttributeValue<decimal>("hil_discountmin");
                       Console.WriteLine("getMaxDiscount ended value is " + min);

                        if (maxDiscoount >= min)
                        {
                            ColumnSet collSet = findEntityColl(ApprovalsMatrixLines[i].GetAttributeValue<string>("hil_mailbody"), _primaryField, service );
                            Entity target = service.Retrieve(entityName, new Guid(entityId), collSet);
                            EntityReference targetOwner = target.GetAttributeValue<EntityReference>("ownerid");

                            _query = new QueryExpression("hil_userbranchmapping");
                            _query.ColumnSet = new ColumnSet("hil_name", "hil_zonalhead", "hil_user", "hil_salesoffice", "hil_buhead", "hil_branchproducthead", "hil_scm");
                            _query.Criteria = new FilterExpression(LogicalOperator.And);
                            _query.Criteria.AddCondition(new ConditionExpression("hil_user", ConditionOperator.Equal, targetOwner.Id));
                            _query.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));

                            EntityCollection userMapingColl = service.RetrieveMultiple(_query);

                            Entity _approval = new Entity("hil_approval");
                            EntityReference approver = new EntityReference();
                            _approval["subject"] = ApprovalsMatrixLines[i].GetAttributeValue<string>("hil_purpose") + "_Level " + ApprovalsMatrixLines[i].GetAttributeValue<int>("hil_level");

                           Console.WriteLine("levrl");
                            if (ApprovalsMatrixLines[i].Contains("hil_approver"))
                                approver = ApprovalsMatrixLines[i].GetAttributeValue<EntityReference>("hil_approver");
                            else if (ApprovalsMatrixLines[i].Contains("hil_approverposition"))
                            {
                                String approvalPosition = ApprovalsMatrixLines[i].GetAttributeValue<EntityReference>("hil_approverposition").Name;
                                if (userMapingColl.Entities.Count > 0)
                                    approver = getApproverByPosition(userMapingColl[0], approvalPosition, service  );
                                else
                                    throw new InvalidPluginExecutionException("Approver not Found");

                            }
                            _approval["ownerid"] = approver;
                            _approval["hil_level"] = new OptionSetValue(ApprovalsMatrixLines[i].GetAttributeValue<int>("hil_level"));
                            _approval["regardingobjectid"] = target.ToEntityReference();// new EntityReference(entityName, new Guid(entityId));
                            if ((ApprovalsMatrixLines[i].GetAttributeValue<int>("hil_level") == 1))
                            {
                                _approval["hil_approvalstatus"] = new OptionSetValue(3);// 3 for - submit for approval    4 for - Draft
                                _approval["hil_requesteddate"] = DateTime.Now.AddMinutes(330);
                            }
                            else
                            {
                                _approval["hil_approvalstatus"] = new OptionSetValue(4);// 4 for - Draft
                            }
                            //if (ApprovalsMatrixLines[i].Contains("hil_optional"))
                            //{
                            //    _approval["hil_optional"] = ApprovalsMatrixLines[i].GetAttributeValue<bool>("hil_optional");

                            //}
                            if (ApprovalsMatrixLines[i].Contains("hil_duehrs"))
                            {
                                _approval["scheduledend"] = DateTime.Now.AddHours(ApprovalsMatrixLines[i].GetAttributeValue<int>("hil_duehrs")).AddMinutes(330);
                            }
                            currentApproval = service.Create(_approval);
                            if ((ApprovalsMatrixLines[i].GetAttributeValue<int>("hil_level") == 1))
                            {
                                // string subject = "Approval Required for " + purpose + " ID " + target[_primaryField];

                                EntityCollection entToList = new EntityCollection();
                                entToList.EntityName = "systemuser";
                                //String approvalPosition = ApprovalsMatrixLines[i].GetAttributeValue<string>("hil_approverposition");
                                if (ApprovalsMatrixLines[i].Contains("hil_copytoposition"))
                                {

                                    string copyto = ApprovalsMatrixLines[i].GetAttributeValue<string>("hil_copytoposition");

                                    if (userMapingColl.Entities.Count > 0 && copyto.Contains(","))
                                    {
                                        entToList = getCopyToData(copyto, userMapingColl, service);
                                    }

                                }
                                Entity entTo = new Entity("activityparty");
                                entTo["partyid"] = targetOwner;
                                entToList.Entities.Add(entTo);

                                string mailbody = createEmailBody(target, ApprovalsMatrixLines[0].GetAttributeValue<string>("hil_mailbody"), approver.Name, collSet, service   );
                                string subject = createEmailSubject(target, _primaryField, ApprovalsMatrixLines[0].GetAttributeValue<string>("hil_kpi"), service   );
                                sendEmal(approver, entToList, target.ToEntityReference(), mailbody, subject, service);
                                lastApproval = currentApproval;


                            }
                            else
                            {
                                Entity entity = new Entity("hil_approval");
                                entity.Id = lastApproval;
                                entity["hil_nextapproval"] = new EntityReference(_approval.LogicalName, currentApproval);
                                //entity["hil_nextisoptional"] = ApprovalsMatrixLines[i].GetAttributeValue<bool>("hil_optional");
                                service.Update(entity);
                                lastApproval = currentApproval;
                            }
                        }
                    }
                    //throw new InvalidPluginExecutionException("hhhh");
                }
                else
                {
                    for (int i = 0; i < ApprovalsMatrixLines.Entities.Count; i++) // Entity entColl[i] in entColl.Entities)
                    {
                        ColumnSet collSet = findEntityColl(ApprovalsMatrixLines[i].GetAttributeValue<string>("hil_mailbody"), _primaryField, service   );
                        Entity target = service.Retrieve(entityName, new Guid(entityId), collSet);
                        EntityReference targetOwner = target.GetAttributeValue<EntityReference>("ownerid");

                        _query = new QueryExpression("hil_userbranchmapping");
                        _query.ColumnSet = new ColumnSet("hil_name", "hil_zonalhead", "hil_user", "hil_salesoffice", "hil_buhead", "hil_branchproducthead", "hil_scm");
                        _query.Criteria = new FilterExpression(LogicalOperator.And);
                        _query.Criteria.AddCondition(new ConditionExpression("hil_user", ConditionOperator.Equal, targetOwner.Id));
                        _query.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));

                        EntityCollection userMapingColl = service.RetrieveMultiple(_query);

                        Entity _approval = new Entity("hil_approval");
                        EntityReference approver = new EntityReference();
                        _approval["subject"] = ApprovalsMatrixLines[i].GetAttributeValue<string>("hil_purpose") + "_Level " + ApprovalsMatrixLines[i].GetAttributeValue<int>("hil_level");

                       Console.WriteLine("levrl");
                        if (ApprovalsMatrixLines[i].Contains("hil_approver"))
                            approver = ApprovalsMatrixLines[i].GetAttributeValue<EntityReference>("hil_approver");
                        else if (ApprovalsMatrixLines[i].Contains("hil_approverposition"))
                        {
                            String approvalPosition = ApprovalsMatrixLines[i].GetAttributeValue<EntityReference>("hil_approverposition").Name;
                            if (userMapingColl.Entities.Count > 0)
                                approver = getApproverByPosition(userMapingColl[0], approvalPosition, service   );
                            else
                                throw new InvalidPluginExecutionException("Approver not Found");

                        }
                        _approval["ownerid"] = approver;
                        _approval["hil_level"] = new OptionSetValue(ApprovalsMatrixLines[i].GetAttributeValue<int>("hil_level"));
                        _approval["regardingobjectid"] = target.ToEntityReference();// new EntityReference(entityName, new Guid(entityId));
                        if ((ApprovalsMatrixLines[i].GetAttributeValue<int>("hil_level") == 1))
                        {
                            _approval["hil_approvalstatus"] = new OptionSetValue(3);// 3 for - submit for approval    4 for - Draft
                            _approval["hil_requesteddate"] = DateTime.Now.AddMinutes(330);
                        }
                        else
                        {
                            _approval["hil_approvalstatus"] = new OptionSetValue(4);// 4 for - Draft
                        }
                        //if (ApprovalsMatrixLines[i].Contains("hil_optional"))
                        //{
                        //    _approval["hil_optional"] = ApprovalsMatrixLines[i].GetAttributeValue<bool>("hil_optional");

                        //}
                        if (ApprovalsMatrixLines[i].Contains("hil_duehrs"))
                        {
                            _approval["scheduledend"] = DateTime.Now.AddHours(ApprovalsMatrixLines[i].GetAttributeValue<int>("hil_duehrs")).AddMinutes(330);
                        }
                        currentApproval = service.Create(_approval);
                        if ((ApprovalsMatrixLines[i].GetAttributeValue<int>("hil_level") == 1))
                        {
                            // string subject = "Approval Required for " + purpose + " ID " + target[_primaryField];

                            EntityCollection entToList = new EntityCollection();
                            entToList.EntityName = "systemuser";
                            //String approvalPosition = ApprovalsMatrixLines[i].GetAttributeValue<string>("hil_approverposition");
                            if (ApprovalsMatrixLines[i].Contains("hil_copytoposition"))
                            {

                                string copyto = ApprovalsMatrixLines[i].GetAttributeValue<string>("hil_copytoposition");

                                if (userMapingColl.Entities.Count > 0 && copyto.Contains(","))
                                {
                                    entToList = getCopyToData(copyto, userMapingColl, service);
                                }

                            }
                            Entity entTo = new Entity("activityparty");
                            entTo["partyid"] = targetOwner;
                            entToList.Entities.Add(entTo);

                            string mailbody = createEmailBody(target, ApprovalsMatrixLines[0].GetAttributeValue<string>("hil_mailbody"), approver.Name, collSet, service   );
                            string subject = createEmailSubject(target, _primaryField, ApprovalsMatrixLines[0].GetAttributeValue<string>("hil_kpi"), service   );
                            // if (target.LogicalName == "hil_tenderbankguarantee")
                            sendEmal(approver, entToList, target.ToEntityReference(), mailbody, subject, service);
                            lastApproval = currentApproval;


                        }
                        else
                        {
                            Entity entity = new Entity("hil_approval");
                            entity.Id = lastApproval;
                            entity["hil_nextapproval"] = new EntityReference(_approval.LogicalName, currentApproval);
                            //entity["hil_nextisoptional"] = ApprovalsMatrixLines[i].GetAttributeValue<bool>("hil_optional");
                            service.Update(entity);
                            lastApproval = currentApproval;
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("|| createApproval Error: " + ex.Message);
            }
        }

        public static decimal getMaxDiscount(Entity orderCheckList, IOrganizationService service )
        {
            Console.WriteLine("getMaxDiscount Started");
            string fetchPODiscount = @"<fetch version='1.0' mapping='logical' distinct='false' aggregate='true'>
	                            <entity name='hil_orderchecklistproduct'>
		                            <attribute name='hil_orderchecklistid' alias='orderchecklist' groupby='true' />
		                            <attribute name='hil_podiscount' alias='amount' aggregate='max' />
		                            <filter type='and'>
			                            <condition attribute='hil_orderchecklistid' operator='eq' value='" + orderCheckList.Id + @"' />
			                            <condition attribute='statecode' operator='eq' value='0' />
		                            </filter>
	                            </entity>
                            </fetch>";
            string fetchOrderBookingDiscount = @"<fetch version='1.0' mapping='logical' distinct='false' aggregate='true'>
	                            <entity name='hil_orderchecklistproduct'>
		                            <attribute name='hil_orderchecklistid' alias='orderchecklist' groupby='true' />
		                            <attribute name='hil_discount' alias='amount' aggregate='max' />
		                            <filter type='and'>
			                            <condition attribute='hil_orderchecklistid' operator='eq' value='" + orderCheckList.Id + @"' />
			                            <condition attribute='statecode' operator='eq' value='0' />
		                            </filter>
	                            </entity>
                            </fetch>";
            if (orderCheckList.Contains("hil_tenderno"))
            {
                EntityCollection eColl = service.RetrieveMultiple(new FetchExpression(fetchPODiscount));
                return (decimal)((AliasedValue)eColl.Entities[0]["amount"]).Value;
            }
            else
            {
                EntityCollection eColl = service.RetrieveMultiple(new FetchExpression(fetchOrderBookingDiscount));
                return (decimal)((AliasedValue)eColl.Entities[0]["amount"]).Value;
            }
        }
        public static EntityReference getApproverByPosition(Entity userMaping, string approvalPosition, IOrganizationService service)
        {
           Console.WriteLine("getApproverByPosition Stratrd");
            EntityReference approvr = new EntityReference();
            if (approvalPosition == "Branch Product Head")
            {
                approvr = userMaping.GetAttributeValue<EntityReference>("hil_branchproducthead");
            }
            else if (approvalPosition == "Enquiry Creator")
            {
                approvr = userMaping.GetAttributeValue<EntityReference>("hil_user");
            }
            else if (approvalPosition == "Zonal Head")
            {
                approvr = userMaping.GetAttributeValue<EntityReference>("hil_zonalhead");
            }
            else if (approvalPosition == "BU Head")
            {
                approvr = userMaping.GetAttributeValue<EntityReference>("hil_buhead");
            }
            else if (approvalPosition == "Sr. Manager Finance")
            {
                QueryExpression query = new QueryExpression("systemuser");
                query.ColumnSet = new ColumnSet(true);
                query.Criteria = new FilterExpression(LogicalOperator.And);
                query.Criteria.AddCondition("positionid", ConditionOperator.Equal, new Guid("6532dc30-40a6-eb11-9442-6045bd72b6fd"));
                EntityCollection _entitys = service.RetrieveMultiple(query);

                if (_entitys.Entities.Count > 0)
                    approvr = _entitys[0].ToEntityReference();
            }
            else if (approvalPosition == "Sr. Manager Treasury")
            {
                QueryExpression query = new QueryExpression("user");
                query.ColumnSet = new ColumnSet(true);
                query.Criteria = new FilterExpression(LogicalOperator.And);
                query.Criteria.AddCondition("positionid", ConditionOperator.Equal, new Guid("{45C29172-3CD0-EB11-BACC-6045BD72E9C2}"));
                EntityCollection _entitys = service.RetrieveMultiple(query);
                if (_entitys.Entities.Count > 0)
                    approvr = _entitys[0].ToEntityReference();
            }
            else if (approvalPosition == "SCM")
            {
                approvr = userMaping.GetAttributeValue<EntityReference>("hil_scm");
            }
            //else if (approvalPosition == "Design Head")
            //{
            //    QueryExpression query = new QueryExpression("systemuser");
            //    query.ColumnSet = new ColumnSet(true);
            //    query.Criteria = new FilterExpression(LogicalOperator.And);
            //    query.Criteria.AddCondition("positionid", ConditionOperator.Equal, new Guid("6532dc30-40a6-eb11-9442-6045bd72b6fd"));
            //    EntityCollection _entitys = service.RetrieveMultiple(query);

            //    _approval["ownerid"] = entColl[i].GetAttributeValue<EntityReference>("hil_user"); = _entitys[0].ToEntityReference();
            //}
           Console.WriteLine("getApproverByPosition end");
            return approvr;
        }
        public static EntityCollection getCopyToData(string copyto, EntityCollection userMapingColl, IOrganizationService service)
        {
            EntityCollection entToList = new EntityCollection();
            entToList.EntityName = "systemuser";
            String[] mailto = copyto.Split(',');
            Entity entTo = new Entity("activityparty");
            foreach (string to in mailto)
            {
                if (to == "Branch Product Head")
                {
                    entTo = new Entity("activityparty");
                    entTo["partyid"] = userMapingColl[0].GetAttributeValue<EntityReference>("hil_branchproducthead");

                }
                //else if (to == "Design Team")
                //{
                //    entTo = new Entity("activityparty");
                //    entTo["partyid"] = userMapingColl[0].GetAttributeValue<EntityReference>("hil_branchproducthead");
                //}
                else if (to == "Enquiry Creator")
                {
                    entTo = new Entity("activityparty");
                    entTo["partyid"] = userMapingColl[0].GetAttributeValue<EntityReference>("hil_user"); ;
                }
                else if (to == "Zonal Head")
                {
                    entTo = new Entity("activityparty");
                    entTo["partyid"] = userMapingColl[0].GetAttributeValue<EntityReference>("hil_zonalhead");
                }
                else if (to == "BU Head")
                {
                    entTo = new Entity("activityparty");
                    entTo["partyid"] = userMapingColl[0].GetAttributeValue<EntityReference>("hil_buhead");
                }
                //else if (to == "Design Head")
                //{
                //    entTo = new Entity("activityparty");
                //    entTo["partyid"] = service.Retrieve("systemuser", (regardingENtity.GetAttributeValue<EntityReference>("hil_designteam")).Id, new ColumnSet("parentsystemuserid")).GetAttributeValue<EntityReference>("parentsystemuserid");
                //}
                else if (to == "Sr. Manager Finance")
                {
                    QueryExpression query = new QueryExpression("systemuser");
                    query.ColumnSet = new ColumnSet(true);
                    query.Criteria = new FilterExpression(LogicalOperator.And);
                    query.Criteria.AddCondition("positionid", ConditionOperator.Equal, new Guid("6532dc30-40a6-eb11-9442-6045bd72b6fd"));
                    EntityCollection _entitys = service.RetrieveMultiple(query);

                    entTo = new Entity("activityparty");
                    entTo["partyid"] = _entitys[0].ToEntityReference();
                }
                else if (to == "Sr. Manager Treasury")
                {
                    QueryExpression query = new QueryExpression("user");
                    query.ColumnSet = new ColumnSet(true);
                    query.Criteria = new FilterExpression(LogicalOperator.And);
                    query.Criteria.AddCondition("positionid", ConditionOperator.Equal, new Guid("{45C29172-3CD0-EB11-BACC-6045BD72E9C2}"));
                    EntityCollection _entitys = service.RetrieveMultiple(query);

                    entTo = new Entity("activityparty");
                    entTo["partyid"] = _entitys[0].ToEntityReference();
                }
                entToList.Entities.Add(entTo);
            }
            return entToList;
        }
        public static void sendEmal(EntityReference approver, EntityCollection copyto, EntityReference regarding, string mailbody, string subject, IOrganizationService service)
        {
            try
            {
                Entity entEmail = new Entity("email");

                Entity entFrom = new Entity("activityparty");
                entFrom["partyid"] = new EntityReference("queue", new Guid("a81ac51b-c9df-eb11-bacb-0022486ea537"));
                Entity[] entFromList = { entFrom };
                entEmail["from"] = entFromList;

                EntityReference to = approver;
                Entity toActivityParty = new Entity("activityparty");
                toActivityParty["partyid"] = to;
                entEmail["to"] = new Entity[] { toActivityParty };

                ;
                Entity ccActivityParty = new Entity("activityparty");
                ccActivityParty["partyid"] = copyto;
                entEmail["cc"] = copyto;

                entEmail["subject"] = subject;
                entEmail["description"] = mailbody;

                entEmail["regardingobjectid"] = regarding;

                Guid emailId = service.Create(entEmail);

                SendEmailRequest sendEmailReq = new SendEmailRequest()
                {
                    EmailId = emailId,
                    IssueSend = true
                };
                SendEmailResponse sendEmailRes = (SendEmailResponse)service.Execute(sendEmailReq);

            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error: " + ex.Message);
            }
        }
        public static string createEmailBody(Entity target, string mailbodyTemp, String approver, ColumnSet collSet, IOrganizationService service)
        {
           Console.WriteLine("createEmailBody Started");
            try
            {
                string recordURL = URL + target.LogicalName + "&id=" + target.Id;
                string clickHere = "<br> For more information please <a href =\"" + recordURL + "\"  target=\"_blank\"><b>click here.<b></a><br>";

                Dictionary<string, string> keyValue1 = new Dictionary<string, string>();
                keyValue1.Add("approverName", approver);
                keyValue1.Add("ClickHere", clickHere);

                foreach (string spl in collSet.Columns)
                {
                    if (target.Contains(spl))
                    {
                        var dataType = target[spl].GetType();
                        if (dataType.Name == "EntityReference")
                        {
                            keyValue1.Add(spl, ((EntityReference)target[spl]).Name);
                        }
                        else if (dataType.Name.ToLower() == "string")
                        {
                            keyValue1.Add(spl, ((string)target[spl]));
                        }
                        else if (dataType.Name.ToLower() == "datetime")
                        {
                            keyValue1.Add(spl, ((DateTime)target[spl]).ToString());
                        }
                        else if (dataType.Name.ToLower() == "decial")
                        {
                            keyValue1.Add(spl, ((decimal)target[spl]).ToString());
                        }
                        else if (dataType.Name.ToLower() == "Money".ToLower())
                        {
                            keyValue1.Add(spl, ((Money)target[spl]).Value.ToString());
                        }
                        else if (dataType.Name.ToLower() == "OptionSetValue".ToLower())
                        {
                            keyValue1.Add(spl, target.FormattedValues[spl].ToString());
                        }
                        else if (dataType.Name.ToLower() == "Boolean".ToLower())
                        {
                            keyValue1.Add(spl, ((bool)target[spl]).ToString());
                        }
                        else
                        {
                            throw new InvalidPluginExecutionException(" Error:- Data Type not found ");
                        }
                    }
                    else
                    {
                        keyValue1.Add(spl, "");
                    }
                }
                foreach (KeyValuePair<string, string> keyValue in keyValue1)
                {
                    string key = "{" + keyValue.Key + "}";
                    mailbodyTemp = mailbodyTemp.Replace(key, keyValue.Value);
                }
                //Console.WriteLine("createEmailBody Ended");
                return mailbodyTemp;
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("createEmailBody Error:- " + ex.Message);
            }
        }
        public static string createEmailSubject(Entity target1, string _primaryField, string subjectTemplate, IOrganizationService service)
        {
            //Console.WriteLine("createEmailBody Started");
            try
            {
                Dictionary<string, string> keyValue1 = new Dictionary<string, string>();


                ColumnSet collSet = findEntityColl(subjectTemplate, _primaryField, service   );
                Entity target = service.Retrieve(target1.LogicalName, target1.Id, collSet);

                foreach (string spl in collSet.Columns)
                {
                    if (target.Contains(spl))
                    {
                        var dataType = target[spl].GetType();
                        if (dataType.Name == "EntityReference")
                        {
                            keyValue1.Add(spl, ((EntityReference)target[spl]).Name);
                        }
                        else if (dataType.Name.ToLower() == "string")
                        {
                            keyValue1.Add(spl, ((string)target[spl]));
                        }
                        else if (dataType.Name.ToLower() == "datetime")
                        {
                            keyValue1.Add(spl, ((DateTime)target[spl]).ToString());
                        }
                        else if (dataType.Name.ToLower() == "decial")
                        {
                            keyValue1.Add(spl, ((decimal)target[spl]).ToString());
                        }
                        else if (dataType.Name.ToLower() == "Money".ToLower())
                        {
                            keyValue1.Add(spl, ((Money)target[spl]).Value.ToString());
                        }
                        else if (dataType.Name.ToLower() == "OptionSetValue".ToLower())
                        {
                            keyValue1.Add(spl, target.FormattedValues[spl].ToString());
                        }
                        else if (dataType.Name.ToLower() == "Boolean".ToLower())
                        {
                            keyValue1.Add(spl, ((bool)target[spl]).ToString());
                        }
                    }
                    else
                    {
                        keyValue1.Add(spl, "");
                    }
                }
                foreach (KeyValuePair<string, string> keyValue in keyValue1)
                {
                    string key = "{" + keyValue.Key + "}";
                    subjectTemplate = subjectTemplate.Replace(key, keyValue.Value);
                }
                //Console.WriteLine("createEmailBody Ended");
                return subjectTemplate;
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("createEmailSubject Error:- " + ex.Message);
            }
        }
        public static ColumnSet findEntityColl(string mailbodyTemp, string _primaryField, IOrganizationService service)
        {
           Console.WriteLine("findEntityColl Started");
            string[] split1 = mailbodyTemp.Split('{');
            string split2 = string.Empty;
           Console.WriteLine("1 " + split1.Length);
            foreach (string spl in split1)
            {
                if (spl.Contains("}"))
                {

                    string[] spl1 = spl.Split('}');
                    if (spl1[0] != "approverName" && spl1[0] != "ClickHere")
                        if (split2 == string.Empty)
                            split2 = split2 + spl1[0];
                        else if (!split2.Contains(spl1[0]))
                            split2 = split2 + "," + spl1[0];
                }
            }
            if (!split2.Contains("ownerid"))
                split2 = split2 + ",ownerid";

            if (!split2.Contains(_primaryField))
                split2 = split2 + _primaryField;
            string[] coll = split2.Split(',');
            ColumnSet cc = new ColumnSet(coll);
           Console.WriteLine("findEntityColl End");
            return cc;
        }
        public static void GetPrimaryIdFieldName(string _entityName, IOrganizationService service, out string _primaryField)
        {
            //Create RetrieveEntityRequest

            _primaryField = "";
            try
            {
                RetrieveEntityRequest retrievesEntityRequest = new RetrieveEntityRequest
                {
                    EntityFilters = EntityFilters.Entity,
                    LogicalName = _entityName
                };
                //Execute Request
                RetrieveEntityResponse retrieveEntityResponse = (RetrieveEntityResponse)service.Execute(retrievesEntityRequest);
                //_primaryKey = retrieveEntityResponse.EntityMetadata.PrimaryIdAttribute;
                _primaryField = retrieveEntityResponse.EntityMetadata.PrimaryNameAttribute;
            }
            catch (Exception ex)
            {
                _primaryField = "ERROR";
                _primaryField = ex.Message;
            }
        }
    }
    public class IntegrationConfiguration
    {
        public string url { get; set; }
        public string userName { get; set; }
        public string password { get; set; }
    }

    public class SendPaymentUrlRequest
    {

        public string PROJECT { get; set; }

        public string command { get; set; }

        public RemotePaymentLinkDetails RemotePaymentLinkDetails { get; set; }
    }

    public class RemotePaymentLinkDetails
    {

        public string amount { get; set; }

        public string txnid { get; set; }

        public string productinfo { get; set; }

        public string firstname { get; set; }

        public string email { get; set; }

        public string phone { get; set; }

        public string address1 { get; set; }

        public string city { get; set; }

        public string state { get; set; }

        public string country { get; set; }

        public string zipcode { get; set; }

        public string template_id { get; set; }

        public string validation_period { get; set; }

        public string send_email_now { get; set; }

        public string send_sms { get; set; }

        public string time_unit { get; set; }
    }

    public class SendPaymentUrlResponse
    {

        public string Email_Id { get; set; }

        public string Transaction_Id { get; set; }

        public string URL { get; set; }

        public string Status { get; set; }

        public string Phone { get; set; }

        public string StatusCode { get; set; }

        public string msg { get; set; }
    }

    public class SendURLD365Request
    {

        public string InvoiceId { get; set; }

        public string mobile { get; set; }

        public string Amount { get; set; }
    }

    public class StatusRequest
    {

        public string PROJECT { get; set; }

        public string command { get; set; }

        public string var1 { get; set; }

    }

    public class TransactionDetail
    {

        public string mihpayid { get; set; }

        public string request_id { get; set; }

        public string bank_ref_num { get; set; }

        public string amt { get; set; }

        public string transaction_amount { get; set; }

        public string txnid { get; set; }

        public string additional_charges { get; set; }

        public string productinfo { get; set; }

        public string firstname { get; set; }

        public string bankcode { get; set; }

        public string udf1 { get; set; }

        public string udf3 { get; set; }

        public string udf4 { get; set; }

        public string udf5 { get; set; }

        public string field2 { get; set; }

        public string field9 { get; set; }

        public string error_code { get; set; }

        public string addedon { get; set; }

        public string payment_source { get; set; }

        public string card_type { get; set; }

        public string error_Message { get; set; }

        public string net_amount_debit { get; set; }

        public string disc { get; set; }

        public string mode { get; set; }

        public string PG_TYPE { get; set; }

        public string card_no { get; set; }

        public string udf2 { get; set; }

        public string status { get; set; }

        public string unmappedstatus { get; set; }

        public string Merchant_UTR { get; set; }

        public string Settled_At { get; set; }
    }

    public class StatusResponse
    {

        public int status { get; set; }

        public string msg { get; set; }

        public List<TransactionDetail> transaction_details { get; set; }
    }

    public class PaymentStatusD365Response
    {

        public string Status { get; set; }
    }
}
