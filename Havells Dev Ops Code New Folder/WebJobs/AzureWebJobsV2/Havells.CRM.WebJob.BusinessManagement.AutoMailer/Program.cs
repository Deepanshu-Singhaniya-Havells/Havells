using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Text;

namespace Havells.CRM.WebJob.BusinessManagement.AutoMailer
{
    public class Program
    {
        #region Global Varialble declaration
        static IOrganizationService service;
        static public EntityReference sender = new EntityReference();
        public static String URL = "";
        #endregion
        static void Main(string[] args)
        {
            Console.WriteLine("Program Started.");

            var connStr = ConfigurationManager.AppSettings["connStr"].ToString();
            var CrmURL = ConfigurationManager.AppSettings["CRMUrl"].ToString();
            string finalString = string.Format(connStr, CrmURL);
            service = HavellsConnection.CreateConnection.createConnection(finalString);
            Console.WriteLine("Connection is Established..");
            try
            {
                OCLDeactivate.getAllDepartment(service);//, 2, new Guid("ce8b92cb-e64c-ec11-8f8e-6045bd733e10"));
                BOEReminder.SendBOEReminder(service);
            }
            catch(Exception ex)
            {
                Console.WriteLine("ERROR in Send BOE Reminder ***Error: " + ex.Message);
            }
            

            //Integration integration = Helper.IntegrationConfiguration(service, "TenderAppURL");
            //URL = integration.uri;
            //Console.WriteLine("Tender App URL retrived.");

            //sender = Helper.getSender("SMS", service);
            //Console.WriteLine("Retrive Sender.");
            //try
            //{
            //    Console.WriteLine("PPC mail for 3 days started.");
            //    PPCMails3Days();
            //}
            //catch (Exception ex)
            //{
            //    Console.WriteLine("ERROR in PPC mail for 3 days ***Error: " + ex.Message);
            //}
            //try
            //{
            //    Console.WriteLine("PPC Email for Todays is started");
            //    PPCMailstoday();
            //}
            //catch (Exception ex)
            //{
            //    Console.WriteLine("ERROR in PPC mail for Today ***Error:  " + ex.Message);
            //}
            //try
            //{
            //    Console.WriteLine("InspectionStatusUpadteToday started");
            //    InspectionStatusUpadteToday();
            //}
            //catch (Exception ex)
            //{
            //    Console.WriteLine("ERROR in InspectionStatusUpadteToday ***Error:   " + ex.Message);
            //}
            //try
            //{
            //    Console.WriteLine("BillingDocumentReviewbefore12days started");
            //    BillingDocumentReviewbefore12days();
            //}
            //catch (Exception ex)
            //{
            //    Console.WriteLine("ERROR in BillingDocumentReviewbefore12days ***Error:   " + ex.Message);
            //}
            //try
            //{
            //    Console.WriteLine("LCDocumentReview5DaysBefore started");
            //    LCDocumentReview5DaysBefore();
            //}
            //catch (Exception ex)
            //{
            //    Console.WriteLine("ERROR in LCDocumentReview5DaysBefore ***Error:   " + ex.Message);
            //}
            //try
            //{
            //    Console.WriteLine("CreditLimitDocumentReview10DaysBefore started");
            //    CreditLimitDocumentReview10DaysBefore();
            //}
            //catch (Exception ex)
            //{
            //    Console.WriteLine("ERROR in CreditLimitDocumentReview10DaysBefore ***Error:   " + ex.Message);
            //}
        }
        #region PPC EMAIL...
        public static void PPCMails3Days()
        {
            try
            {
                DateTime today = DateTime.Today.AddMinutes(330);
                DateTime toDate = DateTime.Today.AddMinutes(330).AddDays(4);
                QueryExpression queryExp = new QueryExpression("hil_oaproduct");
                queryExp.ColumnSet = new ColumnSet("ownerid", "hil_name", "hil_oaheader", "hil_stockreadinessdateppc", "hil_productdescription",
                    "hil_deliverydate", "hil_quantity", "hil_materialgroup", "hil_department");
                queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                queryExp.Criteria.AddCondition("hil_stockreadinessdateppc", ConditionOperator.OnOrAfter, today.AddDays(1));
                queryExp.Criteria.AddCondition("hil_stockreadinessdateppc", ConditionOperator.OnOrBefore, toDate);

                LinkEntity EntityA = new LinkEntity("hil_oaproduct", "hil_oaheader", "hil_oaheader", "hil_oaheaderid", JoinOperator.Inner);
                EntityA.Columns = new ColumnSet("hil_tenderid", "hil_customername", "hil_orderchecklistid");
                EntityA.EntityAlias = "OAHeader";
                queryExp.LinkEntities.Add(EntityA);
                EntityCollection OAProductCol = service.RetrieveMultiple(queryExp);

                Console.WriteLine("Record Retrived");
                if (OAProductCol.Entities.Count > 0)
                {
                    PPCMails(OAProductCol, true);
                }
                Console.WriteLine("PPC email for 3 days is Completed.");
                //}
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
            }
        }
        public static void PPCMailstoday()
        {
            try
            {
                QueryExpression queryExp = new QueryExpression("hil_oaproduct");
                queryExp.ColumnSet = new ColumnSet("ownerid", "hil_name", "hil_oaheader", "hil_stockreadinessdateppc", "hil_productdescription",
                     "hil_deliverydate", "hil_quantity", "hil_materialgroup", "hil_department");
                queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                queryExp.Criteria.AddCondition("hil_stockreadinessdateppc", ConditionOperator.On, DateTime.Today.AddMinutes(330));
                LinkEntity EntityA = new LinkEntity("hil_oaproduct", "hil_oaheader", "hil_oaheader", "hil_oaheaderid", JoinOperator.Inner);
                EntityA.Columns = new ColumnSet("hil_tenderid", "hil_customername", "hil_orderchecklistid");
                EntityA.EntityAlias = "OAHeader";
                queryExp.LinkEntities.Add(EntityA);
                EntityCollection OAProductCol = service.RetrieveMultiple(queryExp);

                if (OAProductCol.Entities.Count > 0)
                {
                    Console.WriteLine("Record Found.");
                    PPCMails(OAProductCol, false);
                }
                Console.WriteLine("PPC email for today is completed");
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
            }
        }
        public static void PPCMails(EntityCollection OAProductCol, bool today)
        {
            try
            {
                Console.WriteLine("PPC mail function started.");
                EntityReference owner = OAProductCol.Entities[0].GetAttributeValue<EntityReference>("ownerid");
                string bodtText = PPCMailBody(OAProductCol);
                Console.WriteLine("Mail Body Created.");
                Entity userConfiguration = Helper.getUserConfiguartion(owner, service);
                Console.WriteLine("User Configuration is retrived");
                EntityCollection entCCList = new EntityCollection();

                //Entity toemailer = getTeamUser(OAProductCol.Entities[0].GetAttributeValue<EntityReference>("hil_omsemailersetup"));
                //  Console.WriteLine("to is set");
                EntityCollection toTeamMembers = new EntityCollection();

                EntityReferenceCollection materialGroup = new EntityReferenceCollection();

                List<Guid> OAHEADERIS = new List<Guid>();
                foreach (Entity entityReference in OAProductCol.Entities)
                {
                    OAHEADERIS.Add(entityReference.GetAttributeValue<EntityReference>("hil_oaheader").Id);
                }


                QueryExpression query = new QueryExpression("hil_oaproduct");
                query.ColumnSet = new ColumnSet("hil_materialgroup");
                query.Distinct = true;
                query.Criteria = new FilterExpression(LogicalOperator.And);
                query.Criteria.AddCondition(new ConditionExpression("hil_oaheader", ConditionOperator.In, OAHEADERIS.ToArray()));
                EntityCollection entCol = service.RetrieveMultiple(query);


                foreach (Entity oaPrd in entCol.Entities)
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
                EntityReference ocl = new EntityReference();
                EntityReference department = OAProductCol[0].GetAttributeValue<EntityReference>("hil_department");
                if (OAProductCol[0].Contains("OAHeader.hil_orderchecklistid"))
                {
                    ocl = (EntityReference)OAProductCol[0].GetAttributeValue<AliasedValue>("OAHeader.hil_orderchecklistid").Value;

                }

                Entity _oclEntity = service.Retrieve(ocl.LogicalName, ocl.Id, new ColumnSet("hil_despatchpoint"));

                EntityReference plant = _oclEntity.GetAttributeValue<EntityReference>("hil_despatchpoint");

                toTeamMembers = Helper.retriveTeamMembers(service, "PPC", materialGroup, department, plant, toTeamMembers);
                toTeamMembers = Helper.retriveTeamMembers(service, "Floar", materialGroup, department, plant, toTeamMembers);
                toTeamMembers = Helper.retriveTeamMembers(service, "Testing", materialGroup, department, plant, toTeamMembers);
                toTeamMembers = Helper.retriveTeamMembers(service, "CRI", materialGroup, department, plant, toTeamMembers);


                EntityCollection entTOList = new EntityCollection();

                foreach (Entity ccEntity in toTeamMembers.Entities)
                {
                    Entity entcc = new Entity("activityparty");
                    entcc["partyid"] = ccEntity.ToEntityReference();
                    entTOList.Entities.Add(entcc);
                }

                string subject = string.Empty;
                if (today)
                {
                    subject = @"3 days prior intimation notice for Scheduled readiness";
                }
                else if (!today)
                {
                    subject = @"Intimation notice for today's Scheduled readiness";
                }
                Console.WriteLine("subject Line " + subject);
                // EntityReference oaheaderregarding = OAProductCol.Entities[0].GetAttributeValue<EntityReference>("hil_oaheader");
                Helper.sendEmail(bodtText, subject, null, entTOList, entCCList, sender, service);
                Console.WriteLine("PPC mail function ended.");
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
            }
        }
        public static string PPCMailBody(EntityCollection entColPPC)
        {
            StringBuilder sbtopBody = new StringBuilder();
            StringBuilder sbtable = new StringBuilder();
            StringBuilder sbbelowBody = new StringBuilder();
            sbtopBody.Append("<Div style='width:703.9pt;margin-left:-.15pt;font-weight:bold'>Dear Team,</Div>");
            sbtopBody.Append("<Div><p style='width:703.9pt;margin-left:-.15pt;margin-bottom:5pt;font-weight:bold'>Scheduled readiness for the below orders has been given to sales team . Please ensure the readiness of Material and stock to be updated on the readiness date mentioned below.</p></Div>");
            sbtable.Append("<Div>");
            sbtable.Append("<table border=1 cellspacing=0 cellpadding=0 width=0 style='width:703.9pt;margin-left:-.15pt;border-collapse:collapse'>");
            sbtable.Append("<tr style='height:24.0pt;font-weight:bold;'>" +
                        "<th> OA Number </th>" +
                        "<th>Enquiry No</th>" +
                        "<th>Customer Name</th>" + "" +
                        "<th>Product Name(Click to open)</th>" +
                        "<th>Delivery Qty</th>" +
                        "<th>Rediness Date</th>" +
                        "<th>Inspection Status</th>" +
                        "<th>Inspection Scheduled Date</th>" +
                    "</tr>");
            foreach (Entity entCol in entColPPC.Entities)
            {
                String partialUrl = URL + entCol.LogicalName + "&id=" + entCol.Id;
                string OaNumber = entCol.GetAttributeValue<EntityReference>("hil_oaheader").Name;
                EntityReference tenderNo = (entCol.Contains("OAHeader.hil_tenderid")) ? (EntityReference)(entCol.GetAttributeValue<AliasedValue>("OAHeader.hil_tenderid").Value) : (EntityReference)(entCol.GetAttributeValue<AliasedValue>("OAHeader.hil_orderchecklistid").Value);
                string enqnumber = tenderNo.Name;
                string inspDate = string.Empty;
                string inspStatus = string.Empty;
                if (entCol.Contains("hil_inspectioncallscheduledon"))
                    inspDate = entCol.GetAttributeValue<DateTime>("hil_inspectioncallscheduledon").AddMinutes(330).ToString("dd/MMM/yyyy");
                if (entCol.Contains("hil_inspectioncallstatus"))
                    inspStatus = entCol.FormattedValues["hil_inspectioncallstatus"];
                ///////////////////////////////********************************************
                string customerName = ((EntityReference)(entCol.GetAttributeValue<AliasedValue>("OAHeader.hil_customername").Value)).Name;
                ///////////////////////////////********************************************
                string prodName = entCol.GetAttributeValue<string>("hil_productdescription");
                string deliveryDate = entCol.GetAttributeValue<DateTime>("hil_stockreadinessdateppc").AddMinutes(330).ToString("dd/MMM/yyyy");
                string deliveryQty = String.Format("{0:0.##}", entCol.GetAttributeValue<Decimal>("hil_quantity")).ToString();
                sbtable.Append("<tr style='height:24.0pt;text-align:center'>" +
                        "<td>" + OaNumber + "</td>" +
                        "<td>" + enqnumber + "</td>" + "" +
                        "<td>" + customerName + "</td>" +
                        "<td><a href='" + partialUrl + "' target='_blank'>" + prodName + "</a></td>" +
                        "<td>" + deliveryQty + "</td>" +
                        "<td>" + deliveryDate + "</td>" +
                        "<td>" + inspStatus + "</td>" +
                        "<td>" + inspDate + "</td>" +
                    "</tr>");
            }
            sbtable.Append("</table></Div>");
            sbbelowBody.Append("<Div style='margin-top:20pt;font-weight:bold'><p>Kind Regards </p></Div>");
            sbbelowBody.Append("<Div style='margin-top:20pt;font-weight:bold'><p>OMS </p></Div>");
            return sbtopBody.ToString() + sbtable.ToString() + sbbelowBody.ToString();
        }
        #endregion
        #region InspectionStatusUpadteToday...
        public static void InspectionStatusUpadteToday()
        {
            QueryExpression queryExp = new QueryExpression("hil_oaproduct");
            queryExp.ColumnSet = new ColumnSet("ownerid");
            queryExp.Distinct = true;
            queryExp.Criteria = new FilterExpression(LogicalOperator.And);
            queryExp.Criteria.AddCondition("hil_stockreadinessdateppc", ConditionOperator.On, DateTime.Today.AddDays(1));
            EntityCollection OAProductDistinctCol = service.RetrieveMultiple(queryExp);
            Console.WriteLine("record Retrived " + OAProductDistinctCol.Entities.Count);
            foreach (Entity entity in OAProductDistinctCol.Entities)
            {
                Console.WriteLine("all OA product is retrived for " + entity.GetAttributeValue<EntityReference>("ownerid").Name);
                queryExp = new QueryExpression("hil_oaproduct");
                queryExp.ColumnSet = new ColumnSet("ownerid", "hil_name", "hil_oaheader", "hil_stockreadinessdateppc", "hil_productdescription", "hil_deliverydate",
                    "hil_quantity");
                queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                queryExp.Criteria.AddCondition("hil_stockreadinessdateppc", ConditionOperator.On, DateTime.Today.AddDays(1));
                queryExp.Criteria.AddCondition("ownerid", ConditionOperator.Equal, entity.GetAttributeValue<EntityReference>("ownerid").Id);
                LinkEntity EntityA = new LinkEntity("hil_oaproduct", "hil_oaheader", "hil_oaheader", "hil_oaheaderid", JoinOperator.Inner);
                EntityA.Columns = new ColumnSet("hil_tenderid", "hil_orderchecklistid");
                EntityA.EntityAlias = "OAHeader";
                queryExp.LinkEntities.Add(EntityA);
                EntityCollection OAProductCol = service.RetrieveMultiple(queryExp);
                Console.WriteLine("User wise OA Product is retrived." + OAProductCol.Entities.Count);
                if (OAProductCol.Entities.Count > 0)
                {
                    EntityReference owner = OAProductCol.Entities[0].GetAttributeValue<EntityReference>("ownerid");
                    string bodtText = InspectionStatusMailBody(OAProductCol);
                    Entity userConfiguration = Helper.getUserConfiguartion(owner, service);
                    EntityCollection entCCList = new EntityCollection();
                    Entity entCC = new Entity("activityparty");
                    entCC["partyid"] = userConfiguration.GetAttributeValue<EntityReference>("hil_branchproducthead");
                    entCCList.Entities.Add(entCC);
                    entCC = new Entity("activityparty");
                    entCC["partyid"] = owner;
                    entCCList.Entities.Add(entCC);
                    string subject = @"Please Update Inspection";
                    Console.WriteLine("Subject " + subject);
                    EntityReference oaheaderregarding = OAProductCol.Entities[0].GetAttributeValue<EntityReference>("hil_oaheader");
                    Helper.sendEmail(bodtText, subject, oaheaderregarding, entCCList, null, sender, service);
                }
            }
            Console.WriteLine("InspectionStatusUpadteToday ended.");
        }
        public static string InspectionStatusMailBody(EntityCollection entColPPC)
        {
            StringBuilder sbtopBody = new StringBuilder();
            StringBuilder sbtable = new StringBuilder();
            StringBuilder sbbelowBody = new StringBuilder();
            sbtopBody.Append("<Div style='width:703.9pt;margin-left:-.15pt;font-weight:bold'>Dear " + entColPPC[0].GetAttributeValue<EntityReference>("ownerid").Name + ",</Div>");
            sbtopBody.Append("<Div><p style='width:703.9pt;margin-left:-.15pt;margin-bottom:5pt;font-weight:bold'>Please update the inspection status as Waived Off/Scheduled for the below orders.</p></Div>");
            sbtable.Append("<Div>");
            sbtable.Append("<table border=1 cellspacing=0 cellpadding=0 width=0 style='width:703.9pt;margin-left:-.15pt;border-collapse:collapse'>");
            sbtable.Append("<tr style='height:24.0pt;font-weight:bold;'><th> OA Number </th><th>Enquiry No</th><th>Product Name</th><th>Readiness Date</th><th>Delivery Qty</th></tr>");
            foreach (Entity entCol in entColPPC.Entities)
            {
                String partialUrl = URL + entCol.LogicalName + "&id=" + entCol.Id;
                string OaNumber = entCol.GetAttributeValue<EntityReference>("hil_oaheader").Name;
                EntityReference tenderNo = (entCol.Contains("OAHeader.hil_tenderid")) ?
                    (EntityReference)(entCol.GetAttributeValue<AliasedValue>("OAHeader.hil_tenderid").Value) :
                    (EntityReference)(entCol.GetAttributeValue<AliasedValue>("OAHeader.hil_orderchecklistid").Value);
                string enqnumber = tenderNo.Name;
                string prodName = entCol.GetAttributeValue<string>("hil_productdescription");
                string deliveryDate = entCol.GetAttributeValue<DateTime>("hil_stockreadinessdateppc").ToString("dd/MMM/yyyy");
                string deliveryQty = String.Format("{0:0.##}", entCol.GetAttributeValue<Decimal>("hil_quantity")).ToString();
                sbtable.Append("<tr style='height:24.0pt;text-align:center'><td>" + OaNumber + "</td><td>" + enqnumber + "</td><td><a href='" + partialUrl + "' target='_blank'>" + prodName + "</a></td><td>" + deliveryDate + "</td><td>" + deliveryQty + "</td></tr>");
            }
            sbtable.Append("</table></Div>");
            sbbelowBody.Append("<Div style='margin-top:20pt;font-weight:bold'><p>Kind Regards </p></Div>");
            sbbelowBody.Append("<Div style='margin-top:20pt;font-weight:bold'><p>OMS </p></Div>");
            return sbtopBody.ToString() + sbtable.ToString() + sbbelowBody.ToString();
        }
        #endregion
        #region BillingDocument
        public static void BillingDocumentReviewbefore12days()
        {
            try
            {
                DateTime today = DateTime.Today.AddMinutes(330);

                QueryExpression queryExpOA = new QueryExpression("hil_oaheader");
                queryExpOA.ColumnSet = new ColumnSet("ownerid", "hil_oaheaderid", "hil_name");
                queryExpOA.Distinct = true;
                queryExpOA.Criteria.AddCondition("hil_dispatchdate", ConditionOperator.On, today.AddDays(12));
                LinkEntity EntityA = new LinkEntity("hil_oaheder", "hil_tender", "hil_tenderid", "hil_tenderid", JoinOperator.Inner);
                EntityA.Columns = new ColumnSet("hil_rm");
                EntityA.EntityAlias = "Pterms";
                queryExpOA.LinkEntities.Add(EntityA);
                EntityCollection OAHeaderCol = service.RetrieveMultiple(queryExpOA);
                foreach (Entity entity in OAHeaderCol.Entities)
                {

                    QueryExpression queryExp = new QueryExpression("hil_attachment");
                    queryExp.ColumnSet = new ColumnSet("hil_docurl");
                    queryExp.Distinct = true;
                    queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                    queryExp.Criteria.AddCondition("regardingobjectid", ConditionOperator.Equal, entity.GetAttributeValue<Guid>("hil_oaheaderid"));
                    queryExp.Criteria.AddCondition("hil_documenttype", ConditionOperator.Equal, "11a4ee13-6732-ec11-b6e6-002248d4cad3");
                    queryExp.Criteria.AddCondition("hil_isdeleted", ConditionOperator.Equal, false);
                    EntityCollection OAHeaderAtt = service.RetrieveMultiple(queryExp);
                    if (OAHeaderAtt.Entities.Count > 0)
                    {
                        Console.WriteLine("Billing Document is attached.");
                    }
                    else
                    {
                        ReviewBillingDocument(entity);
                    }
                }

            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
            }
        }
        public static void ReviewBillingDocument(Entity OaHeaderEntity)
        {
            //To: Sales Rep/ RM CC: ZM
            try
            {
                EntityReference owner = OaHeaderEntity.GetAttributeValue<EntityReference>("ownerid");
                string bodtText = BillingDocMailBody(OaHeaderEntity);
                Entity userConfiguration = Helper.getUserConfiguartion(owner, service);
                EntityCollection entToList = new EntityCollection();
                Entity entTo = new Entity("activityparty");
                entTo["partyid"] = owner;
                entToList.Entities.Add(entTo);
                entTo = new Entity("activityparty");
                entTo["partyid"] = (EntityReference)OaHeaderEntity.GetAttributeValue<AliasedValue>("Pterms.hil_rm").Value;
                entToList.Entities.Add(entTo);


                EntityCollection entCCList = new EntityCollection();
                Entity entCC = new Entity("activityparty");
                entCC["partyid"] = userConfiguration.GetAttributeValue<EntityReference>("hil_zonalhead");
                entCCList.Entities.Add(entCC);
                string subject = @"Please insure to submit the billing documents billed by OA number ";
                EntityReference oaheaderregarding = OaHeaderEntity.ToEntityReference();
                Helper.sendEmail(bodtText, subject, oaheaderregarding, entToList, entCCList, sender, service);
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
            }
        }
        public static string BillingDocMailBody(Entity entCol)
        {

            StringBuilder sbtopBody = new StringBuilder();
            StringBuilder sbtable = new StringBuilder();
            StringBuilder sbbelowBody = new StringBuilder();
            sbtopBody.Append("<Div style='width:700pt;margin-left:-.15pt;font-weight:bold'>Dear Sales,</Div>");
            sbtable.Append("<Div><p>Dispatch Documents for the billing took place vide below OA number is pending to attach. Please ensure to attach the same ASAP </p></Div>");
            sbtable.Append("<Div>");
            sbtable.Append("<ul>");
            String partialUrl = URL + entCol.LogicalName + "&id=" + entCol.Id;
            string OaNumber = entCol.GetAttributeValue<string>("hil_name");
            sbtable.Append("<li style='height:24.0pt;'>" +
            "<a href='" + partialUrl + "' target='_blank'>" + OaNumber + "</a></li>");
            sbtable.Append("</ul></Div>");
            sbbelowBody.Append("<Div style='margin-top:20pt;font-weight:bold'><p>Kind Regards </p></Div>");
            sbbelowBody.Append("<Div style='margin-top:20pt;font-weight:bold'><p>OMS </p></Div>");
            return sbtopBody.ToString() + sbtable.ToString() + sbbelowBody.ToString();
        }

        #endregion BillingDocument
        #region LCDocument
        public static void LCDocumentReview5DaysBefore()
        {
            try
            {
                DateTime today = DateTime.Today.AddMinutes(330);
                QueryExpression queryExpOwner = new QueryExpression("hil_oaheader");
                queryExpOwner.ColumnSet = new ColumnSet("ownerid");
                queryExpOwner.Distinct = true;
                queryExpOwner.Criteria = new FilterExpression(LogicalOperator.And);
                queryExpOwner.Criteria.AddCondition("hil_dispatchdate", ConditionOperator.On, today.AddDays(5));
                LinkEntity EntityAOwner = new LinkEntity("hil_oaheader", "hil_orderchecklist", "hil_orderchecklistid", "hil_orderchecklistid", JoinOperator.Inner);
                EntityAOwner.Columns = new ColumnSet(false);
                EntityAOwner.LinkCriteria.AddCondition("hil_paymentterms", ConditionOperator.Equal, 1);
                EntityAOwner.EntityAlias = "Pterms";
                queryExpOwner.LinkEntities.Add(EntityAOwner);
                EntityCollection OAHeaderColOwner = service.RetrieveMultiple(queryExpOwner);
                foreach (Entity entity in OAHeaderColOwner.Entities)
                {
                    QueryExpression queryExp = new QueryExpression("hil_oaheader");
                    queryExp.ColumnSet = new ColumnSet("ownerid", "hil_name", "hil_tenderid", "hil_omsemailersetup", "hil_oaheaderid", "hil_department");
                    queryExp.Distinct = true;
                    queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                    queryExp.Criteria.AddCondition("hil_dispatchdate", ConditionOperator.On, today.AddDays(5));
                    queryExp.Criteria.AddCondition("ownerid", ConditionOperator.Equal, entity.GetAttributeValue<EntityReference>("ownerid").Id);
                    LinkEntity EntityA = new LinkEntity("hil_oaheader", "hil_orderchecklist", "hil_orderchecklistid", "hil_orderchecklistid", JoinOperator.Inner);
                    EntityA.Columns = new ColumnSet(false);
                    EntityA.LinkCriteria.AddCondition("hil_paymentterms", ConditionOperator.Equal, 1);
                    EntityA.EntityAlias = "Pterms";
                    queryExp.LinkEntities.Add(EntityA);
                    EntityCollection OAHeaderCol = service.RetrieveMultiple(queryExp);
                    if (OAHeaderCol.Entities.Count > 0)
                    {
                        ReviewMailLC(OAHeaderCol);
                    }
                }

            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
            }
        }
        public static string LCDocumentMailBody(EntityCollection entColLC)
        {
            StringBuilder sbtopBody = new StringBuilder();
            StringBuilder sbtable = new StringBuilder();
            StringBuilder sbbelowBody = new StringBuilder();
            sbtopBody.Append("<Div style='width:200pt;margin-left:-.15pt;font-weight:bold'>Dear Sales/CMT/Dispatch Team,</Div>");
            sbtable.Append("<Div><p>Please ensure, the conditions of required documents has been fulfilled in line with the Terms of LC for below OA# .</p></Div>");
            sbtable.Append("<Div>");
            sbtable.Append("<ul>");
            foreach (Entity entCol in entColLC.Entities)
            {
                String partialUrl = URL + entCol.LogicalName + "&id=" + entCol.Id;
                string OaNumber = entCol.GetAttributeValue<string>("hil_name");
                sbtable.Append("<li style='height:24.0pt;'>" +
                        "<a href='" + partialUrl + "' target='_blank'>" + OaNumber + "</a></li>");
            }
            sbtable.Append("</ul>");
            sbtable.Append("</Div>");
            sbbelowBody.Append("<Div style='margin-top:20pt;font-weight:bold'><p>Kind Regards </p></Div>");
            sbbelowBody.Append("<Div style='margin-top:20pt;font-weight:bold'><p>OMS </p></Div>");
            return sbtopBody.ToString() + sbtable.ToString() + sbbelowBody.ToString();
        }
        public static void ReviewMailLC(EntityCollection Oacollection)
        {
            //Sales Rep, CMT, Dispatch

            try
            {
                EntityReference owner = Oacollection.Entities[0].GetAttributeValue<EntityReference>("ownerid");
                string bodtText = LCDocumentMailBody(Oacollection);

                EntityCollection toTeamMembers = new EntityCollection();

                EntityReference department = Oacollection[0].GetAttributeValue<EntityReference>("hil_department");
                toTeamMembers = Helper.retriveTeamMembers(service, "CMT", null, department, null, toTeamMembers);

                EntityCollection entTOList = new EntityCollection();

                foreach (Entity ccEntity in toTeamMembers.Entities)
                {
                    Entity entcc = new Entity("activityparty");
                    entcc["partyid"] = ccEntity.ToEntityReference();
                    entTOList.Entities.Add(entcc);
                }

                Entity entTO = new Entity("activityparty");
                entTO["partyid"] = owner;
                entTOList.Entities.Add(entTO);

                string subject = @"Required documents in Terms of LC";
                EntityReference oaheaderregarding = Oacollection.Entities[0].ToEntityReference();
                Helper.sendEmail(bodtText, subject, oaheaderregarding, entTOList, null, sender, service);
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
            }
        }
        #endregion LCDocument
        #region CreditDocument
        public static void CreditLimitDocumentReview10DaysBefore()
        {
            try
            {
                DateTime today = DateTime.Today.AddMinutes(330);

                QueryExpression queryExpOwner = new QueryExpression("hil_oaheader");
                queryExpOwner.ColumnSet = new ColumnSet("ownerid");
                queryExpOwner.Distinct = true;
                queryExpOwner.Criteria = new FilterExpression(LogicalOperator.And);
                queryExpOwner.Criteria.AddCondition("hil_dispatchdate", ConditionOperator.On, today.AddDays(10));
                LinkEntity EntityAOwner = new LinkEntity("hil_oaheader", "hil_orderchecklist", "hil_orderchecklistid", "hil_orderchecklistid", JoinOperator.Inner);
                EntityAOwner.Columns = new ColumnSet(false);
                EntityAOwner.LinkCriteria.AddCondition("hil_paymentterms", ConditionOperator.Equal, 8);
                EntityAOwner.EntityAlias = "Ptermss";
                LinkEntity EntityRM = new LinkEntity("hil_oaheder", "hil_tender", "hil_tenderid", "hil_tenderid", JoinOperator.Inner);
                EntityRM.Columns = new ColumnSet("hil_rm");
                EntityRM.EntityAlias = "Pterms";
                queryExpOwner.LinkEntities.Add(EntityAOwner);
                queryExpOwner.LinkEntities.Add(EntityRM);
                EntityCollection OAHeaderColowner = service.RetrieveMultiple(queryExpOwner);
                foreach (Entity entity in OAHeaderColowner.Entities)
                {
                    QueryExpression queryExp = new QueryExpression("hil_oaheader");
                    queryExp.ColumnSet = new ColumnSet("ownerid", "hil_name", "hil_tenderid", "hil_omsemailersetup", "hil_oaheaderid");
                    queryExp.Distinct = true;
                    queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                    queryExp.Criteria.AddCondition("hil_dispatchdate", ConditionOperator.On, today.AddDays(10));
                    queryExp.Criteria.AddCondition("ownerid", ConditionOperator.Equal, entity.GetAttributeValue<EntityReference>("ownerid").Id);
                    LinkEntity EntityA = new LinkEntity("hil_oaheader", "hil_orderchecklist", "hil_orderchecklistid", "hil_orderchecklistid", JoinOperator.Inner);
                    EntityA.Columns = new ColumnSet(false);
                    EntityA.LinkCriteria.AddCondition("hil_paymentterms", ConditionOperator.Equal, 8);
                    EntityA.EntityAlias = "Ptermss";
                    LinkEntity EntityRMM = new LinkEntity("hil_oaheder", "hil_tender", "hil_tenderid", "hil_tenderid", JoinOperator.Inner);
                    EntityRMM.Columns = new ColumnSet("hil_rm");
                    EntityRMM.EntityAlias = "Pterms";
                    queryExp.LinkEntities.Add(EntityA);
                    queryExp.LinkEntities.Add(EntityRMM);
                    EntityCollection OAHeaderCol = service.RetrieveMultiple(queryExp);
                    if (OAHeaderCol.Entities.Count > 0)
                    {
                        ReviewMailCredit(OAHeaderCol);
                    }
                }


            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
            }
        }

        public static void ReviewMailCredit(EntityCollection Oacollection)
        {
            // To: Sales Rep/ RM CC: ZM
            try
            {
                EntityReference owner = Oacollection.Entities[0].GetAttributeValue<EntityReference>("ownerid");
                string bodtText = CreditDocumentMailBody(Oacollection);
                Entity userConfiguration = Helper.getUserConfiguartion(owner,service);
                EntityCollection entToList = new EntityCollection();
                Entity entTo = new Entity("activityparty");
                entTo["partyid"] = owner;
                entToList.Entities.Add(entTo);
                entTo = new Entity("activityparty");
                entTo["partyid"] = (EntityReference)Oacollection.Entities[0].GetAttributeValue<AliasedValue>("Pterms.hil_rm").Value;
                entToList.Entities.Add(entTo);

                EntityCollection entCCList = new EntityCollection();
                Entity entCC = new Entity("activityparty");
                entCC["partyid"] = userConfiguration.GetAttributeValue<EntityReference>("hil_zonalhead");
                entCCList.Entities.Add(entCC);
                string subject = @"Please insure to submit the billing documents.";
                EntityReference oaheaderregarding = Oacollection.Entities[0].ToEntityReference();
                Helper.sendEmail(bodtText, subject, oaheaderregarding, entToList, entCCList, sender,service);


            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
            }
        }

        public static string CreditDocumentMailBody(EntityCollection entColLC)
        {
            StringBuilder sbtopBody = new StringBuilder();
            StringBuilder sbtable = new StringBuilder();
            StringBuilder sbbelowBody = new StringBuilder();
            sbtopBody.Append("<Div style='width:200pt;margin-left:-.15pt;font-weight:bold'>Dear Sales Team,</Div>");
            sbtable.Append("<Div><p>Please insure to submit the billing documents billed by OA number without any further delay. for below OA# .</p></Div>");
            sbtable.Append("<Div>");
            sbtable.Append("<ul>");

            foreach (Entity entCol in entColLC.Entities)
            {
                String partialUrl = URL + entCol.LogicalName + "&id=" + entCol.Id;
                string OaNumber = entCol.GetAttributeValue<string>("hil_name");
                sbtable.Append("<li style='height:24.0pt;'>" +
                        "<a href='" + partialUrl + "' target='_blank'>" + OaNumber + "</a></li>");
            }
            sbtable.Append("</ul></Div>");
            sbbelowBody.Append("<Div style='margin-top:20pt;font-weight:bold'><p>Kind Regards </p></Div>");
            sbbelowBody.Append("<Div style='margin-top:20pt;font-weight:bold'><p>OMS </p></Div>");
            return sbtopBody.ToString() + sbtable.ToString() + sbbelowBody.ToString();
        }
        #endregion CreditDocument
    }
}
