using Microsoft.Crm.Sdk.Messages;
using Microsoft.Extensions.DependencyModel;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreateFieldsFromExcelToD365
{
    public static class ProcessMigration
    {
        public static void MainFunction(IOrganizationService service, IOrganizationService serviceSer)
        {
            String processName = "Accept Proposed Booking;AcceptTeamRecommendation;ActionBomTemplateLineItem;AddMembersBatch;CapacityProfilesToUsers;Claim Operation;ConsumerApp_ChangePassword;ConsumerApp_ForgetPassword;ConsumerApp_ProductRegistration;ConsumerApp_ProfileUpdate;ConsumerApp_RegisterCustomer;ConsumerApp_Validate;ConsumerApp_WorkOrderDemo;ConsumerApp_WorkOrderService;Copy Managing Partner from parent account;Create Approvals;Create Task Action;CustomerCreation;CustomerOutstanding;CustomiseRibbon;Project Service - Apply work template (Deprecated in v3.0);Project Service - AssignGenericResource (Deprecated in v3.0);Project Service - AutoGenerateProjectTeam (Deprecated in v3.0);Project Service - Fulfill Resource Demand (Deprecated in v3.0);Project Service - Get Resource Availability (Deprecated in v3.0);Project Service - Get Resource Demand TimeLine (Deprecated in v3.0);Project Service - GetCollectionData (Deprecated in v3.0);Project Service - GetGenericResourceDetails (Deprecated in v3.0);Project Service - GetMyChangedSkills (Deprecated in v3.0);Project Service - GetResourceAvailabilitySummary (Deprecated in v3.0);Project Service - GetResourceBookingByProject (Deprecated in v3.0);Project Service - GetTimelineData (Deprecated in v3.0);Project Service - LogFindWorkEvent (Deprecated in v3.0);Project Service - MSProject_GetFindResourcesURL (Deprecated in v3.4);Project Service - Project Team Member Sign up process accept (Deprecated in v3.0);Project Service - Project Team Member Sign-Up Process (Deprecated in v3.0);Project Service - Project Team Update Membership Status (Deprecated in v3.0);Project Service - ResGetResourceDetail (Deprecated in v3.0);Project Service - ResourceReservationCancel (Deprecated in v3.0);Project Service - ResourceUtilization (Deprecated in v3.0);Project Service - ResourceUtilizationChart (Deprecated in v3.0);Project Service - UpdateChangedSkills (Deprecated in v3.0);Sales Insights Get Sequence Analytics KPIs;Send email Action (TCR);Send Email to Try team;Send Sms Confirmation to Contact;Update Final Status;Update Inspection Status";
            //"Account mandatory for Franchisee/DSE/Technician;AMC Plan Visibility;Amount zero validation;Approved Data Sheet (GTP);Asset Approval Lock;Asset Invoice Details;Bank Guarantee approval;Base charge validation;Business rule for PR header form;CallBack Validation;Campaign Code;Campaign run by Business Rule;Checks and Validations;Close Ticket Validation;Configure By-Spare Part;Copy of New business rule;Customer Email is Recommended in AMC Jobs;Customer Feedback Validations;Data Sheet;Date range validation;Date validation;Dealer field;DG Power Management;Disable Product Request detail fields;Drum Length/Schedule: ;Email Mandatory;Emergency Remarks Mandatory;Enable/Disable De-link PO Option;Enquiry Lost Other Reason Validation;Enquiry Lost Reason Validations;Enquiry Number;Error Message on Consumer Category;Freight Charge Default set 0;Fulfill Qty -Req Qty- Stock in Hand Validation;Hide and Show B2B Consumer Name;Hide Tender No If Tender No Blank;IDU to ODU Wire Joint Condition;IDU to Plug Point Wire Joint Condition;IF Approval Status equals \"Approved\" THEN Set Cleared for Manufacturing to \"Yes\";if industry type empty then hide industry sub type;IF L.P (Rs/Mtr) equals 0 THEN Clear L.P (Rs/Mtr);IF LD PENALTY YES THEN SET LD TERMS 0.5% per week subject to maximum of 5% of undelivered Portion;IF LOADING FACTOR (in %) does not contain data THEN Set LOADING FACTOR (in %) to 0;if the type of roof is SLANT;If Type Of Roof is RCC;INSPECTION;Inspection Status;Is SAP Verification is mandatory- Product Registration date wef;Job Closure Date Validation;KKG Disposition - Create Child Job Id.;Lock Approvals if Status equals Closed;Lock CallBack Field;lock fields  based on emd required;Lock Fields if SO Number generated;lock fields on grid;Lock Product Division;Lock SAW Activity;Lock Serial Number;Lock the Message Field Based on Sms Template Data;lock Vaccination status;Lock Warranty crucial Information if Asset is Approved;lower limit;LT Penalty Validation;Make Closure remarks mandatory;Make KKG Code Required Mandatory;Make KKG Mandatory on Job Closure;Make Old Job Reference Mandatory;Make Web Closure remarks Mandatory;Mandatory comments for Job Re-assign to BSH.;Mandatory reason for Job Re-assign to BSH;Manufacturing Clearance;MCB available condition;On Acceptance status Partially;Payable Charges Calculation;Payment Terms;PIC Position;PIC User;please enter Correct Temperature;PO Rate Validation if HO Price Given;Price List Default Value;Price-condition make Required on Price Basis;Product Extensions;Product rule;QAP;QAP Cleared On;Quantity Validation;Receipt Amount mandatory in AMC Jobs;remarks mandatory;Required Dispatch Date;Re-Sync Enable;SAW Approval Remark;Set % Margin Added on HO Price to 0;Set Consumer Type and Consumer Category for old jobs and not cc jobs;Set Default Value;Set default value of DISCOUNT % to 0;Set Duration autoset when job service is Used;Set Job Service Default mandatory fields;set PO Quantity;Set Rejection Reason as Business Required;Set Required Remark field when BG approval status is rejected;Set Required Remark field when DD approval status is rejected;Set Source \"Customer Portal\" for Portal;Set Value in Consumer Category;Set Visibility of Extra Fields;Show Customer on created record;Show fields on created record;Show Hide Boe details;Show hide on DD Required;Show hide on PMD Required;show hide resend payment link;Show Marking details description on marking details checked available;Source is mandatory;Stabliser available condition;Technician Type;Threshold job count validation;Tracker code is mandatory;upper limit;Validation based on Type of BG;Visible Serial Number;Warranty Void Logic;Warranty Void Reason Validation";
            string[] processNames = processName.Split(';');
            ActivateWorkflow(processNames, serviceSer);
            Console.WriteLine("Done");
        }
        public static void craeteworkflow(IOrganizationService service, IOrganizationService serviceSer)
        {
            var query = new QueryExpression("workflow");
            query.Criteria.AddCondition("name", ConditionOperator.Equal, "AdvisoryLineSync");
            query.ColumnSet = new ColumnSet(true);
            EntityCollection entityCollection = service.RetrieveMultiple(query);
            int total = entityCollection.Entities.Count;
            int done = 0;
            int skip = 0;
            int error = 0;
            foreach (Entity e in entityCollection.Entities)
            {
                try
                {
                    Entity entity = e;

                    var query1 = new QueryExpression("workflow");
                    query1.Criteria.AddCondition("name", ConditionOperator.Equal, e.GetAttributeValue<string>("name"));
                    query1.ColumnSet = new ColumnSet(true);
                    EntityCollection entityCollection1 = serviceSer.RetrieveMultiple(query1);
                    if (entityCollection1.Entities.Count > 0)
                    {
                        //Console.WriteLine(e.FormattedValues["category"]);
                        //e["ownerid"] = Program.SystemAdmin;
                        //e["owninguser"] = Program.SystemAdmin;
                        //e["ismanaged"] = false;
                        //serviceSer.Update(e);
                        skip++;
                        try
                        {
                            if (entity.GetAttributeValue<OptionSetValue>("statuscode").Value == 2)
                            {
                                SetStateRequest req = new SetStateRequest();
                                req.EntityMoniker = new EntityReference(entity.LogicalName, entityCollection1[0].Id);
                                req.State = new OptionSetValue(1);
                                req.Status = new OptionSetValue(2);
                                serviceSer.Execute(req);
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Process with name !" + e.GetAttributeValue<string>("name") + "! Error in Update status " + ex.Message);
                        }

                        //Console.WriteLine("!!!!!!!!!!!!!!!!! Process with name " + e.GetAttributeValue<string>("name") + " of Type " + e.FormattedValues["category"] + " is Skiped " + skip + "/" + total);
                    }
                    //else
                    //{
                    //    //Console.WriteLine(e.FormattedValues["category"] + "  Name  " + e.GetAttributeValue<string>("name"));
                    //    e["ownerid"] = Program.SystemAdmin;
                    //    e["owninguser"] = Program.SystemAdmin;
                    //    e["ismanaged"] = false;
                    //    entity["ismanaged"] = false;
                    //    e["type"] = new OptionSetValue(1);
                    //    e["statuscode"] = new OptionSetValue(1);
                    //    e["statecode"] = new OptionSetValue(0);
                    //    if (e.Contains("xaml"))
                    //    {
                    //        if (e.GetAttributeValue<string>("xaml").Contains("&amp;#2352;&amp;#2369;&amp;#2346;&amp;#2351;&amp;#2366;"))
                    //        {
                    //            e["xaml"] = e.GetAttributeValue<string>("xaml").Replace("&amp;#2352;&amp;#2369;&amp;#2346;&amp;#2351;&amp;#2366;", "INR");
                    //            Console.WriteLine("d");
                    //        }
                    //    }
                    //    //Guid proc = serviceSer.Create(e);
                    //    serviceSer.Update(entity);
                    //    //service.Delete(entity.LogicalName, entity.Id);
                    //    if (entity.GetAttributeValue<string>("xaml").Contains("&amp;#2352;&amp;#2369;&amp;#2346;&amp;#2351;&amp;#2366;"))
                    //    {
                    //        //Console.WriteLine("d");
                    //    }
                    //    done++;
                    //    Console.WriteLine("~~~~~~~~~~~~~~~~~~~ Process with name " + e.GetAttributeValue<string>("name") + " of Type " + e.FormattedValues["category"] + " is completed " + done + "/" + total);
                    //}

                }
                catch (Exception ex)
                {
                    if (!ex.Message.Contains("The entity with a name ="))
                    {
                        error++;
                        Console.WriteLine("___________Failed " + error + "/" + total + " in process Name !" + e.GetAttributeValue<string>("name") + "! type " +
                            e.FormattedValues["category"] + " Error is " + ex.Message);
                    }
                    else
                        Console.WriteLine("Skip");
                }

            }

        }
        public static void UpdateWorkflow(IOrganizationService service, string processName, IOrganizationService serviceSer)
        {
            var query = new QueryExpression("workflow");
            query.Criteria.AddCondition("name", ConditionOperator.Equal, processName);
            query.ColumnSet = new ColumnSet(true);
            EntityCollection entityCollection = service.RetrieveMultiple(query);
            int total = entityCollection.Entities.Count;
            int done = 0;
            int skip = 0;
            int error = 0;
            foreach (Entity e in entityCollection.Entities)
            {
                try
                {
                    Entity entity = e;

                    var query1 = new QueryExpression("workflow");
                    query1.Criteria.AddCondition("name", ConditionOperator.Equal, e.GetAttributeValue<string>("name"));
                    query1.ColumnSet = new ColumnSet(true);
                    EntityCollection entityCollection1 = serviceSer.RetrieveMultiple(query1);
                    if (entityCollection1.Entities.Count > 0)
                    {
                        skip++;
                        try
                        {
                            if (entity.GetAttributeValue<OptionSetValue>("statuscode").Value == 2)
                            {
                                SetStateRequest req = new SetStateRequest();
                                req.EntityMoniker = new EntityReference(entity.LogicalName, entityCollection1[0].Id);
                                req.State = new OptionSetValue(1);
                                req.Status = new OptionSetValue(2);
                                serviceSer.Execute(req);
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Process with name !" + e.GetAttributeValue<string>("name") + "! Error in Update status " + ex.Message);
                        }
                    }
                }
                catch (Exception ex)
                {
                    if (!ex.Message.Contains("The entity with a name ="))
                    {
                        error++;
                        Console.WriteLine("___________Failed " + error + "/" + total + " in process Name !" + e.GetAttributeValue<string>("name") + "! type " +
                            e.FormattedValues["category"] + " Error is " + ex.Message);
                    }
                    else
                        Console.WriteLine("Skip");
                }

            }

        }
        private static void ActivateWorkflow(string[] processName, IOrganizationService serviceSer)
        {
            var query = new QueryExpression("workflow");
            query.Criteria.AddCondition("name", ConditionOperator.In, processName);
            query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
            query.ColumnSet = new ColumnSet(true);
            EntityCollection entityCollection = serviceSer.RetrieveMultiple(query);

            foreach (Entity entity in entityCollection.Entities)
            {
                try
                {
                    SetStateRequest req = new SetStateRequest();
                    req.EntityMoniker = new EntityReference(entity.LogicalName, entity.Id);
                    req.State = new OptionSetValue(1);
                    req.Status = new OptionSetValue(2);
                    serviceSer.Execute(req);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Process With Name " + entity.GetAttributeValue<string>("name"));
                }
                
            }
        }


    }
}
