using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Havells.CRM.WebJob.BusinessManagement.AutoMailer
{
    public class OCLDeactivate
    {
        public static void getAllDepartment(IOrganizationService service)
        {
            QueryExpression query = new QueryExpression("hil_enquirydepartment");
            query.ColumnSet = new ColumnSet("hil_ocllifecycleindays");
            query.Criteria = new FilterExpression(LogicalOperator.And);
            query.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));

            EntityCollection _departmentCollection = service.RetrieveMultiple(query);
            foreach (Entity departmetn in _departmentCollection.Entities)
            {
                int days = departmetn.GetAttributeValue<int>("hil_ocllifecycleindays");
                getOCLForIntimation(service, (days - 5), departmetn.Id);
                getOCLForDeactivaition(service, days, departmetn.Id);
            }
        }
        public static void getOCLForIntimation(IOrganizationService service, int days, Guid departmentId)
        {
            string fetch = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                              <entity name='hil_orderchecklist'>
                                <attribute name='hil_orderchecklistid' />
                                <attribute name='hil_name' />
                                <attribute name='createdon' />
                                <attribute name='hil_nameofclientcustomercode' />
                                <attribute name='hil_rm' />
                                <attribute name='hil_zonalhead' />
                                <attribute name='hil_projectname' />
                                <attribute name='ownerid' />
                                <order attribute='ownerid' descending='false' />
                                <filter type='and'>
                                  <condition attribute='createdon' operator='olderthan-x-days' value='{days}' />
                                  <condition attribute='hil_approvalstatus' operator='ne' value='1' />
                                  <condition attribute='statecode' operator='eq' value='0' />
                                  <condition attribute='hil_department' operator='eq' uiname='Cable' uitype='hil_enquirydepartment' value='{departmentId}' />
                                </filter>
                              </entity>
                            </fetch>";
            EntityReference owner = null;
            EntityReference OCLRef = null;
            EntityReference ZonalHead = null;
            EntityReference RM = null;
            string table = null;
            EntityCollection _OCLCollection = service.RetrieveMultiple(new FetchExpression(fetch));
            foreach (Entity ocl in _OCLCollection.Entities)
            {
                if (owner == null)
                {
                    owner = ocl.GetAttributeValue<EntityReference>("ownerid");
                    ZonalHead = ocl.GetAttributeValue<EntityReference>("hil_zonalhead");
                    RM = ocl.GetAttributeValue<EntityReference>("hil_rm");
                    OCLRef = ocl.ToEntityReference();
                    table = table + $@"<tr>
                                        <td
                                            style='background-color:#cccccc; border-bottom:1px solid #666666; border-left:1px solid #666666; border-right:1px solid #666666; border-top:none; padding:0cm 7px 0cm 7px; vertical-align:top; width:132px'>
                                            {ocl.GetAttributeValue<string>("hil_name")}</td>
                                        <td
                                            style='background-color:#cccccc; border-bottom:1px solid #666666; border-left:none; border-right:1px solid #666666; border-top:none; padding:0cm 7px 0cm 7px; vertical-align:top; width:142px'>
                                            {ocl.GetAttributeValue<EntityReference>("hil_nameofclientcustomercode").Name}</td>
                                        <td
                                            style='background-color:#cccccc; border-bottom:1px solid #666666; border-left:none; border-right:1px solid #666666; border-top:none; padding:0cm 7px 0cm 7px; vertical-align:top; width:327px'>
                                             {ocl.GetAttributeValue<string>("hil_projectname")}</td>
                                    </tr>";
                }
               else if (owner.Id == ocl.GetAttributeValue<EntityReference>("ownerid").Id)
                {
                    OCLRef = ocl.ToEntityReference();
                    table = table + $@"<tr>
                                        <td
                                            style='background-color:#cccccc; border-bottom:1px solid #666666; border-left:1px solid #666666; border-right:1px solid #666666; border-top:none; padding:0cm 7px 0cm 7px; vertical-align:top; width:132px'>
                                            {ocl.GetAttributeValue<string>("hil_name")}</td>
                                        <td
                                            style='background-color:#cccccc; border-bottom:1px solid #666666; border-left:none; border-right:1px solid #666666; border-top:none; padding:0cm 7px 0cm 7px; vertical-align:top; width:142px'>
                                            {ocl.GetAttributeValue<EntityReference>("hil_nameofclientcustomercode").Name}</td>
                                        <td
                                            style='background-color:#cccccc; border-bottom:1px solid #666666; border-left:none; border-right:1px solid #666666; border-top:none; padding:0cm 7px 0cm 7px; vertical-align:top; width:327px'>
                                             {ocl.GetAttributeValue<string>("hil_projectname")}</td>
                                    </tr>";
                }
                else
                {
                    CreateMailBody(service, table, owner, ZonalHead, RM, 2);
                    table = "";
                    OCLRef = ocl.ToEntityReference();
                    owner = ocl.GetAttributeValue<EntityReference>("ownerid");
                    ZonalHead = ocl.GetAttributeValue<EntityReference>("hil_zonalhead");
                    RM = ocl.GetAttributeValue<EntityReference>("hil_rm");
                    table = table + $@"<tr>
                                        <td
                                            style='background-color:#cccccc; border-bottom:1px solid #666666; border-left:1px solid #666666; border-right:1px solid #666666; border-top:none; padding:0cm 7px 0cm 7px; vertical-align:top; width:132px'>
                                            {ocl.GetAttributeValue<string>("hil_name")}</td>
                                        <td
                                            style='background-color:#cccccc; border-bottom:1px solid #666666; border-left:none; border-right:1px solid #666666; border-top:none; padding:0cm 7px 0cm 7px; vertical-align:top; width:142px'>
                                            {ocl.GetAttributeValue<EntityReference>("hil_nameofclientcustomercode").Name}</td>
                                        <td
                                            style='background-color:#cccccc; border-bottom:1px solid #666666; border-left:none; border-right:1px solid #666666; border-top:none; padding:0cm 7px 0cm 7px; vertical-align:top; width:327px'>
                                             {ocl.GetAttributeValue<string>("hil_projectname")}</td>
                                    </tr>";
                }
            }
        }
        public static void getOCLForDeactivaition(IOrganizationService service, int days, Guid departmentId)
        {
            string fetch = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                              <entity name='hil_orderchecklist'>
                                <attribute name='hil_orderchecklistid' />
                                <attribute name='hil_name' />
                                <attribute name='createdon' />
                                <attribute name='hil_nameofclientcustomercode' />
                                <attribute name='hil_rm' />
                                <attribute name='hil_zonalhead' />
                                <attribute name='hil_projectname' />
                                <attribute name='ownerid' />
                                <order attribute='ownerid' descending='false' />
                                <filter type='and'>
                                  <condition attribute='createdon' operator='olderthan-x-days' value='{days}' />
                                  <condition attribute='hil_approvalstatus' operator='ne' value='1' />
                                  <condition attribute='statecode' operator='eq' value='0' />
                                  <condition attribute='hil_department' operator='eq' uiname='Cable' uitype='hil_enquirydepartment' value='{departmentId}' />
                                </filter>
                              </entity>
                            </fetch>";
            EntityReference owner = null;
            EntityReference OCLRef = null;
            EntityReference ZonalHead = null;
            EntityReference RM = null;
            string table = null;
            EntityCollection _OCLCollection = service.RetrieveMultiple(new FetchExpression(fetch));
            foreach (Entity ocl in _OCLCollection.Entities)
            {
                if (owner == null)
                {
                    owner = ocl.GetAttributeValue<EntityReference>("ownerid");
                    ZonalHead = ocl.GetAttributeValue<EntityReference>("hil_zonalhead");
                    RM = ocl.GetAttributeValue<EntityReference>("hil_rm");
                    OCLRef = ocl.ToEntityReference();
                    deactivateOCL(service, ocl.ToEntityReference());
                    table = table + $@"<tr>
                                        <td
                                            style='background-color:#cccccc; border-bottom:1px solid #666666; border-left:1px solid #666666; border-right:1px solid #666666; border-top:none; padding:0cm 7px 0cm 7px; vertical-align:top; width:132px'>
                                            {ocl.GetAttributeValue<string>("hil_name")}</td>
                                        <td
                                            style='background-color:#cccccc; border-bottom:1px solid #666666; border-left:none; border-right:1px solid #666666; border-top:none; padding:0cm 7px 0cm 7px; vertical-align:top; width:142px'>
                                            {ocl.GetAttributeValue<EntityReference>("hil_nameofclientcustomercode").Name}</td>
                                        <td
                                            style='background-color:#cccccc; border-bottom:1px solid #666666; border-left:none; border-right:1px solid #666666; border-top:none; padding:0cm 7px 0cm 7px; vertical-align:top; width:327px'>
                                             {ocl.GetAttributeValue<string>("hil_projectname")}</td>
                                    </tr>";
                }
                if (owner.Id == ocl.GetAttributeValue<EntityReference>("ownerid").Id)
                {
                    OCLRef = ocl.ToEntityReference();
                    deactivateOCL(service, ocl.ToEntityReference());
                    table = table + $@"<tr>
                                        <td
                                            style='background-color:#cccccc; border-bottom:1px solid #666666; border-left:1px solid #666666; border-right:1px solid #666666; border-top:none; padding:0cm 7px 0cm 7px; vertical-align:top; width:132px'>
                                            {ocl.GetAttributeValue<string>("hil_name")}</td>
                                        <td
                                            style='background-color:#cccccc; border-bottom:1px solid #666666; border-left:none; border-right:1px solid #666666; border-top:none; padding:0cm 7px 0cm 7px; vertical-align:top; width:142px'>
                                            {ocl.GetAttributeValue<EntityReference>("hil_nameofclientcustomercode").Name}</td>
                                        <td
                                            style='background-color:#cccccc; border-bottom:1px solid #666666; border-left:none; border-right:1px solid #666666; border-top:none; padding:0cm 7px 0cm 7px; vertical-align:top; width:327px'>
                                             {ocl.GetAttributeValue<string>("hil_projectname")}</td>
                                    </tr>";
                }
                else
                {
                    CreateMailBody(service, table, owner, ZonalHead, RM, 1);
                    table = "";
                    OCLRef = ocl.ToEntityReference();
                    owner = ocl.GetAttributeValue<EntityReference>("ownerid");
                    ZonalHead = ocl.GetAttributeValue<EntityReference>("hil_zonalhead");
                    RM = ocl.GetAttributeValue<EntityReference>("hil_rm");
                    deactivateOCL(service, ocl.ToEntityReference());
                    table = table + $@"<tr>
                                        <td
                                            style='background-color:#cccccc; border-bottom:1px solid #666666; border-left:1px solid #666666; border-right:1px solid #666666; border-top:none; padding:0cm 7px 0cm 7px; vertical-align:top; width:132px'>
                                            {ocl.GetAttributeValue<string>("hil_name")}</td>
                                        <td
                                            style='background-color:#cccccc; border-bottom:1px solid #666666; border-left:none; border-right:1px solid #666666; border-top:none; padding:0cm 7px 0cm 7px; vertical-align:top; width:142px'>
                                            {ocl.GetAttributeValue<EntityReference>("hil_nameofclientcustomercode").Name}</td>
                                        <td
                                            style='background-color:#cccccc; border-bottom:1px solid #666666; border-left:none; border-right:1px solid #666666; border-top:none; padding:0cm 7px 0cm 7px; vertical-align:top; width:327px'>
                                             {ocl.GetAttributeValue<string>("hil_projectname")}</td>
                                    </tr>";
                }
            }
        }
        static void CreateMailBody(IOrganizationService service, string table, EntityReference oclOwner, EntityReference ZonalHead, EntityReference RM, int msg)
        {

            string bodtText = mailBody(msg).Replace("{#Table#}", table);
            string subject = "OCL is Deactivated";
            EntityCollection entTOList = new EntityCollection();
            Entity entcc = new Entity("activityparty");
            entcc["partyid"] = oclOwner;
            entTOList.Entities.Add(entcc);
            EntityCollection entCCList = new EntityCollection();

            entcc = new Entity("activityparty");
            entcc["partyid"] = RM;
            entCCList.Entities.Add(entcc);

            entcc = new Entity("activityparty");
            entcc["partyid"] = ZonalHead;
            entCCList.Entities.Add(entcc);


            EntityReference sender = Helper.getSender("OMS", service);

            Helper.sendEmail(bodtText, subject, null, entTOList, entCCList, sender, service);
        }
        static void deactivateOCL(IOrganizationService service, EntityReference OclId)
        {
            SetStateRequest state = new SetStateRequest();
            state.State = new OptionSetValue(1);
            state.Status = new OptionSetValue(2);
            state.EntityMoniker = new EntityReference(OclId.LogicalName, OclId.Id);
            SetStateResponse stateSet = (SetStateResponse)service.Execute(state);

            QueryExpression query = new QueryExpression("hil_orderchecklistproduct");
            query.ColumnSet = new ColumnSet(false);
            query.Criteria = new FilterExpression(LogicalOperator.And);
            query.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
            query.Criteria.AddCondition(new ConditionExpression("hil_orderchecklistid", ConditionOperator.Equal, OclId.Id));

            EntityCollection _OCLPrdCollection = service.RetrieveMultiple(query);
            foreach (Entity OCL in _OCLPrdCollection.Entities)
            {
                state = new SetStateRequest();
                state.State = new OptionSetValue(1);
                state.Status = new OptionSetValue(2);
                state.EntityMoniker = new EntityReference(OCL.LogicalName, OCL.Id);
                stateSet = (SetStateResponse)service.Execute(state);
            }
        }
        public static string mailBody(int msg)
        {
            if (msg == 1)
            {
                string body = @"<div data-wrapper='true' style='font-size:9pt;font-family:' Segoe UI','Helvetica Neue',sans-serif;'>
                                <div>
                                    <div><span style='font-size:11pt'><span style='line-height:normal'><span
                                                    style='font-family:Calibri,sans-serif'><span style='font-size:12.0pt'><span
                                                            style='font-family:&quot;Times New Roman&quot;,serif'>Dear
                                                            User,</span></span></span></span></span><br><br><span style='font-size:11pt'><span
                                                style='line-height:normal'><span style='font-family:Calibri,sans-serif'>Below OCLs, created by you
                                                    are hereby deactivated due to draft overtime in line with the OCL norms.</span></span></span>
                                    </div>
                                    <div>&nbsp;</div>
                                    <div>
                                        <table cellspacing='0' class='MsoTable15Grid4' style='border-collapse:collapse; border:none'>
                                            <tbody>
                                                <tr>
                                                    <td
                                                        style='background-color:#b54646; border-bottom:1px solid #b54646; border-left:1px solid #b54646; border-right:none; border-top:1px solid #b54646; padding:0cm 7px 0cm 7px; vertical-align:top; width:132px'>
                                                        <span style='font-size:11pt'><span style='line-height:normal'><span
                                                                    style='font-family:Calibri,sans-serif'><span style='color:white'>OCL
                                                                        No.</span></span></span></span>
                                                    </td>
                                                    <td
                                                        style='background-color:#b54646; border-bottom:1px solid #b54646; border-left:1px solid #b54646; border-right:none; border-top:1px solid #b54646; padding:0cm 7px 0cm 7px; vertical-align:top; width:132px'>
                                                        <span style='font-size:11pt'><span style='line-height:normal'><span
                                                                    style='font-family:Calibri,sans-serif'><span
                                                                        style='color:white'>Customer</span></span></span></span>
                                                    </td>
                                                    <td
                                                        style='background-color:#b54646; border-bottom:1px solid #b54646; border-left:1px solid #b54646; border-right:none; border-top:1px solid #b54646; padding:0cm 7px 0cm 7px; vertical-align:top; width:132px'>
                                                        <span style='font-size:11pt'><span style='line-height:normal'><span
                                                                    style='font-family:Calibri,sans-serif'><span style='color:white'>Project
                                                                        Name</span></span></span></span>
                                                    </td>
                                                </tr>
                                                {#Table#}
                                            </tbody>
                                        </table>
                                    </div>
                                    <table cellspacing='0' class='MsoTableGrid' style='border-collapse:collapse; border:none'>
                                        <tbody></tbody>
                                    </table><br><span style='font-size:11pt'><span style='line-height:107%'><span
                                                style='font-family:Calibri,sans-serif'>You are requested to kindly recreate the OCL and trigger for
                                                approval within 24 hours of creation.</span></span></span><br><br><span style='font-size:11pt'><span
                                            style='line-height:107%'><span
                                                style='font-family:Calibri,sans-serif'>Regards</span></span></span><br><span
                                        style='font-size:11pt'><span style='line-height:107%'><span
                                                style='font-family:Calibri,sans-serif'>OMS</span></span></span>
                                </div>
                            </div>";
                return body;
            }
            else if (msg == 2)
            {
                string body = @"<div data-wrapper='true' style='font-size:9pt;font-family:' Segoe UI','Helvetica Neue',sans-serif;'>
                                <div>
                                    <div><span style='font-size:11pt'><span style='line-height:normal'><span
                                                    style='font-family:Calibri,sans-serif'><span style='font-size:12.0pt'><span
                                                            style='font-family:&quot;Times New Roman&quot;,serif'>Dear
                                                            User,</span></span></span></span></span><br><br><span style='font-size:11pt'><span
                                                style='line-height:normal'><span style='font-family:Calibri,sans-serif'>Below OCLs, created by you are hereby deactivated after 5 days, due to draft overtime in line with the OCL norms.</span></span></span>
                                    </div>
                                    <div>&nbsp;</div>
                                    <div>
                                        <table cellspacing='0' class='MsoTable15Grid4' style='border-collapse:collapse; border:none'>
                                            <tbody>
                                                <tr>
                                                    <td
                                                        style='background-color:#b54646; border-bottom:1px solid #b54646; border-left:1px solid #b54646; border-right:none; border-top:1px solid #b54646; padding:0cm 7px 0cm 7px; vertical-align:top; width:132px'>
                                                        <span style='font-size:11pt'><span style='line-height:normal'><span
                                                                    style='font-family:Calibri,sans-serif'><span style='color:white'>OCL
                                                                        No.</span></span></span></span>
                                                    </td>
                                                    <td
                                                        style='background-color:#b54646; border-bottom:1px solid #b54646; border-left:1px solid #b54646; border-right:none; border-top:1px solid #b54646; padding:0cm 7px 0cm 7px; vertical-align:top; width:132px'>
                                                        <span style='font-size:11pt'><span style='line-height:normal'><span
                                                                    style='font-family:Calibri,sans-serif'><span
                                                                        style='color:white'>Customer</span></span></span></span>
                                                    </td>
                                                    <td
                                                        style='background-color:#b54646; border-bottom:1px solid #b54646; border-left:1px solid #b54646; border-right:none; border-top:1px solid #b54646; padding:0cm 7px 0cm 7px; vertical-align:top; width:132px'>
                                                        <span style='font-size:11pt'><span style='line-height:normal'><span
                                                                    style='font-family:Calibri,sans-serif'><span style='color:white'>Project
                                                                        Name</span></span></span></span>
                                                    </td>
                                                </tr>
                                                {#Table#}
                                            </tbody>
                                        </table>
                                    </div>
                                    <table cellspacing='0' class='MsoTableGrid' style='border-collapse:collapse; border:none'>
                                        <tbody></tbody>
                                    </table><br><span style='font-size:11pt'><span style='line-height:107%'><span
                                                style='font-family:Calibri,sans-serif'>You are requested to kindly process these OCLs and trigger them for approval within 24 hours.</span></span></span><br><br><span style='font-size:11pt'><span
                                            style='line-height:107%'><span
                                                style='font-family:Calibri,sans-serif'>Regards</span></span></span><br><span
                                        style='font-size:11pt'><span style='line-height:107%'><span
                                                style='font-family:Calibri,sans-serif'>OMS</span></span></span>
                                </div>
                            </div>";
                return body;
            }
            else
                return null;
        }
    }
}
