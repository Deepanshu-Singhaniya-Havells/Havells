using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Configuration;
using System.Data;
using Microsoft.Office.Interop;

namespace ClaimGeneration
{
    public class Common
    {
        public static void PostClosureSAWActivityApprovals(IOrganizationService _service) {
            var _fromDate = DateTime.Now;

            string _currDate = _fromDate.Year.ToString() + "-" + _fromDate.Month.ToString().PadLeft(2, '0') + "-" + _fromDate.Day.ToString().PadLeft(2, '0');

            DateTime _processDate = Convert.ToDateTime(_currDate);
            EntityReference _erFiscalMonth = null;
            QueryExpression queryExp = new QueryExpression("hil_claimperiod");
            queryExp.ColumnSet = new ColumnSet("hil_fromdate", "hil_todate");
            queryExp.Criteria = new FilterExpression(LogicalOperator.And);
            queryExp.Criteria.AddCondition("hil_fromdate", ConditionOperator.OnOrBefore, _processDate);
            queryExp.Criteria.AddCondition("hil_todate", ConditionOperator.OnOrAfter, _processDate);
            EntityCollection entcoll = _service.RetrieveMultiple(queryExp);
            if (entcoll.Entities.Count > 0)
            {
                _erFiscalMonth = entcoll.Entities[0].ToEntityReference();
            }
            //<condition attribute='hil_fiscalmonth' operator='eq' value='{_erFiscalMonth.Id}' />
            /*
             <link-entity name='hil_jobsextension' from='hil_jobsextensionid' to='hil_jobextension' link-type='inner' alias='ab'>
                    <attribute name='hil_jobsextensionid' />
                    <filter type='and'>
                    <condition attribute='hil_isjobclosurescreeningcompleted' operator='ne' value='1' />
                    </filter>
                </link-entity>
             */
            string _fetchXML11 = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                <entity name='msdyn_workorder'>
                <attribute name='msdyn_name' />
                <attribute name='msdyn_workorderid' />
                <attribute name='msdyn_timeclosed' />
                <order attribute='msdyn_timeclosed' descending='false' />
                <filter type='and'>
                    <condition attribute='msdyn_substatus' operator='eq' value='{{1727FA6C-FA0F-E911-A94E-000D3AF060A1}}' />
                    <condition attribute='hil_closeticket' operator='eq' value='1' />
                    <condition attribute='msdyn_name' operator='eq' value='10062322279837' />
                    <condition attribute='hil_isocr' operator='ne' value='1' />
                    <condition attribute='hil_typeofassignee' operator='ne' value='{{7D1ECBAB-1208-E911-A94D-000D3AF0694E}}' />
                </filter>
                </entity>
                </fetch>";
            EntityCollection entCol11 = null;
            string _fetchXMLSAW = string.Empty;
            EntityCollection _entSAW = null;

            int i = 1;
            while (true)
            {
                entCol11 = _service.RetrieveMultiple(new FetchExpression(_fetchXML11));

                foreach (Entity entity in entCol11.Entities)
                {
                    _fetchXMLSAW = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                      <entity name='hil_sawactivity'>
                        <attribute name='hil_sawactivityid' />
                        <attribute name='hil_name' />
                        <attribute name='createdon' />
                        <order attribute='hil_name' descending='false' />
                        <filter type='and'>
                          <condition attribute='hil_sawcategory' operator='eq' value='{{B0918D74-44ED-EA11-A815-000D3AF05D7B}}' />
                          <filter type='or'>
                            <condition attribute='hil_repeatreferencejob' operator='eq' value='{entity.Id}' />
                            <condition attribute='hil_jobid' operator='eq' value='{entity.Id}' />
                          </filter>
                          <condition attribute='statecode' operator='eq' value='0' />
                        </filter>
                      </entity>
                    </fetch>";
                    _entSAW = _service.RetrieveMultiple(new FetchExpression(_fetchXMLSAW));
                    if (_entSAW.Entities.Count == 0)
                    {
                        #region Post Closure SAW Activity Approvals
                        msdyn_workorder _enJob1 = (msdyn_workorder)_service.Retrieve(msdyn_workorder.EntityLogicalName, entity.Id, new ColumnSet("msdyn_customerasset", "msdyn_timeclosed", "hil_typeofassignee", "hil_isgascharged", "hil_laborinwarranty", "hil_callsubtype", "hil_productsubcategory", "hil_customerref", "createdon", "hil_isocr", "hil_claimstatus"));
                        if (_enJob1 != null)
                        {
                            bool _isOCR = false;
                            bool _gasCharged = false;
                            bool _laborInWarranty = false;
                            bool _underReview = false;
                            OptionSetValue optVal = null;
                            EntityReference erTypeOfAssignee = null;

                            if (_enJob1.Attributes.Contains("hil_typeofassignee"))
                            {
                                erTypeOfAssignee = _enJob1.GetAttributeValue<EntityReference>("hil_typeofassignee");
                            }
                            if (_enJob1.Attributes.Contains("hil_claimstatus"))
                            {
                                optVal = _enJob1.GetAttributeValue<OptionSetValue>("hil_claimstatus");
                            }
                            if (_enJob1.Attributes.Contains("hil_isocr"))
                            {
                                _isOCR = (bool)_enJob1["hil_isocr"];
                            }
                            if (_enJob1.Attributes.Contains("hil_isgascharged"))
                            {
                                _gasCharged = (bool)_enJob1["hil_isgascharged"];
                            }
                            if (_enJob1.Attributes.Contains("hil_laborinwarranty"))
                            {
                                _laborInWarranty = (bool)_enJob1["hil_laborinwarranty"];
                            }
                            if (!_isOCR)
                            {
                                if (erTypeOfAssignee.Id != new Guid("7D1ECBAB-1208-E911-A94D-000D3AF0694E")) // Labor InWarranty and Type of Assignee !=DSE
                                {
                                    #region RepeatRepair Approval
                                    DateTime _createdOn = _enJob1.GetAttributeValue<DateTime>("createdon").AddDays(-15);
                                    DateTime _ClosedOn = DateTime.Now.AddDays(-15);
                                    string _strCreatedOn = _createdOn.Year.ToString() + "-" + _createdOn.Month.ToString() + "-" + _createdOn.Day.ToString();
                                    string _strClosedOn = _ClosedOn.Year.ToString() + "-" + _ClosedOn.Month.ToString() + "-" + _ClosedOn.Day.ToString();
                                    EntityCollection entCol;

                                    string _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                        <entity name='msdyn_workorder'>
                                        <attribute name='msdyn_name' />
                                        <attribute name='createdon' />
                                        <attribute name='hil_productsubcategory' />
                                        <attribute name='hil_customerref' />
                                        <attribute name='msdyn_customerasset' />
                                        <attribute name='hil_callsubtype' />
                                        <attribute name='msdyn_workorderid' />
                                        <attribute name='msdyn_timeclosed' />
                                        <attribute name='msdyn_closedby' />
                                        <order attribute='createdon' descending='true' />
                                        <filter type='and'>
                                            <condition attribute='hil_isocr' operator='ne' value='1' />
                                            <condition attribute='hil_typeofassignee' operator='ne' value='{7D1ECBAB-1208-E911-A94D-000D3AF0694E}' />
                                            <condition attribute='msdyn_workorderid' operator='ne' value='" + entity.Id + @"' />
                                            <condition attribute='hil_customerref' operator='eq' value='" + _enJob1.GetAttributeValue<EntityReference>("hil_customerref").Id + @"' />
                                            <condition attribute='hil_callsubtype' operator='eq' value='" + _enJob1.GetAttributeValue<EntityReference>("hil_callsubtype").Id + @"' />
                                            <condition attribute='hil_callsubtype' operator='ne' value='{8D80346B-3C0B-E911-A94E-000D3AF06CD4}' />
                                            <condition attribute='hil_productsubcategory' operator='eq' value='" + _enJob1.GetAttributeValue<EntityReference>("hil_productsubcategory").Id + @"' />
                                            <filter type='or'>
                                                <condition attribute='msdyn_timeclosed' operator='on-or-after' value='" + _strCreatedOn + @"' />
                                                <condition attribute='msdyn_timeclosed' operator='on-or-after' value='" + _strClosedOn + @"' />
                                            </filter>
                                            <condition attribute='msdyn_substatus' operator='in'>
                                            <value>{1727FA6C-FA0F-E911-A94E-000D3AF060A1}</value>
                                            <value>{2927FA6C-FA0F-E911-A94E-000D3AF060A1}</value>
                                            <value>{7E85074C-9C54-E911-A951-000D3AF0677F}</value>
                                            </condition>
                                        </filter>
                                        </entity>
                                        </fetch>";
                                    entCol = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                                    if (entCol.Entities.Count > 0)
                                    {

                                        string _remarks = string.Empty;
                                        string _remarks1 = string.Empty;
                                        EntityReference _entref = null;

                                        foreach (Entity ent in entCol.Entities)
                                        {
                                            _entref = ent.ToEntityReference();
                                            if (ent.GetAttributeValue<EntityReference>("msdyn_customerasset").Id == _enJob1.GetAttributeValue<EntityReference>("msdyn_customerasset").Id)
                                            {
                                                _remarks += ent.GetAttributeValue<string>("msdyn_name") + ",";
                                            }
                                            else
                                            {
                                                _remarks1 += ent.GetAttributeValue<string>("msdyn_name") + ",";
                                            }
                                        }
                                        _remarks = ((_remarks == string.Empty ? "" : "Repeated Jobs with Same Serial Number: " + _remarks + ":\n") + (_remarks1 == string.Empty ? "" : "Repeated Jobs with Same Product Subcategory: " + _remarks1 + ":")).Replace(",:", "") + " - Screening";
                                        CommonLib obj = new CommonLib();
                                        CommonLib objReturn = obj.CreateSAWActivity(_enJob1.Id, 0, SAWCategoryConst._RepeatRepair, _service, _remarks, _entref);
                                        if (objReturn.statusRemarks == "OK")
                                        {
                                            _underReview = true;
                                        }
                                    }
                                    #endregion

                                    if (_underReview)
                                    {
                                        if (optVal != null && optVal.Value != 3)
                                        {
                                            Entity Ticket = new Entity("msdyn_workorder");
                                            Ticket.Id = entity.Id;
                                            Ticket["hil_claimstatus"] = new OptionSetValue(1); //Claim Under Review
                                            _service.Update(Ticket);
                                        }
                                    }
                                }
                            }
                        }
                        #endregion
                    }
                    if (entity.Contains("ab.hil_jobsextensionid"))
                    {
                        Entity _entUpdate = new Entity("hil_jobsextension", new Guid(entity.GetAttributeValue<AliasedValue>("ab.hil_jobsextensionid").Value.ToString()));
                        _entUpdate["hil_isjobclosurescreeningcompleted"] = true;
                        _service.Update(_entUpdate);
                    }

                    Console.WriteLine("Processing ..." + i++.ToString());
                }
            }
        }
        public static void GenerateClaimLines(Entity ent, Guid _performaInvoiceId, IOrganizationService _service, EntityReference _erClaimPeriod)
        {
            if (ent != null)
            {
                OptionSetValue opBrand = null;
                OptionSetValue opClaimStatus = null;
                OptionSetValue opWarrantyStatus = null;
                bool _isOCR = false;
                bool _LaborInWarranty = false;
                EntityReference erProdSubCatg = null;
                EntityReference erProdCatg = null;
                EntityReference erSalesOffice = null;
                EntityReference erTypeOfAssignee = null;
                EntityReference erCallSubType = null;
                EntityReference erSchemeCode = null;
                EntityReference erTatCategory = null;
                EntityReference _erRelatedFranchise = null;
                OptionSetValue _opRelatedFranchiseCategory = null;
                OptionSetValue _opSourceOfJob = null;
                OptionSetValue opCountryClassification = null;
                int _Jobqty = 1;
                bool _isGascharged = false;
                bool _sparepartuse = false;
                bool _jobclosemobile = false;
                string kkgcode = string.Empty;
                DateTime _closedOn = new DateTime(1900, 1, 1);
                QueryExpression queryExp;
                EntityCollection entcoll;
                Guid _prodCategoryId_LLOYDAirCon = new Guid("D51EDD9D-16FA-E811-A94C-000D3AF0694E");
                Guid _callSubTypeId_Installation = new Guid("E3129D79-3C0B-E911-A94E-000D3AF06CD4");
                Guid _callSubTypeId_Breakdown = new Guid("6560565A-3C0B-E911-A94E-000D3AF06CD4");
                Guid _callSubTypeId_PMS = new Guid("E2129D79-3C0B-E911-A94E-000D3AF06CD4");
                Guid _callSubTypeId_AMC = new Guid("55A71A52-3C0B-E911-A94E-000D3AF06CD4");

                string _strJobClosedOn = string.Empty;
                string _fetchXML = string.Empty;

                if (ent.Attributes.Contains("hil_isocr"))
                {
                    _isOCR = ent.GetAttributeValue<bool>("hil_isocr");
                }
                if (!_isOCR)
                {
                    //
                    if ((ent.Attributes.Contains("hil_claimstatus") && ent.Attributes.Contains("msdyn_timeclosed") && ent.Attributes.Contains("hil_laborinwarranty") && ent.Attributes.Contains("hil_warrantystatus")) || ent.GetAttributeValue<OptionSetValue>("hil_claimstatus").Value==7)
                    {
                        #region Local Variables
                        opClaimStatus = ent.GetAttributeValue<OptionSetValue>("hil_claimstatus");
                        opBrand = ent.GetAttributeValue<OptionSetValue>("hil_brand");

                        if (ent.GetAttributeValue<OptionSetValue>("hil_claimstatus").Value == 7)
                        {
                            QueryExpression qrExp = new QueryExpression("hil_sawactivity");
                            qrExp.ColumnSet = new ColumnSet("createdon");
                            qrExp.Criteria = new FilterExpression(LogicalOperator.And);
                            qrExp.Criteria.AddCondition("hil_jobid", ConditionOperator.Equal, ent.Id);
                            qrExp.Criteria.AddCondition("hil_sawcategory", ConditionOperator.Equal, new Guid("E123BD08-8ADD-EA11-A813-000D3AF05A4B"));
                            EntityCollection entCol50 = _service.RetrieveMultiple(qrExp);
                            if (entCol50.Entities.Count > 0)
                            {
                                _closedOn = ent.GetAttributeValue<DateTime>("createdon").AddMinutes(330);
                            }
                            else {
                                _closedOn = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);
                            }
                        }
                        else {
                            _closedOn = ent.GetAttributeValue<DateTime>("msdyn_timeclosed");
                        }

                        opWarrantyStatus = ent.GetAttributeValue<OptionSetValue>("hil_warrantystatus");
                        _LaborInWarranty = ent.GetAttributeValue<bool>("hil_laborinwarranty");

                        if (ent.Attributes.Contains("hil_quantity"))
                        {
                            _Jobqty = ent.GetAttributeValue<int>("hil_quantity");
                        }
                        if (ent.Attributes.Contains("hil_owneraccount"))
                        {
                            _erRelatedFranchise = ent.GetAttributeValue<EntityReference>("hil_owneraccount");
                        }
                        if (ent.Attributes.Contains("custCatg.hil_category"))
                        {
                            _opRelatedFranchiseCategory = ((OptionSetValue)ent.GetAttributeValue<AliasedValue>("custCatg.hil_category").Value);
                        }
                        if (ent.Attributes.Contains("hil_countryclassification"))
                        {
                            opCountryClassification = ent.GetAttributeValue<OptionSetValue>("hil_countryclassification");
                        }
                        if (ent.Attributes.Contains("hil_jobclosemobile"))
                        {
                            _jobclosemobile = ent.GetAttributeValue<bool>("hil_jobclosemobile");
                        }
                        if (ent.Attributes.Contains("hil_sparepartuse"))
                        {
                            _sparepartuse = ent.GetAttributeValue<bool>("hil_sparepartuse");
                        }
                        if (ent.Attributes.Contains("hil_isgascharged"))
                        {
                            _isGascharged = ent.GetAttributeValue<bool>("hil_isgascharged");
                        }

                        if (ent.Attributes.Contains("hil_callsubtype"))
                        {
                            erCallSubType = ent.GetAttributeValue<EntityReference>("hil_callsubtype");
                        }
                        if (ent.Attributes.Contains("hil_kkgcode"))
                        {
                            kkgcode = ent.GetAttributeValue<string>("hil_kkgcode");
                        }
                        if (ent.Attributes.Contains("hil_schemecode"))
                        {
                            erSchemeCode = ent.GetAttributeValue<EntityReference>("hil_schemecode");
                        }
                        if (ent.Attributes.Contains("hil_typeofassignee"))
                        {
                            erTypeOfAssignee = ent.GetAttributeValue<EntityReference>("hil_typeofassignee");
                        }
                        if (ent.Attributes.Contains("hil_salesoffice"))
                        {
                            erSalesOffice = ent.GetAttributeValue<EntityReference>("hil_salesoffice");
                        }
                        if (ent.Attributes.Contains("hil_tatcategory"))
                        {
                            erTatCategory = ent.GetAttributeValue<EntityReference>("hil_tatcategory");
                        }
                        if (ent.Attributes.Contains("hil_productcategory"))
                        {
                            erProdCatg = ent.GetAttributeValue<EntityReference>("hil_productcategory");
                        }
                        if (ent.Attributes.Contains("hil_productsubcategory"))
                        {
                            erProdSubCatg = ent.GetAttributeValue<EntityReference>("hil_productsubcategory");
                        }
                        if (ent.Attributes.Contains("hil_sourceofjob"))
                        {
                            _opSourceOfJob = ent.GetAttributeValue<OptionSetValue>("hil_sourceofjob");
                        }
                        #endregion

                        if (opClaimStatus.Value == 7) //Abnormal Penalty Imposed
                        {

                            queryExp = new QueryExpression("hil_wrongclosurepenalty");
                            queryExp.ColumnSet = new ColumnSet("hil_amount");
                            queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                            queryExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                            queryExp.Criteria.AddCondition("hil_brand", ConditionOperator.Equal, opBrand.Value);
                            queryExp.Criteria.AddCondition("hil_wrongclosurecategory", ConditionOperator.Equal, 1);
                            queryExp.Criteria.AddCondition(new ConditionExpression("hil_fromdate", ConditionOperator.OnOrBefore, _closedOn)); //
                            queryExp.Criteria.AddCondition(new ConditionExpression("hil_todate", ConditionOperator.OnOrAfter, _closedOn)); //
                            entcoll = _service.RetrieveMultiple(queryExp);
                            if (entcoll.Entities.Count > 0)
                            {
                                UpsertClaimLine(_service, _performaInvoiceId, (-1 * entcoll.Entities[0].GetAttributeValue<Money>("hil_amount").Value), ClaimCategoryConst._KKGAuditFailed.ToString(), _erClaimPeriod, _erRelatedFranchise, ent.ToEntityReference());
                            }
                        }
                        else if (opClaimStatus.Value == 4 && opWarrantyStatus.Value == 1)//Claim Approved & IN Warranty
                        {
                            #region KKG Closure
                            if (kkgcode != string.Empty)
                            {
                                queryExp = new QueryExpression("hil_travelclosureincentive");
                                queryExp.ColumnSet = new ColumnSet("hil_amount");
                                queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                                queryExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                                queryExp.Criteria.AddCondition("hil_brand", ConditionOperator.Equal, opBrand.Value);
                                queryExp.Criteria.AddCondition("hil_incentivetype", ConditionOperator.Equal, 1);
                                queryExp.Criteria.AddCondition(new ConditionExpression("hil_startdate", ConditionOperator.OnOrBefore, _closedOn)); //
                                queryExp.Criteria.AddCondition(new ConditionExpression("hil_enddate", ConditionOperator.OnOrAfter, _closedOn)); //
                                entcoll = _service.RetrieveMultiple(queryExp);
                                if (entcoll.Entities.Count > 0)
                                {
                                    if (erCallSubType.Id != _callSubTypeId_AMC) // NON AMC JOB
                                    {
                                        UpsertClaimLine(_service, _performaInvoiceId, (entcoll.Entities[0].GetAttributeValue<Money>("hil_amount").Value), ClaimCategoryConst._KKGClosure.ToString(), _erClaimPeriod, _erRelatedFranchise, ent.ToEntityReference());
                                    }
                                }
                            }
                            #endregion

                            if (_LaborInWarranty)//Labor IN Warranty
                            {
                                #region Abnormal Closure

                                #endregion

                                #region Base Call Charge/Upcountry Charge
                                _strJobClosedOn = _closedOn.Year.ToString() + "-" + _closedOn.Month.ToString().PadLeft(2, '0') + "-" + _closedOn.Day.ToString().PadLeft(2, '0');

                                _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                    <entity name='hil_jobbasecharge'>
                                    <attribute name='hil_jobbasechargeid' />
                                    <attribute name='hil_upcountrycharge' />
                                    <attribute name='hil_rate' />
                                    <order attribute='hil_name' descending='false' />
                                    <filter type='and'>
                                    <filter type='or'>
                                        <condition attribute='hil_channelpartnercode' operator='null' />
                                        <condition attribute='hil_channelpartnercode' operator='eq' value='{" + _erRelatedFranchise.Id + @"}' />
                                    </filter>
                                    <condition attribute='hil_productsubcategory' operator='eq' value='{" + erProdSubCatg.Id + @"}' />
                                    <condition attribute='hil_callsubtype' operator='eq' value='{" + erCallSubType.Id + @"}' />
                                    <condition attribute='hil_gascharged' operator='eq' value='" + (_isGascharged ? 1 : 0).ToString() + @"' />
                                    <filter type='or'>
                                        <condition attribute='hil_sparepartuse' operator='eq' value='0' />
                                        <condition attribute='hil_sparepartuse' operator='eq' value='" + (_sparepartuse ? 1 : 2).ToString() + @"' />
                                    </filter>
                                     <condition attribute='statecode' operator='eq' value='0' />
                                     <condition attribute='hil_startdate' operator='on-or-before' value='" + _strJobClosedOn + @"' />
                                     <condition attribute='hil_enddate' operator='on-or-after' value='" + _strJobClosedOn + @"' /> 
                                </filter>
                                </entity>
                                </fetch>";

                                entcoll = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                                if (entcoll.Entities.Count > 0)
                                {
                                    decimal _baseRate = 0;
                                    if (entcoll.Entities[0].Attributes.Contains("hil_rate"))
                                    {
                                        if (opBrand.Value == 1 && erCallSubType.Id == _callSubTypeId_Installation)
                                        {
                                            _baseRate = entcoll.Entities[0].GetAttributeValue<Money>("hil_rate").Value * _Jobqty;
                                        }
                                        else
                                        {
                                            _baseRate = entcoll.Entities[0].GetAttributeValue<Money>("hil_rate").Value;
                                        }
                                    }
                                    if (_baseRate > 0)
                                    {
                                        UpsertClaimLine(_service, _performaInvoiceId, _baseRate, ClaimCategoryConst._BaseCallCharge.ToString(), _erClaimPeriod, _erRelatedFranchise, ent.ToEntityReference());
                                    }
                                    #region Upcountry Travel Charge
                                    if (entcoll.Entities[0].Attributes.Contains("hil_upcountrycharge"))
                                    {
                                        if (opCountryClassification.Value == 2 && entcoll.Entities[0].GetAttributeValue<Money>("hil_upcountrycharge").Value > 0)
                                        {
                                            UpsertClaimLine(_service, _performaInvoiceId, (entcoll.Entities[0].GetAttributeValue<Money>("hil_upcountrycharge").Value), ClaimCategoryConst._UpcountryTravelCharge.ToString(), _erClaimPeriod, _erRelatedFranchise, ent.ToEntityReference());
                                        }
                                    }
                                    #endregion
                                }
                                #endregion

                                #region Minimum Guarantee

                                #endregion

                                #region Mobile Closure
                                if (_jobclosemobile)
                                {
                                    queryExp = new QueryExpression("hil_travelclosureincentive");
                                    queryExp.ColumnSet = new ColumnSet("hil_amount");
                                    queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                                    queryExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                                    queryExp.Criteria.AddCondition("hil_brand", ConditionOperator.Equal, opBrand.Value);
                                    queryExp.Criteria.AddCondition("hil_incentivetype", ConditionOperator.Equal, 2);
                                    queryExp.Criteria.AddCondition(new ConditionExpression("hil_startdate", ConditionOperator.OnOrBefore, _closedOn)); //
                                    queryExp.Criteria.AddCondition(new ConditionExpression("hil_enddate", ConditionOperator.OnOrAfter, _closedOn)); //
                                    entcoll = _service.RetrieveMultiple(queryExp);
                                    if (entcoll.Entities.Count > 0)
                                    {
                                        if (entcoll.Entities[0].GetAttributeValue<Money>("hil_amount").Value > 0)
                                        {
                                            UpsertClaimLine(_service, _performaInvoiceId, (entcoll.Entities[0].GetAttributeValue<Money>("hil_amount").Value), ClaimCategoryConst._MobileClosureIncentive.ToString(), _erClaimPeriod, _erRelatedFranchise, ent.ToEntityReference());
                                        }
                                    }
                                }
                                #endregion

                                #region Scheme Incentive
                                if (erSchemeCode != null)
                                {
                                    queryExp = new QueryExpression("hil_schemeincentive");
                                    queryExp.ColumnSet = new ColumnSet("hil_amount");
                                    queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                                    queryExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                                    queryExp.Criteria.AddCondition("hil_schemeincentiveid", ConditionOperator.Equal, erSchemeCode.Id);
                                    entcoll = _service.RetrieveMultiple(queryExp);
                                    if (entcoll.Entities.Count > 0)
                                    {
                                        if (entcoll.Entities[0].GetAttributeValue<Money>("hil_amount").Value > 0)
                                        {
                                            UpsertClaimLine(_service, _performaInvoiceId, (entcoll.Entities[0].GetAttributeValue<Money>("hil_amount").Value), ClaimCategoryConst._SchemeIncentive.ToString(), _erClaimPeriod, _erRelatedFranchise, ent.ToEntityReference());
                                        }
                                    }
                                }
                                #endregion

                                #region TAT Breach Penalty
                                if (erTatCategory != null)
                                {
                                    queryExp = new QueryExpression("hil_tatbreachpenalty");
                                    queryExp.ColumnSet = new ColumnSet("hil_amount");
                                    queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                                    queryExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                                    queryExp.Criteria.AddCondition("hil_brand", ConditionOperator.Equal, opBrand.Value);
                                    queryExp.Criteria.AddCondition("hil_tathour", ConditionOperator.Equal, erTatCategory.Id);
                                    queryExp.Criteria.AddCondition(new ConditionExpression("hil_fromdate", ConditionOperator.OnOrBefore, _closedOn)); //
                                    queryExp.Criteria.AddCondition(new ConditionExpression("hil_to", ConditionOperator.OnOrAfter, _closedOn)); //
                                    entcoll = _service.RetrieveMultiple(queryExp);
                                    if (entcoll.Entities.Count > 0)
                                    {
                                        if (entcoll.Entities[0].GetAttributeValue<Money>("hil_amount").Value > 0)
                                        {
                                            if (opBrand.Value == 3)
                                            { // If Brand is Havells Aqua Allow only for (Installation, Breakdown & Non System generated PMS)
                                                if (erCallSubType.Id == _callSubTypeId_Breakdown || erCallSubType.Id == _callSubTypeId_Installation || (erCallSubType.Id == _callSubTypeId_PMS && _opSourceOfJob.Value != 10))
                                                {
                                                    UpsertClaimLine(_service, _performaInvoiceId, (entcoll.Entities[0].GetAttributeValue<Money>("hil_amount").Value * -1), ClaimCategoryConst._TATBreachPenalty.ToString(), _erClaimPeriod, _erRelatedFranchise, ent.ToEntityReference());
                                                }
                                            }
                                            else // Other Brands
                                            {
                                                UpsertClaimLine(_service, _performaInvoiceId, (entcoll.Entities[0].GetAttributeValue<Money>("hil_amount").Value * -1), ClaimCategoryConst._TATBreachPenalty.ToString(), _erClaimPeriod, _erRelatedFranchise, ent.ToEntityReference());
                                            }
                                        }
                                    }
                                }
                                #endregion

                                #region TAT Incentive for LLOYD Jobs applicable on AC+Installation
                                if (opBrand.Value == 2 && _prodCategoryId_LLOYDAirCon == erProdCatg.Id && _callSubTypeId_Installation == erCallSubType.Id)
                                {
                                    bool _validForTATIncentive = true;
                                    if (_isGascharged || erSchemeCode == null)
                                    {
                                        _validForTATIncentive = false;
                                    }
                                    if (_validForTATIncentive)
                                    {
                                        entcoll = RetrieveTATIncentiveLLOYD(_service, _sparepartuse, opBrand, _opRelatedFranchiseCategory, erCallSubType, erProdSubCatg, erTatCategory, _closedOn);
                                        if (entcoll.Entities.Count > 0)
                                        {
                                            if (entcoll.Entities[0].GetAttributeValue<Money>("hil_amount").Value > 0)
                                            {
                                                UpsertClaimLine(_service, _performaInvoiceId, (entcoll.Entities[0].GetAttributeValue<Money>("hil_amount").Value), ClaimCategoryConst._TATIncentive.ToString(), _erClaimPeriod, _erRelatedFranchise, ent.ToEntityReference());
                                            }
                                        }
                                    }
                                }
                                #endregion

                                #region Special Incentive
                                _strJobClosedOn = _closedOn.Year.ToString() + "-" + _closedOn.Month.ToString().PadLeft(2, '0') + "-" + _closedOn.Day.ToString().PadLeft(2, '0');

                                _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                <entity name='hil_specialincentive'>
                                <attribute name='hil_amount' />
                                <filter type='and'>
                                    <condition attribute='statecode' operator='eq' value='0' />
                                    <condition attribute='hil_callsubtype' operator='eq' value='{" + erCallSubType.Id + @"}' />
                                    <condition attribute='hil_fromdate' operator='on-or-before' value='" + _strJobClosedOn + @"' />
                                    <condition attribute='hil_todate' operator='on-or-after' value='" + _strJobClosedOn + @"' />
                                    <condition attribute='hil_productcategory' operator='eq' value='{" + erProdCatg.Id + @"}' />
                                    <filter type='or'>
                                    <condition attribute='hil_salesoffice' operator='eq' value='{90503976-8FD1-EA11-A813-000D3AF0563C}' />
                                    <condition attribute='hil_salesoffice' operator='eq' value='{" + erSalesOffice.Id + @"}' />
                                    </filter>
                                </filter>
                                </entity>
                                </fetch>";

                                entcoll = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                                if (entcoll.Entities.Count > 0)
                                {
                                    decimal _baseRate = 0;
                                    if (entcoll.Entities[0].Attributes.Contains("hil_amount"))
                                    {
                                        _baseRate = entcoll.Entities[0].GetAttributeValue<Money>("hil_amount").Value;
                                    }
                                    if (_baseRate > 0)
                                    {
                                        UpsertClaimLine(_service, _performaInvoiceId, _baseRate, ClaimCategoryConst._SpecialIncentive.ToString(), _erClaimPeriod, _erRelatedFranchise, ent.ToEntityReference());
                                    }
                                }
                                #endregion
                            }
                        }
                    }
                }
            }
        }

        private static EntityCollection RetrieveTATIncentiveLLOYD(IOrganizationService _service, bool _sparepartuse, OptionSetValue opBrand, OptionSetValue _opRelatedFranchiseCategory, EntityReference erCallSubType, EntityReference erProdSubCatg, EntityReference erTatCategory, DateTime _closedOn)
        {
            string _franchiseCategory = _opRelatedFranchiseCategory == null ? "910590000" : _opRelatedFranchiseCategory.Value.ToString();
            string _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                <entity name='hil_tatincentive'>
                <attribute name='hil_tatincentiveid' />
                <attribute name='hil_amount' />
                <order attribute='hil_amount' descending='false' />
                <filter type='and'>
                <condition attribute='statecode' operator='eq' value='0' />
                <condition attribute='hil_brand' operator='eq' value='" + opBrand.Value.ToString() + @"' />
                <condition attribute='hil_callsubtype' operator='eq' value='{" + erCallSubType.Id + @"}' />
                <condition attribute='hil_channelpartnercategory' operator='eq' value='" + _franchiseCategory + @"' />
                <condition attribute='hil_productsubcategory' operator='eq' value='{" + erProdSubCatg.Id + @"}' />
                <condition attribute='hil_tathourslab' operator='eq' value='{" + erTatCategory.Id + @"}' />
                <condition attribute='hil_startdate' operator='on-or-before' value='" + (_closedOn.Year.ToString() + "-" + _closedOn.Month.ToString().PadLeft(2, '0') + "-" + _closedOn.Day.ToString().PadLeft(2, '0')) + @"' />
                <condition attribute='hil_enddate' operator='on-or-after' value='" + (_closedOn.Year.ToString() + "-" + _closedOn.Month.ToString().PadLeft(2, '0') + "-" + _closedOn.Day.ToString().PadLeft(2, '0')) + @"' />
                <filter type='or'>
                    <condition attribute='hil_sparepartuse' operator='eq' value='1' />
                    <condition attribute='hil_sparepartuse' operator='eq' value='" + (_sparepartuse ? 2 : 3).ToString() + @"' />
                </filter>
            </filter>
            </entity>
            </fetch>";
            return _service.RetrieveMultiple(new FetchExpression(_fetchXML));
        }

        private static EntityCollection RetrieveTATIncentiveNonLLOYD(IOrganizationService _service, bool _sparepartuse, OptionSetValue opBrand, OptionSetValue _opRelatedFranchiseCategory, EntityReference erCallSubType, EntityReference erProdSubCatg, EntityReference erTatSlab, DateTime _closedOn)
        {
            QueryExpression queryExp;
            EntityCollection entcoll;

            string _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
            <entity name='hil_tatincentive'>
            <attribute name='hil_tatincentiveid' />
            <attribute name='hil_amount' />
            <attribute name='createdon' />
            <order attribute='hil_name' descending='false' />
            <filter type='and'>
                <condition attribute='statecode' operator='eq' value='0' />
                <condition attribute='hil_brand' operator='eq' value='" + opBrand.Value.ToString() + @"' />
                <condition attribute='hil_callsubtype' operator='eq' value='{" + erCallSubType.Id + @"}' />
                <condition attribute='hil_channelpartnercategory' operator='eq' value='" + _opRelatedFranchiseCategory.Value.ToString() + @"' />
                <condition attribute='hil_productsubcategory' operator='eq' value='{" + erProdSubCatg.Id + @"}' />
                <condition attribute='hil_tatacheivementslab' operator='eq' value='{" + erTatSlab.Id + @"}' />
                <condition attribute='hil_startdate' operator='on-or-before' value='" + (_closedOn.Year.ToString() + "-" + _closedOn.Month.ToString().PadLeft(2, '0') + "-" + _closedOn.Day.ToString().PadLeft(2, '0')) + @"' />
                <condition attribute='hil_enddate' operator='on-or-after' value='" + (_closedOn.Year.ToString() + "-" + _closedOn.Month.ToString().PadLeft(2, '0') + "-" + _closedOn.Day.ToString().PadLeft(2, '0')) + @"' />
                <filter type='or'>
                    <condition attribute='hil_sparepartuse' operator='eq' value='1' />
                    <condition attribute='hil_sparepartuse' operator='eq' value='" + (_sparepartuse ? 2 : 3).ToString() + @"' />
                </filter>
            </filter>
            </entity>
            </fetch>";
            entcoll = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
            return entcoll;
        }
        public static void GenerateClaimOverHeads(IOrganizationService _service, Guid _performaInvoiceId)
        {
            QueryExpression queryExp;
            EntityCollection entcoll;

            LinkEntity lnkEntCC = new LinkEntity
            {
                LinkFromEntityName = "hil_claimoverheadline",
                LinkToEntityName = "product",
                LinkFromAttributeName = "hil_productcategory",
                LinkToAttributeName = "productid",
                Columns = new ColumnSet("hil_claimcategory"),
                EntityAlias = "cc",
                JoinOperator = JoinOperator.LeftOuter
            };
            queryExp = new QueryExpression("hil_claimoverheadline");
            queryExp.ColumnSet = new ColumnSet("hil_productcategory", "hil_performainvoice", "hil_callsubtype", "hil_amount");
            queryExp.Criteria = new FilterExpression(LogicalOperator.And);
            queryExp.Criteria.AddCondition("hil_performainvoice", ConditionOperator.Equal, _performaInvoiceId);
            queryExp.Criteria.AddCondition("hil_productcategory", ConditionOperator.NotNull);
            queryExp.Criteria.AddCondition("hil_callsubtype", ConditionOperator.NotNull);
            queryExp.Criteria.AddCondition("hil_amount", ConditionOperator.NotNull);
            queryExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
            queryExp.LinkEntities.Add(lnkEntCC);

            entcoll = _service.RetrieveMultiple(queryExp);
            if (entcoll.Entities.Count > 0)
            {
                LinkEntity lnkEntSO = new LinkEntity
                {
                    LinkFromEntityName = "account",
                    LinkToEntityName = "hil_salesoffice",
                    LinkFromAttributeName = "hil_salesoffice",
                    LinkToAttributeName = "hil_salesofficeid",
                    Columns = new ColumnSet("hil_state"),
                    EntityAlias = "so",
                    JoinOperator = JoinOperator.Inner
                };
                LinkEntity lnkEntCP = new LinkEntity
                {
                    LinkFromEntityName = "hil_claimheader",
                    LinkToEntityName = "account",
                    LinkFromAttributeName = "hil_franchisee",
                    LinkToAttributeName = "accountid",
                    Columns = new ColumnSet("hil_state", "hil_salesoffice", "ownerid"),
                    EntityAlias = "cp",
                    JoinOperator = JoinOperator.Inner
                };
                lnkEntCP.LinkEntities.Add(lnkEntSO);

                queryExp = new QueryExpression("hil_claimheader");
                queryExp.ColumnSet = new ColumnSet("hil_franchisee", "hil_fiscalmonth");
                queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                queryExp.Criteria.AddCondition("hil_claimheaderid", ConditionOperator.Equal, _performaInvoiceId);
                queryExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                queryExp.LinkEntities.Add(lnkEntCP);
                EntityCollection entcollCH = _service.RetrieveMultiple(queryExp);
                if (entcollCH.Entities.Count > 0)
                {
                    int _gstCatg;
                    EntityReference _erCP = entcollCH.Entities[0].GetAttributeValue<EntityReference>("hil_franchisee");
                    EntityReference _erCM = entcollCH.Entities[0].GetAttributeValue<EntityReference>("hil_fiscalmonth");
                    Guid _ownerId = ((EntityReference)entcollCH.Entities[0].GetAttributeValue<AliasedValue>("cp.ownerid").Value).Id;

                    if (entcollCH.Entities[0].Contains("cp.hil_state") && entcollCH.Entities[0].Contains("so.hil_state")) {
                        EntityReference _erCPState = ((EntityReference)entcollCH.Entities[0].GetAttributeValue<AliasedValue>("cp.hil_state").Value);
                        EntityReference _erSOState = ((EntityReference)entcollCH.Entities[0].GetAttributeValue<AliasedValue>("so.hil_state").Value);
                        _gstCatg = _erCPState.Id == _erSOState.Id ? 1 : 2;
                        QueryExpression queryExpTemp;
                        EntityCollection entcollTemp;
                        string _activityCode = string.Empty;
                        string _claimCatwgory = string.Empty;

                        foreach (Entity ent in entcoll.Entities)
                        {
                            _activityCode = "";
                            _claimCatwgory = "";
                            queryExpTemp = new QueryExpression("hil_claimpostingsetup");
                            queryExpTemp.ColumnSet = new ColumnSet("hil_activitycode");
                            queryExpTemp.Criteria = new FilterExpression(LogicalOperator.And);
                            queryExpTemp.Criteria.AddCondition("hil_callsubtype", ConditionOperator.Equal, ent.GetAttributeValue<EntityReference>("hil_callsubtype").Id);
                            queryExpTemp.Criteria.AddCondition("hil_productcategory", ConditionOperator.Equal, ent.GetAttributeValue<EntityReference>("hil_productcategory").Id);
                            queryExpTemp.Criteria.AddCondition("hil_activitygstslab", ConditionOperator.Equal, _gstCatg);
                            queryExpTemp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                            entcollTemp = _service.RetrieveMultiple(queryExpTemp);
                            if (entcollTemp.Entities.Count > 0) {
                                _activityCode = entcollTemp.Entities[0].GetAttributeValue<string>("hil_activitycode");
                            }
                            if (ent.Contains("cc.hil_claimcategory"))
                            {
                                _claimCatwgory = ((EntityReference)ent.GetAttributeValue<AliasedValue>("cc.hil_claimcategory").Value).Id.ToString();
                            }
                            else {
                                _claimCatwgory = "DEAC05B5-8E68-EC11-8943-6045BDA5EF08"; //Lumpsum Payment
                            }
                            ClaimLineDTO _claimLine = new ClaimLineDTO()
                            {
                                CallSubtype = ent.GetAttributeValue<EntityReference>("hil_callsubtype"),
                                ChannelPartner = _erCP,
                                ClaimAmount = ent.GetAttributeValue<Money>("hil_amount").Value,
                                ClaimCatwgory = _claimCatwgory,
                                ClaimPeriod = _erCM,
                                OwnerId = _ownerId,
                                PerformInvoiceId = ent.GetAttributeValue<EntityReference>("hil_performainvoice"),
                                ProdCatg = ent.GetAttributeValue<EntityReference>("hil_productcategory"),
                                ActivityCode = _activityCode
                            };
                            CreateOverheadClaimLine(_service, _claimLine);
                        }
                    }
                }
            }
        }
        private static int GetJobCountClaimMonth(IOrganizationService _service,Guid _franchiseId, Guid _claimMonthId) {
            QueryExpression queryExp;
            EntityCollection entcoll;

            try
            {
                queryExp = new QueryExpression(msdyn_workorder.EntityLogicalName);
                queryExp.ColumnSet = new ColumnSet(false);
                queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                queryExp.Criteria.AddCondition("hil_owneraccount", ConditionOperator.Equal, _franchiseId);
                queryExp.Criteria.AddCondition("hil_tatcategory", ConditionOperator.NotNull);
                queryExp.Criteria.AddCondition("hil_isocr", ConditionOperator.NotEqual, true);
                queryExp.Criteria.AddCondition("hil_claimstatus", ConditionOperator.Equal, 4);// Claim Approved
                queryExp.Criteria.AddCondition("msdyn_substatus", ConditionOperator.Equal, new Guid("1727FA6C-FA0F-E911-A94E-000D3AF060A1")); //Closed
                queryExp.Criteria.AddCondition("hil_fiscalmonth", ConditionOperator.Equal, _claimMonthId); //
                entcoll = _service.RetrieveMultiple(queryExp);
                return entcoll.Entities.Count;
            }
            catch 
            {
                return 0;
            }
        }
        private static EntityCollection GetProdCatgWiseJobCountClaimMonth(IOrganizationService _service, Guid _franchiseId, Guid _claimMonthId)
        {
            EntityCollection entcoll;

            try
            {
                string _fetchXML = @"<fetch distinct='false' mapping='logical' aggregate='true'>
                <entity name='msdyn_workorder'>
                <attribute name='msdyn_workorderid' alias='rowcount' aggregate='count'/> 
                <attribute name='hil_productcategory' alias='prodcatg' groupby='true' /> 
                <filter type='and'>
                    <condition attribute='hil_isocr' operator='ne' value='1' />
                    <condition attribute='msdyn_substatus' operator='eq' value='{1727FA6C-FA0F-E911-A94E-000D3AF060A1}' />
                    <condition attribute='hil_claimstatus' operator='eq' value='4' />
                    <condition attribute='hil_tatcategory' operator='not-null' />
                    <condition attribute='hil_owneraccount' operator='eq' value='{" + _franchiseId + @"}' />
                    <condition attribute='hil_fiscalmonth' operator='eq' value='{" + _claimMonthId + @"}' />
                </filter>
                </entity>
                </fetch>";
                entcoll = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                return entcoll;
            }
            catch
            {
                return null;
            }
        }
        public static void UpdatePerformaInvoice(IOrganizationService _service, Guid _performaInvoiceId)
        {
            QueryExpression queryExp;
            EntityCollection entcoll;
            EntityReference _erSalesOffice = null;
            EntityReference _erFiscalMonth = null;
            Guid _fixedCompensationId = new Guid("824A1BCC-6EE5-EA11-A817-000D3AF0501C");
            try
            {
                if (_performaInvoiceId != Guid.Empty)
                {
                    queryExp = new QueryExpression("hil_claimheader");
                    queryExp.ColumnSet = new ColumnSet("hil_fiscalmonth");
                    queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                    queryExp.Criteria.AddCondition("hil_claimheaderid", ConditionOperator.Equal, _performaInvoiceId);
                    entcoll = _service.RetrieveMultiple(queryExp);
                    if (entcoll.Entities.Count > 0)
                    {
                        _erFiscalMonth = entcoll.Entities[0].GetAttributeValue<EntityReference>("hil_fiscalmonth");
                    }
                }
                else
                {
                    queryExp = new QueryExpression("hil_claimperiod");
                    queryExp.ColumnSet = new ColumnSet("hil_fromdate", "hil_todate");
                    queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                    queryExp.Criteria.AddCondition("hil_isapplicable", ConditionOperator.Equal, true);
                    queryExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                    entcoll = _service.RetrieveMultiple(queryExp);
                    if (entcoll.Entities.Count > 0)
                    {
                        _erFiscalMonth = entcoll.Entities[0].ToEntityReference();
                    }
                }
                if (_erFiscalMonth == null)
                {
                    Console.WriteLine("Fiscal Month is not defined." + DateTime.Now.ToString());
                    return;
                }

                if (_erFiscalMonth != null)
                {
                    LinkEntity lnkEntCP = new LinkEntity
                    {
                        LinkFromEntityName = "hil_claimheader",
                        LinkToEntityName = "account",
                        LinkFromAttributeName = "hil_franchisee",
                        LinkToAttributeName = "accountid",
                        Columns = new ColumnSet("hil_salesoffice"),
                        EntityAlias = "cp",
                        JoinOperator = JoinOperator.Inner
                    };
                    queryExp = new QueryExpression("hil_claimheader");
                    queryExp.ColumnSet = new ColumnSet("hil_claimheaderid", "hil_name");
                    queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                    if(_performaInvoiceId!=Guid.Empty)
                        queryExp.Criteria.AddCondition("hil_claimheaderid", ConditionOperator.Equal, _performaInvoiceId);

                    queryExp.Criteria.AddCondition("hil_fiscalmonth", ConditionOperator.Equal, _erFiscalMonth.Id);
                    //queryExp.Criteria.AddCondition("hil_performastatus", ConditionOperator.Equal, 3);

                    queryExp.LinkEntities.Add(lnkEntCP);
                    entcoll = _service.RetrieveMultiple(queryExp);
                    int _cnt = 0, _jobCount;
                    Decimal _claimAmount, _fixedCharge, _overHeadsAmount, _netClaimAmount;

                    foreach (Entity ent in entcoll.Entities)
                    {
                        try
                        {
                            _cnt += 1;
                            Console.WriteLine("Processing Performa Invoice# " + _cnt.ToString() + " : " + ent.GetAttributeValue<string>("hil_name"));

                            _performaInvoiceId = ent.Id;
                            _jobCount = 0;
                            _claimAmount = 0;
                            _fixedCharge = 0;
                            _overHeadsAmount = 0;
                            _netClaimAmount = 0;

                            if (ent.Contains("cp.hil_salesoffice"))
                            {
                                _erSalesOffice = (EntityReference)ent.GetAttributeValue<AliasedValue>("cp.hil_salesoffice").Value;
                            }
                            string _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                            <entity name='hil_claimline'>
                            <attribute name='hil_jobid' />
                            <filter type='and'>
                                <condition attribute='hil_claimheader' operator='eq' value='{" + _performaInvoiceId + @"}' />
                                <condition attribute='statecode' operator='eq' value='0' />
                                <condition attribute='hil_jobid' operator='not-null' />
                            </filter>
                            </entity>
                            </fetch>";
                            entcoll = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                            if (entcoll.Entities.Count > 0)
                            {
                                _jobCount = entcoll.Entities.Count;
                            }
                            _fetchXML = @"<fetch distinct='false' mapping='logical' aggregate='true'>
                            <entity name='hil_claimline'>
                            <attribute name='hil_claimamount' alias='claimamount' aggregate='sum'/> 
                            <filter type='and'>
                                <condition attribute='hil_claimheader' operator='eq' value='{" + _performaInvoiceId + @"}' />
                                <condition attribute='statecode' operator='eq' value='0' />
                                <condition attribute='hil_claimcategory' operator='not-in'>
                                    <value uiname='Fixed Compensation' uitype='hil_claimcategory'>{824A1BCC-6EE5-EA11-A817-000D3AF0501C}</value>
                                    <value uiname='Lumpsum Payment' uitype='hil_claimcategory'>{DEAC05B5-8E68-EC11-8943-6045BDA5EF08}</value>
                                    <value uiname='Non-Standardized Charges' uitype='hil_claimcategory'>{501D69F1-8E68-EC11-8943-6045BDA5EF08}</value>
                                    <value uiname='Reverse Logistics' uitype='hil_claimcategory'>{51B1D88E-8E68-EC11-8943-6045BDA5EF08}</value>
                                    <value uiname='Abnormal Closure' uitype='hil_claimcategory'>{8dbfb0c4-33f3-ea11-a815-000d3af05a4b}</value>
                                </condition>
                            </filter>
                            </entity>
                            </fetch>";
                            entcoll = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                            if (entcoll.Entities.Count > 0)
                            {
                                if (entcoll.Entities[0].Contains("claimamount") && entcoll.Entities[0].GetAttributeValue<AliasedValue>("claimamount").Value != null)
                                    _claimAmount = ((Money)entcoll.Entities[0].GetAttributeValue<AliasedValue>("claimamount").Value).Value;
                            }
                            _fetchXML = @"<fetch distinct='false' mapping='logical' aggregate='true'>
                                <entity name='hil_claimline'>
                                <attribute name='hil_claimamount' alias='claimamount' aggregate='sum'/> 
                                <filter type='and'>
                                    <condition attribute='hil_claimheader' operator='eq' value='{" + _performaInvoiceId + @"}' />
                                    <condition attribute='statecode' operator='eq' value='0' />
                                    <condition attribute='hil_claimcategory' operator='eq' value='{" + _fixedCompensationId + @"}' />
                                </filter>
                                </entity>
                                </fetch>";
                            entcoll = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                            if (entcoll.Entities.Count > 0)
                            {
                                if (entcoll.Entities[0].Contains("claimamount") && entcoll.Entities[0].GetAttributeValue<AliasedValue>("claimamount").Value != null)
                                    _fixedCharge = ((Money)entcoll.Entities[0].GetAttributeValue<AliasedValue>("claimamount").Value).Value;
                            }
                            //_fetchXML = @"<fetch distinct='false' mapping='logical' aggregate='true'>
                            //    <entity name='hil_claimoverheadline'>
                            //    <attribute name='hil_amount' alias='claimamount' aggregate='sum'/> 
                            //    <filter type='and'>
                            //        <condition attribute='hil_performainvoice' operator='eq' value='{" + _performaInvoiceId + @"}' />
                            //        <condition attribute='statecode' operator='eq' value='0' />
                            //    </filter>
                            //    </entity>
                            //    </fetch>";
                            _fetchXML = @"<fetch distinct='false' mapping='logical' aggregate='true'>
                                <entity name='hil_claimline'>
                                <attribute name='hil_claimamount' alias='claimamount' aggregate='sum'/> 
                                <filter type='and'>
                                    <condition attribute='hil_claimheader' operator='eq' value='{" + _performaInvoiceId + @"}' />
                                    <condition attribute='statecode' operator='eq' value='0' />
                                    <condition attribute='hil_claimcategory' operator='in'>
                                        <value uiname='Lumpsum Payment' uitype='hil_claimcategory'>{DEAC05B5-8E68-EC11-8943-6045BDA5EF08}</value>
                                        <value uiname='Non-Standardized Charges' uitype='hil_claimcategory'>{501D69F1-8E68-EC11-8943-6045BDA5EF08}</value>
                                        <value uiname='Reverse Logistics' uitype='hil_claimcategory'>{51B1D88E-8E68-EC11-8943-6045BDA5EF08}</value>
                                        <value uiname='Abnormal Closure' uitype='hil_claimcategory'>{8dbfb0c4-33f3-ea11-a815-000d3af05a4b}</value>
                                    </condition>
                                </filter>
                                </entity>
                                </fetch>";
                            entcoll = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                            if (entcoll.Entities.Count > 0)
                            {
                                if (entcoll.Entities[0].Contains("claimamount") && entcoll.Entities[0].GetAttributeValue<AliasedValue>("claimamount").Value != null)
                                    _overHeadsAmount = ((Money)entcoll.Entities[0].GetAttributeValue<AliasedValue>("claimamount").Value).Value;
                            }
                            _netClaimAmount = _claimAmount + _fixedCharge + _overHeadsAmount;

                            Entity _entPerformaInv = new Entity("hil_claimheader", _performaInvoiceId);
                            _entPerformaInv["hil_branch"] = _erSalesOffice;
                            _entPerformaInv["hil_totaljobsclosed"] = _jobCount;
                            _entPerformaInv["hil_totalclaimamount"] = _claimAmount;
                            _entPerformaInv["hil_fixedcharges"] = _fixedCharge;
                            _entPerformaInv["hil_totalclaimvalue"] = _netClaimAmount;
                            _entPerformaInv["hil_expenseoverheads"] = _overHeadsAmount;
                            _service.Update(_entPerformaInv);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("ERRPR!!! " + ex.Message + " Processing Performa Invoice# " + _cnt.ToString() + " : " + ent.GetAttributeValue<string>("hil_name"));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR!!!" + ex.Message);
            }
        }
        public static void UpdateActivityOnClaimLinesJun2022(IOrganizationService _service)
        {
            #region Variable declaration
            QueryExpression queryExp;
            EntityCollection entcoll;
            Guid _performaInvoiceId = Guid.Empty;
            EntityReference _erFiscalMonth = null;
            #endregion

            //DateTime _processDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);
            DateTime _processDate = new DateTime(2022, 06, 20);
            queryExp = new QueryExpression("hil_claimperiod");
            queryExp.ColumnSet = new ColumnSet("hil_fromdate", "hil_todate");
            queryExp.Criteria = new FilterExpression(LogicalOperator.And);
            queryExp.Criteria.AddCondition("hil_fromdate", ConditionOperator.OnOrBefore, _processDate);
            queryExp.Criteria.AddCondition("hil_todate", ConditionOperator.OnOrAfter, _processDate);
            entcoll = _service.RetrieveMultiple(queryExp);
            if (entcoll.Entities.Count > 0)
            {
                _erFiscalMonth = entcoll.Entities[0].ToEntityReference();
            }
            if (_erFiscalMonth == null)
            {
                Console.WriteLine("Fiscal Month is not defined." + DateTime.Now.ToString());
                return;
            }

            LinkEntity lnkEntJB = new LinkEntity
            {
                LinkFromEntityName = "hil_claimline",
                LinkToEntityName = "msdyn_workorder",
                LinkFromAttributeName = "hil_jobid",
                LinkToAttributeName = "msdyn_workorderid",
                Columns = new ColumnSet("hil_productcategory", "hil_callsubtype"),
                EntityAlias = "jb",
                JoinOperator = JoinOperator.Inner
            };

            LinkEntity lnkEntSO = new LinkEntity
            {
                LinkFromEntityName = "account",
                LinkToEntityName = "hil_salesoffice",
                LinkFromAttributeName = "hil_salesoffice",
                LinkToAttributeName = "hil_salesofficeid",
                Columns = new ColumnSet("hil_state"),
                EntityAlias = "so",
                JoinOperator = JoinOperator.Inner
            };

            LinkEntity lnkEntCP = new LinkEntity
            {
                LinkFromEntityName = "hil_claimheader",
                LinkToEntityName = "account",
                LinkFromAttributeName = "hil_franchisee",
                LinkToAttributeName = "accountid",
                Columns = new ColumnSet("hil_state", "hil_salesoffice", "ownerid"),
                EntityAlias = "cp",
                JoinOperator = JoinOperator.Inner
            };

            LinkEntity lnkEntCH = new LinkEntity
            {
                LinkFromEntityName = "hil_claimline",
                LinkToEntityName = "hil_claimheader",
                LinkFromAttributeName = "hil_claimheader",
                LinkToAttributeName = "hil_claimheaderid",
                Columns = new ColumnSet("hil_performastatus"),
                EntityAlias = "ch",
                JoinOperator = JoinOperator.Inner
            };

            lnkEntCP.LinkEntities.Add(lnkEntSO);
            lnkEntCH.LinkEntities.Add(lnkEntCP);

            queryExp = new QueryExpression("hil_claimline");
            queryExp.ColumnSet = new ColumnSet("hil_jobid", "hil_franchisee", "hil_claimperiod", "hil_claimheader", "hil_claimamount", "hil_activitycode");
            queryExp.Criteria = new FilterExpression(LogicalOperator.And);
            queryExp.Criteria.AddCondition("hil_claimperiod", ConditionOperator.Equal, _erFiscalMonth.Id);
            queryExp.Criteria.AddCondition("hil_claimcategory", ConditionOperator.Equal, new Guid("8C4A1BCC-6EE5-EA11-A817-000D3AF0501C"));
            queryExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
            queryExp.Criteria.AddCondition("hil_jobid", ConditionOperator.NotNull);
            queryExp.Criteria.AddCondition("hil_activitycode", ConditionOperator.Null);
            queryExp.LinkEntities.Add(lnkEntCH);
            queryExp.LinkEntities.Add(lnkEntJB);
            queryExp.AddOrder("createdon", OrderType.Descending);
            queryExp.NoLock = true;

            entcoll = _service.RetrieveMultiple(queryExp);
            if (entcoll.Entities.Count > 0)
            {
                int _gstCatg;
                EntityReference _erCP = null;
                QueryExpression queryExpTemp;
                EntityCollection entcollTemp;
                Entity _updateClaimLine = null;
                string _activityCodeCL = string.Empty;
                string _activityCode = string.Empty;
                int _rowCount = 0;
                foreach (Entity ent in entcoll.Entities)
                {
                    _erCP = ent.GetAttributeValue<EntityReference>("hil_franchisee");
                    _activityCode = string.Empty;
                    if (ent.Contains("cp.hil_state") && ent.Contains("so.hil_state") && ent.Contains("jb.hil_productcategory") && ent.Contains("jb.hil_callsubtype"))
                    {
                        EntityReference _erCPState = ((EntityReference)ent.GetAttributeValue<AliasedValue>("cp.hil_state").Value);
                        EntityReference _erSOState = ((EntityReference)ent.GetAttributeValue<AliasedValue>("so.hil_state").Value);
                        EntityReference _erProdCategory = ((EntityReference)ent.GetAttributeValue<AliasedValue>("jb.hil_productcategory").Value);
                        EntityReference _erCallSubtype = ((EntityReference)ent.GetAttributeValue<AliasedValue>("jb.hil_callsubtype").Value);

                        _gstCatg = _erCPState.Id == _erSOState.Id ? 1 : 2;

                        queryExpTemp = new QueryExpression("hil_claimpostingsetup");
                        queryExpTemp.ColumnSet = new ColumnSet("hil_activitycode");
                        queryExpTemp.Criteria = new FilterExpression(LogicalOperator.And);
                        queryExpTemp.Criteria.AddCondition("hil_callsubtype", ConditionOperator.Equal, _erCallSubtype.Id); //Breakdown
                        queryExpTemp.Criteria.AddCondition("hil_productcategory", ConditionOperator.Equal, _erProdCategory.Id);
                        queryExpTemp.Criteria.AddCondition("hil_activitygstslab", ConditionOperator.Equal, _gstCatg);
                        queryExpTemp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                        entcollTemp = _service.RetrieveMultiple(queryExpTemp);
                        if (entcollTemp.Entities.Count > 0)
                        {
                            _activityCode = entcollTemp.Entities[0].GetAttributeValue<string>("hil_activitycode");
                            _updateClaimLine = new Entity(ent.LogicalName, ent.Id);
                            _updateClaimLine["hil_callsubtype"] = _erCallSubtype;
                            _updateClaimLine["hil_productcategory"] = _erProdCategory;
                            _updateClaimLine["hil_activitycode"] = _activityCode;
                            _service.Update(_updateClaimLine);
                        }
                        else
                        {
                            _updateClaimLine = new Entity(ent.LogicalName, ent.Id);
                            _updateClaimLine["hil_activitycode"] = null;
                            _service.Update(_updateClaimLine);
                        }
                    }
                    _rowCount++;
                    Console.WriteLine("Processing Claim Activity Code... " + ent.GetAttributeValue<EntityReference>("hil_jobid").Name + "/" +_rowCount.ToString() + "/" + entcoll.Entities.Count.ToString());
                }
            }
        }
        public static void UpdateActivityOnClaimLines(IOrganizationService _service)
        {
            #region Variable declaration
            QueryExpression queryExp;
            EntityCollection entcoll;
            Guid _performaInvoiceId = Guid.Empty;
            EntityReference _erFiscalMonth = null;
            #endregion

            //DateTime _processDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);
            DateTime _processDate = new DateTime(2024, 04, 20);
            queryExp = new QueryExpression("hil_claimperiod");
            queryExp.ColumnSet = new ColumnSet("hil_fromdate", "hil_todate");
            queryExp.Criteria = new FilterExpression(LogicalOperator.And);
            queryExp.Criteria.AddCondition("hil_fromdate", ConditionOperator.OnOrBefore, _processDate);
            queryExp.Criteria.AddCondition("hil_todate", ConditionOperator.OnOrAfter, _processDate);
            entcoll = _service.RetrieveMultiple(queryExp);
            if (entcoll.Entities.Count > 0)
            {
                _erFiscalMonth = entcoll.Entities[0].ToEntityReference();
            }
            if (_erFiscalMonth == null)
            {
                Console.WriteLine("Fiscal Month is not defined." + DateTime.Now.ToString());
                return;
            }
            //FilterExpression _temp = new FilterExpression();
            //_temp.AddCondition("hil_warrantysubstatus", ConditionOperator.Equal, 4); //Under AMC
            //_temp.AddCondition("hil_callsubtype", ConditionOperator.NotEqual, new Guid("706D80B8-D7AA-EE11-A569-6045BDAD2B59"));//Special PMS
            //LinkCriteria = _temp
            LinkEntity lnkEntJB = new LinkEntity
            {
                LinkFromEntityName = "hil_claimline",
                LinkToEntityName = "msdyn_workorder",
                LinkFromAttributeName = "hil_jobid",
                LinkToAttributeName = "msdyn_workorderid",
                Columns = new ColumnSet("hil_productcategory", "hil_callsubtype", "hil_warrantysubstatus", "msdyn_name"),
                EntityAlias = "jb",
                JoinOperator = JoinOperator.Inner
            };
            LinkEntity lnkEntSO = new LinkEntity
            {
                LinkFromEntityName = "account",
                LinkToEntityName = "hil_salesoffice",
                LinkFromAttributeName = "hil_salesoffice",
                LinkToAttributeName = "hil_salesofficeid",
                Columns = new ColumnSet("hil_state"),
                EntityAlias = "so",
                JoinOperator = JoinOperator.Inner
            };

            LinkEntity lnkEntCP = new LinkEntity
            {
                LinkFromEntityName = "hil_claimheader",
                LinkToEntityName = "account",
                LinkFromAttributeName = "hil_franchisee",
                LinkToAttributeName = "accountid",
                Columns = new ColumnSet("hil_state", "hil_salesoffice", "ownerid"),
                EntityAlias = "cp",
                JoinOperator = JoinOperator.Inner
            };

            LinkEntity lnkEntCH = new LinkEntity
            {
                LinkFromEntityName = "hil_claimline",
                LinkToEntityName = "hil_claimheader",
                LinkFromAttributeName = "hil_claimheader",
                LinkToAttributeName = "hil_claimheaderid",
                Columns = new ColumnSet("hil_performastatus"),
                EntityAlias = "ch",
                JoinOperator = JoinOperator.Inner
            };

            lnkEntCP.LinkEntities.Add(lnkEntSO);
            lnkEntCH.LinkEntities.Add(lnkEntCP);

            queryExp = new QueryExpression("hil_claimline");
            queryExp.ColumnSet = new ColumnSet("createdon", "hil_franchisee", "hil_claimperiod", "hil_claimheader", "hil_claimamount", "hil_activitycode");
            queryExp.Criteria = new FilterExpression(LogicalOperator.And);
            queryExp.Criteria.AddCondition("hil_claimperiod", ConditionOperator.Equal, _erFiscalMonth.Id); // Fiscal Month
            queryExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
            queryExp.Criteria.AddCondition("hil_jobid", ConditionOperator.NotNull);
            queryExp.Criteria.AddCondition("hil_jobid", ConditionOperator.Equal, new Guid("348ffb70-37e7-ee11-a204-6045bdac5a1d"));
            queryExp.Criteria.AddCondition("hil_activitycode", ConditionOperator.Null);
            queryExp.LinkEntities.Add(lnkEntCH);
            queryExp.LinkEntities.Add(lnkEntJB);
            queryExp.NoLock = true;
            queryExp.AddOrder("createdon", OrderType.Ascending);
            entcoll = _service.RetrieveMultiple(queryExp);
            if (entcoll.Entities.Count > 0)
            {
                int _gstCatg;
                EntityReference _erCP = null;
                QueryExpression queryExpTemp;
                EntityCollection entcollTemp;
                Entity _updateClaimLine = null;
                string _activityCode = string.Empty;
                int _rowCount = 0;
                foreach (Entity ent in entcoll.Entities)
                {
                    _erCP = ent.GetAttributeValue<EntityReference>("hil_franchisee");
                    _activityCode = string.Empty;
                    OptionSetValue _osWarrantySubstatus = null;

                    if (ent.Contains("jb.hil_warrantysubstatus"))
                    {
                        _osWarrantySubstatus = ((OptionSetValue)ent.GetAttributeValue<AliasedValue>("jb.hil_warrantysubstatus").Value);
                    }
                    else { 
                    
                    }
                    if (ent.Contains("cp.hil_state") && ent.Contains("so.hil_state") && ent.Contains("jb.hil_productcategory") && ent.Contains("jb.hil_callsubtype") && ent.Contains("jb.hil_warrantysubstatus"))
                    {
                        EntityReference _erCPState = ((EntityReference)ent.GetAttributeValue<AliasedValue>("cp.hil_state").Value);
                        EntityReference _erSOState = ((EntityReference)ent.GetAttributeValue<AliasedValue>("so.hil_state").Value);
                        EntityReference _erProdCategory = ((EntityReference)ent.GetAttributeValue<AliasedValue>("jb.hil_productcategory").Value);
                        EntityReference _erCallSubtype = ((EntityReference)ent.GetAttributeValue<AliasedValue>("jb.hil_callsubtype").Value);
                        //OptionSetValue _osWarrantySubstatus = ((OptionSetValue)ent.GetAttributeValue<AliasedValue>("jb.hil_warrantysubstatus").Value);

                        _gstCatg = _erCPState.Id == _erSOState.Id ? 1 : 2;
                        Console.WriteLine(_erCallSubtype.Name);
                        queryExpTemp = new QueryExpression("hil_claimpostingsetup");
                        queryExpTemp.ColumnSet = new ColumnSet("hil_activitycode");
                        queryExpTemp.Criteria = new FilterExpression(LogicalOperator.And);
                        queryExpTemp.Criteria.AddCondition("hil_callsubtype", ConditionOperator.Equal, _erCallSubtype.Id); //Breakdown
                        queryExpTemp.Criteria.AddCondition("hil_productcategory", ConditionOperator.Equal, _erProdCategory.Id);
                        queryExpTemp.Criteria.AddCondition("hil_activitygstslab", ConditionOperator.Equal, _gstCatg);
                        queryExpTemp.Criteria.AddCondition("hil_warrantysubstatus", ConditionOperator.Equal, _osWarrantySubstatus.Value);
                        queryExpTemp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                        entcollTemp = _service.RetrieveMultiple(queryExpTemp);
                        if (entcollTemp.Entities.Count > 0)
                        {
                            _activityCode = entcollTemp.Entities[0].GetAttributeValue<string>("hil_activitycode");
                            _updateClaimLine = new Entity(ent.LogicalName, ent.Id);
                            _updateClaimLine["hil_callsubtype"] = _erCallSubtype;
                            _updateClaimLine["hil_productcategory"] = _erProdCategory;
                            _updateClaimLine["hil_activitycode"] = _activityCode;
                            _service.Update(_updateClaimLine);
                        }
                        else {
                            _updateClaimLine = new Entity(ent.LogicalName, ent.Id);
                            _updateClaimLine["hil_callsubtype"] = _erCallSubtype;
                            _updateClaimLine["hil_productcategory"] = _erProdCategory;
                            _service.Update(_updateClaimLine);
                        }
                    }
                    _rowCount++;
                    Console.WriteLine("Processing Claim Activity Code... " + ent.GetAttributeValue<DateTime>("createdon").ToString()  + " - "+ _rowCount.ToString() + "/" + entcoll.Entities.Count.ToString());
                }
            }
        }

        public static void UpdateActivityOnClaimLines23Feb2024(IOrganizationService _service,string _jobId)
        {
            #region Variable declaration
            QueryExpression queryExp;
            EntityCollection entcoll;
            Guid _performaInvoiceId = Guid.Empty;
            EntityReference _erFiscalMonth = null;
            #endregion

            //DateTime _processDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);
            DateTime _processDate = new DateTime(2024, 02, 20);
            queryExp = new QueryExpression("hil_claimperiod");
            queryExp.ColumnSet = new ColumnSet("hil_fromdate", "hil_todate");
            queryExp.Criteria = new FilterExpression(LogicalOperator.And);
            queryExp.Criteria.AddCondition("hil_fromdate", ConditionOperator.OnOrBefore, _processDate);
            queryExp.Criteria.AddCondition("hil_todate", ConditionOperator.OnOrAfter, _processDate);
            entcoll = _service.RetrieveMultiple(queryExp);
            if (entcoll.Entities.Count > 0)
            {
                _erFiscalMonth = entcoll.Entities[0].ToEntityReference();
            }
            if (_erFiscalMonth == null)
            {
                Console.WriteLine("Fiscal Month is not defined." + DateTime.Now.ToString());
                return;
            }

            //FilterExpression _temp = new FilterExpression();
            //_temp.AddCondition("hil_warrantysubstatus", ConditionOperator.Equal, 4); //Under AMC
            //_temp.AddCondition("hil_callsubtype", ConditionOperator.NotEqual, new Guid("706D80B8-D7AA-EE11-A569-6045BDAD2B59"));//Special PMS
            LinkEntity lnkEntJB = new LinkEntity
            {
                LinkFromEntityName = "hil_claimline",
                LinkToEntityName = "msdyn_workorder",
                LinkFromAttributeName = "hil_jobid",
                LinkToAttributeName = "msdyn_workorderid",
                Columns = new ColumnSet("hil_productcategory", "hil_callsubtype", "hil_warrantysubstatus", "msdyn_name"),
                EntityAlias = "jb",
                JoinOperator = JoinOperator.Inner
                //,LinkCriteria = _temp
            };
            LinkEntity lnkEntSO = new LinkEntity
            {
                LinkFromEntityName = "account",
                LinkToEntityName = "hil_salesoffice",
                LinkFromAttributeName = "hil_salesoffice",
                LinkToAttributeName = "hil_salesofficeid",
                Columns = new ColumnSet("hil_state"),
                EntityAlias = "so",
                JoinOperator = JoinOperator.Inner
            };

            LinkEntity lnkEntCP = new LinkEntity
            {
                LinkFromEntityName = "hil_claimheader",
                LinkToEntityName = "account",
                LinkFromAttributeName = "hil_franchisee",
                LinkToAttributeName = "accountid",
                Columns = new ColumnSet("hil_state", "hil_salesoffice", "ownerid"),
                EntityAlias = "cp",
                JoinOperator = JoinOperator.Inner
            };

            LinkEntity lnkEntCH = new LinkEntity
            {
                LinkFromEntityName = "hil_claimline",
                LinkToEntityName = "hil_claimheader",
                LinkFromAttributeName = "hil_claimheader",
                LinkToAttributeName = "hil_claimheaderid",
                Columns = new ColumnSet("hil_performastatus"),
                EntityAlias = "ch",
                JoinOperator = JoinOperator.Inner
            };

            lnkEntCP.LinkEntities.Add(lnkEntSO);
            lnkEntCH.LinkEntities.Add(lnkEntCP);

            queryExp = new QueryExpression("hil_claimline");
            queryExp.ColumnSet = new ColumnSet("createdon", "hil_franchisee", "hil_claimperiod", "hil_claimheader", "hil_claimamount", "hil_activitycode");
            queryExp.Criteria = new FilterExpression(LogicalOperator.And);
            queryExp.Criteria.AddCondition("hil_claimperiod", ConditionOperator.Equal, _erFiscalMonth.Id); // Fiscal Month
            queryExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
            queryExp.Criteria.AddCondition("hil_jobid", ConditionOperator.NotNull);
            queryExp.Criteria.AddCondition("hil_jobid", ConditionOperator.Equal, new Guid(_jobId));
            //queryExp.Criteria.AddCondition("hil_activitycode", ConditionOperator.Null);
            queryExp.LinkEntities.Add(lnkEntCH);
            queryExp.LinkEntities.Add(lnkEntJB);
            queryExp.NoLock = true;
            queryExp.AddOrder("createdon", OrderType.Ascending);
            entcoll = _service.RetrieveMultiple(queryExp);
            if (entcoll.Entities.Count > 0)
            {
                int _gstCatg;
                EntityReference _erCP = null;
                QueryExpression queryExpTemp;
                EntityCollection entcollTemp;
                Entity _updateClaimLine = null;
                string _activityCode = string.Empty;
                int _rowCount = 0;
                foreach (Entity ent in entcoll.Entities)
                {
                    _erCP = ent.GetAttributeValue<EntityReference>("hil_franchisee");
                    _activityCode = string.Empty;
                    if (ent.Contains("cp.hil_state") && ent.Contains("so.hil_state") && ent.Contains("jb.hil_productcategory") && ent.Contains("jb.hil_callsubtype") && ent.Contains("jb.hil_warrantysubstatus"))
                    {
                        EntityReference _erCPState = ((EntityReference)ent.GetAttributeValue<AliasedValue>("cp.hil_state").Value);
                        EntityReference _erSOState = ((EntityReference)ent.GetAttributeValue<AliasedValue>("so.hil_state").Value);
                        EntityReference _erProdCategory = ((EntityReference)ent.GetAttributeValue<AliasedValue>("jb.hil_productcategory").Value);
                        EntityReference _erCallSubtype = ((EntityReference)ent.GetAttributeValue<AliasedValue>("jb.hil_callsubtype").Value);
                        OptionSetValue _osWarrantySubstatus = ((OptionSetValue)ent.GetAttributeValue<AliasedValue>("jb.hil_warrantysubstatus").Value);

                        _gstCatg = _erCPState.Id == _erSOState.Id ? 1 : 2;
                        Console.WriteLine(_erCallSubtype.Name);
                        queryExpTemp = new QueryExpression("hil_claimpostingsetup");
                        queryExpTemp.ColumnSet = new ColumnSet("hil_activitycode");
                        queryExpTemp.Criteria = new FilterExpression(LogicalOperator.And);
                        queryExpTemp.Criteria.AddCondition("hil_callsubtype", ConditionOperator.Equal, _erCallSubtype.Id); //Breakdown
                        queryExpTemp.Criteria.AddCondition("hil_productcategory", ConditionOperator.Equal, _erProdCategory.Id);
                        queryExpTemp.Criteria.AddCondition("hil_activitygstslab", ConditionOperator.Equal, _gstCatg);
                        queryExpTemp.Criteria.AddCondition("hil_warrantysubstatus", ConditionOperator.Equal, _osWarrantySubstatus.Value);
                        queryExpTemp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                        entcollTemp = _service.RetrieveMultiple(queryExpTemp);
                        if (entcollTemp.Entities.Count > 0)
                        {
                            _activityCode = entcollTemp.Entities[0].GetAttributeValue<string>("hil_activitycode");
                            _updateClaimLine = new Entity(ent.LogicalName, ent.Id);
                            _updateClaimLine["hil_callsubtype"] = _erCallSubtype;
                            _updateClaimLine["hil_productcategory"] = _erProdCategory;
                            _updateClaimLine["hil_activitycode"] = _activityCode;
                            _service.Update(_updateClaimLine);
                        }
                        else
                        {
                            _updateClaimLine = new Entity(ent.LogicalName, ent.Id);
                            _updateClaimLine["hil_callsubtype"] = _erCallSubtype;
                            _updateClaimLine["hil_productcategory"] = _erProdCategory;
                            _service.Update(_updateClaimLine);
                        }
                    }
                    _rowCount++;
                    Console.WriteLine("Processing Claim Activity Code... " + ent.GetAttributeValue<DateTime>("createdon").ToString() + " - " + _rowCount.ToString() + "/" + entcoll.Entities.Count.ToString());
                }
            }
        }
        public static void UpdateActivityOnClaimLines1(IOrganizationService _service, Guid _workOrderId, Entity entity)
        {
            #region Variable declaration
            QueryExpression queryExp;
            EntityCollection entcoll;
            Guid _performaInvoiceId = Guid.Empty;
            #endregion


            LinkEntity lnkEntSO = new LinkEntity
            {
                LinkFromEntityName = "account",
                LinkToEntityName = "hil_salesoffice",
                LinkFromAttributeName = "hil_salesoffice",
                LinkToAttributeName = "hil_salesofficeid",
                Columns = new ColumnSet("hil_state"),
                EntityAlias = "so",
                JoinOperator = JoinOperator.Inner
            };

            LinkEntity lnkEntCP = new LinkEntity
            {
                LinkFromEntityName = "msdyn_workorder",
                LinkToEntityName = "account",
                LinkFromAttributeName = "hil_owneraccount",
                LinkToAttributeName = "accountid",
                Columns = new ColumnSet("hil_state", "hil_salesoffice", "ownerid"),
                EntityAlias = "cp",
                JoinOperator = JoinOperator.Inner
            };

            lnkEntCP.LinkEntities.Add(lnkEntSO);

            queryExp = new QueryExpression("msdyn_workorder");
            queryExp.ColumnSet = new ColumnSet("hil_productcategory", "hil_callsubtype", "hil_owneraccount");
            queryExp.Criteria = new FilterExpression(LogicalOperator.And);
            queryExp.Criteria.AddCondition("msdyn_workorderid", ConditionOperator.Equal, _workOrderId);
            queryExp.LinkEntities.Add(lnkEntCP);
            entcoll = _service.RetrieveMultiple(queryExp);
            if (entcoll.Entities.Count > 0)
            {
                int _gstCatg;
                QueryExpression queryExpTemp;
                EntityCollection entcollTemp;
                string _activityCode = string.Empty;
                foreach (Entity ent in entcoll.Entities)
                {
                    _activityCode = string.Empty;
                    if (ent.Contains("cp.hil_state") && ent.Contains("so.hil_state") && ent.Contains("hil_productcategory") && ent.Contains("hil_callsubtype"))
                    {
                        EntityReference _erCPState = ((EntityReference)ent.GetAttributeValue<AliasedValue>("cp.hil_state").Value);
                        EntityReference _erSOState = ((EntityReference)ent.GetAttributeValue<AliasedValue>("so.hil_state").Value);
                        EntityReference _erProdCategory = ent.GetAttributeValue<EntityReference>("hil_productcategory");
                        EntityReference _erCallSubtype = ent.GetAttributeValue<EntityReference>("hil_callsubtype");

                        _gstCatg = _erCPState.Id == _erSOState.Id ? 1 : 2;

                        queryExpTemp = new QueryExpression("hil_claimpostingsetup");
                        queryExpTemp.ColumnSet = new ColumnSet("hil_activitycode");
                        queryExpTemp.Criteria = new FilterExpression(LogicalOperator.And);
                        queryExpTemp.Criteria.AddCondition("hil_callsubtype", ConditionOperator.Equal, _erCallSubtype.Id); //Breakdown
                        queryExpTemp.Criteria.AddCondition("hil_productcategory", ConditionOperator.Equal, _erProdCategory.Id);
                        queryExpTemp.Criteria.AddCondition("hil_activitygstslab", ConditionOperator.Equal, _gstCatg);
                        queryExpTemp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                        entcollTemp = _service.RetrieveMultiple(queryExpTemp);
                        if (entcollTemp.Entities.Count > 0)
                        {
                            _activityCode = entcollTemp.Entities[0].GetAttributeValue<string>("hil_activitycode");
                            entity["hil_callsubtype"] = _erCallSubtype;
                            entity["hil_productcategory"] = _erProdCategory;
                            entity["hil_activitycode"] = _activityCode;
                        }
                        else
                        {
                            entity["hil_callsubtype"] = _erCallSubtype;
                            entity["hil_productcategory"] = _erProdCategory;
                        }
                    }
                }
            }
        }

        public static void GenerateClaimSummary(IOrganizationService _service, Guid _performaInvoiceId)
        {
            QueryExpression queryExp;
            EntityCollection entcoll;

            LinkEntity lnkEntSO = new LinkEntity
            {
                LinkFromEntityName = "account",
                LinkToEntityName = "hil_salesoffice",
                LinkFromAttributeName = "hil_salesoffice",
                LinkToAttributeName = "hil_salesofficeid",
                Columns = new ColumnSet("hil_sapcode"),
                EntityAlias = "so",
                JoinOperator = JoinOperator.Inner
            };

            LinkEntity lnkEntCP = new LinkEntity
            {
                LinkFromEntityName = "hil_claimheader",
                LinkToEntityName = "account",
                LinkFromAttributeName = "hil_franchisee",
                LinkToAttributeName = "accountid",
                Columns = new ColumnSet("hil_salesoffice", "ownerid", "hil_vendorcode"),
                EntityAlias = "cp",
                JoinOperator = JoinOperator.Inner
            };


            lnkEntCP.LinkEntities.Add(lnkEntSO);

            queryExp = new QueryExpression("hil_claimheader");
            queryExp.ColumnSet = new ColumnSet("hil_fiscalmonth", "hil_franchisee", "hil_performastatus");
            queryExp.Criteria = new FilterExpression(LogicalOperator.And);
            queryExp.Criteria.AddCondition("hil_claimheaderid", ConditionOperator.Equal, _performaInvoiceId);
            queryExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
            queryExp.LinkEntities.Add(lnkEntCP);

            entcoll = _service.RetrieveMultiple(queryExp);
            if (entcoll.Entities.Count > 0)
            {
                string _vendorCode = string.Empty;
                string _fetchXML = string.Empty;
                if (entcoll.Entities[0].Contains("cp.hil_vendorcode") && entcoll.Entities[0].Contains("so.hil_sapcode"))
                {
                    #region Deactivating all Claim Summary Lines
                    _fetchXML = @"<fetch distinct='false' mapping='logical'>
                        <entity name='hil_claimsummary'>  
                        <attribute name='hil_claimsummaryid' />
                        <filter type='and'>
                            <condition attribute='hil_performainvoiceid' operator='eq' value='{" + _performaInvoiceId + @"}' />
                            <condition attribute='statecode' operator='eq' value='0' />
                        </filter>
                        </entity>
                    </fetch>";
                    EntityCollection entcollCS = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                    foreach (Entity ent in entcollCS.Entities)
                    {
                        SetStateRequest setStateRequest = new SetStateRequest()
                        {
                            EntityMoniker = new EntityReference
                            {
                                Id = ent.Id,
                                LogicalName = ent.LogicalName,
                            },
                            State = new OptionSetValue(1),
                            Status = new OptionSetValue(2)
                        };
                        _service.Execute(setStateRequest);
                    }
                    #endregion

                    int _rowCount = 0;
                    _fetchXML = @"<fetch distinct='false' mapping='logical' aggregate='true'>
                      <entity name='hil_claimline'>
                        <attribute name='hil_activitycode' groupby='true' alias='activity'/>
                        <attribute name='hil_claimamount' alias='claimamount' aggregate='sum'/>
                        <filter type='and'>
                          <condition attribute='hil_claimheader' operator='eq' value='{" + _performaInvoiceId + @"}' />
                          <condition attribute='hil_activitycode' operator='not-null' />
                          <condition attribute='statecode' operator='eq' value='0' />
                          <condition attribute='hil_productcategory' operator='not-null' />
                        </filter>
                        <link-entity name='product' from='productid' to='hil_productcategory' visible='false' link-type='outer' alias='pd'>
                          <attribute name='hil_sapcode' groupby='true' alias='prodcode'/>
                        </link-entity>
                      </entity>
                    </fetch>";
                    EntityCollection entcoll1 = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                    Entity _entClaimSum = null;
                    Money _claimAmount = null;

                    foreach (Entity ent in entcoll1.Entities)
                    {
                        _claimAmount = ((Money)ent.GetAttributeValue<AliasedValue>("claimamount").Value);
                        if (_claimAmount.Value == 0)
                            continue;

                        _entClaimSum = new Entity("hil_claimsummary");

                        _entClaimSum["hil_salesoffice"] = (EntityReference)entcoll.Entities[0].GetAttributeValue<AliasedValue>("cp.hil_salesoffice").Value;
                        _entClaimSum["hil_performainvoiceid"] = entcoll.Entities[0].ToEntityReference();
                        _entClaimSum["hil_franchise"] = entcoll.Entities[0].GetAttributeValue<EntityReference>("hil_franchisee");
                        _entClaimSum["hil_claimmonth"] = entcoll.Entities[0].GetAttributeValue<EntityReference>("hil_fiscalmonth");
                        _entClaimSum["hil_franchisecode"] = entcoll.Entities[0].GetAttributeValue<AliasedValue>("cp.hil_vendorcode").Value.ToString();
                        _entClaimSum["hil_claimmonthcode"] = entcoll.Entities[0].GetAttributeValue<EntityReference>("hil_fiscalmonth").Name;
                        _entClaimSum["hil_salesofficecode"] = entcoll.Entities[0].GetAttributeValue<AliasedValue>("so.hil_sapcode").Value.ToString();
                        _entClaimSum["hil_activitycode"] = ent.GetAttributeValue<AliasedValue>("activity").Value.ToString();
                        _entClaimSum["hil_productdivision"] = ent.GetAttributeValue<AliasedValue>("prodcode").Value.ToString();
                        _entClaimSum["hil_claimamount"] = _claimAmount;
                        _entClaimSum["hil_qty"] = 1;
                        _service.Create(_entClaimSum);

                        _rowCount++;
                    }

                    #region Adjusting Negative Claim amount with Poisitive
                    _fetchXML = @"<fetch distinct='false' mapping='logical' aggregate='true'>
                      <entity name='hil_claimsummary'>
                        <attribute name='hil_claimamount' alias='claimamount' aggregate='sum'/>
                        <filter type='and'>
                          <condition attribute='hil_performainvoiceid' operator='eq' value='{" + _performaInvoiceId + @"}' />
                          <condition attribute='hil_claimamount' operator='lt' value='0' />
                        </filter>
                      </entity>
                    </fetch>";
                    entcoll1 = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                    if (entcoll1.Entities.Count > 0)
                    {
                        if (entcoll1.Entities[0].Contains("claimamount"))
                        {
                            if (entcoll1.Entities[0].GetAttributeValue<AliasedValue>("claimamount").Value != null)
                            {
                                Money _negativeClaimAmount = null;
                                _negativeClaimAmount = ((Money)entcoll1.Entities[0].GetAttributeValue<AliasedValue>("claimamount").Value);

                                _fetchXML = @"<fetch distinct='false' mapping='logical'>
                                    <entity name='hil_claimsummary'>
                                    <attribute name='hil_claimamount' />
                                    <filter type='and'>
                                        <condition attribute='hil_performainvoiceid' operator='eq' value='{" + _performaInvoiceId + @"}' />
                                        <condition attribute='hil_claimamount' operator='ge' value='" + (_negativeClaimAmount.Value * -1) + @"' />
                                    </filter>
                                    </entity>
                                    </fetch>";
                                entcoll1 = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                                if (entcoll1.Entities.Count > 0)
                                {
                                    _claimAmount = entcoll1.Entities[0].GetAttributeValue<Money>("hil_claimamount");
                                    entcoll1.Entities[0]["hil_claimamount"] = _claimAmount.Value - (_negativeClaimAmount.Value * -1);
                                    _service.Update(entcoll1.Entities[0]);
                                }

                                _fetchXML = @"<fetch distinct='false' mapping='logical'>
                                  <entity name='hil_claimsummary'>  
                                    <attribute name='hil_claimsummaryid' />
                                    <filter type='and'>
                                      <condition attribute='hil_performainvoiceid' operator='eq' value='{" + _performaInvoiceId + @"}' />
                                      <condition attribute='hil_claimamount' operator='lt' value='0' />
                                    </filter>
                                  </entity>
                                </fetch>";
                                entcoll1 = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                                foreach (Entity ent in entcoll1.Entities)
                                {
                                    SetStateRequest setStateRequest = new SetStateRequest()
                                    {
                                        EntityMoniker = new EntityReference
                                        {
                                            Id = ent.Id,
                                            LogicalName = ent.LogicalName,
                                        },
                                        State = new OptionSetValue(1),
                                        Status = new OptionSetValue(2)
                                    };
                                    _service.Execute(setStateRequest);
                                }
                            }
                        }
                    }
                    #endregion

                }
            }
        }

        //public static void GenerateClaimSummary(IOrganizationService _service,Guid _performaInvoiceId)
        //{
        //    QueryExpression queryExp;
        //    EntityCollection entcoll = null;
        //    string _fetchXML = string.Empty;
        //    EntityCollection entcoll1 = null;
        //    LinkEntity lnkEntSO = new LinkEntity
        //    {
        //        LinkFromEntityName = "account",
        //        LinkToEntityName = "hil_salesoffice",
        //        LinkFromAttributeName = "hil_salesoffice",
        //        LinkToAttributeName = "hil_salesofficeid",
        //        Columns = new ColumnSet("hil_sapcode"),
        //        EntityAlias = "so",
        //        JoinOperator = JoinOperator.Inner
        //    };

        //    LinkEntity lnkEntCP = new LinkEntity
        //    {
        //        LinkFromEntityName = "hil_claimheader",
        //        LinkToEntityName = "account",
        //        LinkFromAttributeName = "hil_franchisee",
        //        LinkToAttributeName = "accountid",
        //        Columns = new ColumnSet("hil_salesoffice", "ownerid", "hil_vendorcode"),
        //        EntityAlias = "cp",
        //        JoinOperator = JoinOperator.Inner
        //    };

        //    lnkEntCP.LinkEntities.Add(lnkEntSO);

        //    queryExp = new QueryExpression("hil_claimheader");
        //    queryExp.ColumnSet = new ColumnSet("hil_fiscalmonth", "hil_franchisee", "hil_performastatus");
        //    queryExp.Criteria = new FilterExpression(LogicalOperator.And);
        //    queryExp.Criteria.AddCondition("hil_claimheaderid", ConditionOperator.Equal, _performaInvoiceId);
        //    queryExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
        //    queryExp.LinkEntities.Add(lnkEntCP);

        //    entcoll = _service.RetrieveMultiple(queryExp);
        //    if (entcoll.Entities.Count > 0)
        //    {
        //        string _vendorCode = string.Empty;
        //        if (entcoll.Entities[0].Contains("cp.hil_vendorcode") && entcoll.Entities[0].Contains("so.hil_sapcode"))
        //        {
        //            int _rowCount = 0;
        //            _fetchXML = @"<fetch distinct='false' mapping='logical' aggregate='true'>
        //              <entity name='hil_claimline'>
        //                <attribute name='hil_activitycode' groupby='true' alias='activity'/>
        //                <attribute name='hil_claimamount' alias='claimamount' aggregate='sum'/>
        //                <filter type='and'>
        //                  <condition attribute='hil_claimheader' operator='eq' value='{" + _performaInvoiceId + @"}' />
        //                  <condition attribute='hil_activitycode' operator='not-null' />
        //                  <condition attribute='statecode' operator='eq' value='0' />
        //                  <condition attribute='hil_productcategory' operator='not-null' />
        //                </filter>
        //                <link-entity name='product' from='productid' to='hil_productcategory' visible='false' link-type='outer' alias='pd'>
        //                  <attribute name='hil_sapcode' groupby='true' alias='prodcode'/>
        //                </link-entity>
        //              </entity>
        //            </fetch>";
        //            entcoll1 = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
        //            Entity _entClaimSum = null;
        //            Money _claimAmount = null;
        //            foreach (Entity ent in entcoll1.Entities)
        //            {
        //                _claimAmount = ((Money)ent.GetAttributeValue<AliasedValue>("claimamount").Value);
        //                if (_claimAmount.Value == 0)
        //                    continue;

        //                //_entClaimSum = new Entity("hil_claimsummary");
        //                //_entClaimSum["hil_salesoffice"] = (EntityReference)entcoll.Entities[0].GetAttributeValue<AliasedValue>("cp.hil_salesoffice").Value;
        //                //_entClaimSum["hil_performainvoiceid"] = entcoll.Entities[0].ToEntityReference();
        //                //_entClaimSum["hil_franchise"] = entcoll.Entities[0].GetAttributeValue<EntityReference>("hil_franchisee");
        //                //_entClaimSum["hil_claimmonth"] = entcoll.Entities[0].GetAttributeValue<EntityReference>("hil_fiscalmonth");
        //                //_entClaimSum["hil_franchisecode"] = entcoll.Entities[0].GetAttributeValue<AliasedValue>("cp.hil_vendorcode").Value.ToString();
        //                //_entClaimSum["hil_claimmonthcode"] = entcoll.Entities[0].GetAttributeValue<EntityReference>("hil_fiscalmonth").Name;
        //                //_entClaimSum["hil_salesofficecode"] = entcoll.Entities[0].GetAttributeValue<AliasedValue>("so.hil_sapcode").Value.ToString();
        //                //_entClaimSum["hil_activitycode"] = ent.GetAttributeValue<AliasedValue>("activity").Value.ToString();
        //                //_entClaimSum["hil_productdivision"] = ent.GetAttributeValue<AliasedValue>("prodcode").Value.ToString();
        //                //_entClaimSum["hil_claimamount"] = _claimAmount;
        //                //_entClaimSum["hil_qty"] = 1;
        //                //_service.Create(_entClaimSum);

        //                _rowCount++;
        //            }

        //            //#region Adjusting Negative Claim amount with Poisitive
        //            //_fetchXML = @"<fetch distinct='false' mapping='logical' aggregate='true'>
        //            //  <entity name='hil_claimsummary'>
        //            //    <attribute name='hil_claimamount' alias='claimamount' aggregate='sum'/>
        //            //    <filter type='and'>
        //            //      <condition attribute='hil_performainvoiceid' operator='eq' value='{" + _performaInvoiceId + @"}' />
        //            //      <condition attribute='hil_claimamount' operator='lt' value='0' />
        //            //    </filter>
        //            //  </entity>
        //            //</fetch>";
        //            //entcoll1 = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
        //            //if (entcoll1.Entities.Count > 0)
        //            //{
        //            //    if (entcoll1.Entities[0].Contains("claimamount"))
        //            //    {
        //            //        if (entcoll1.Entities[0].GetAttributeValue<AliasedValue>("claimamount").Value != null)
        //            //        {
        //            //            Money _negativeClaimAmount = null;
        //            //            _negativeClaimAmount = ((Money)entcoll1.Entities[0].GetAttributeValue<AliasedValue>("claimamount").Value);

        //            //            _fetchXML = @"<fetch distinct='false' mapping='logical'>
        //            //            <entity name='hil_claimsummary'>
        //            //            <attribute name='hil_claimamount' />
        //            //            <filter type='and'>
        //            //                <condition attribute='hil_performainvoiceid' operator='eq' value='{" + _performaInvoiceId + @"}' />
        //            //                <condition attribute='hil_claimamount' operator='ge' value='" + (_negativeClaimAmount.Value * -1) + @"' />
        //            //            </filter>
        //            //            </entity>
        //            //            </fetch>";
        //            //            entcoll1 = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
        //            //            if (entcoll1.Entities.Count > 0)
        //            //            {
        //            //                _claimAmount = entcoll1.Entities[0].GetAttributeValue<Money>("hil_claimamount");
        //            //                entcoll1.Entities[0]["hil_claimamount"] = _claimAmount.Value - (_negativeClaimAmount.Value * -1);
        //            //                _service.Update(entcoll1.Entities[0]);
        //            //            }

        //            //            _fetchXML = @"<fetch distinct='false' mapping='logical'>
        //            //              <entity name='hil_claimsummary'>  
        //            //                <attribute name='hil_claimsummaryid' />
        //            //                <filter type='and'>
        //            //                  <condition attribute='hil_performainvoiceid' operator='eq' value='{" + _performaInvoiceId + @"}' />
        //            //                  <condition attribute='hil_claimamount' operator='lt' value='0' />
        //            //                </filter>
        //            //              </entity>
        //            //            </fetch>";
        //            //            entcoll1 = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
        //            //            foreach (Entity ent in entcoll1.Entities)
        //            //            {
        //            //                SetStateRequest setStateRequest = new SetStateRequest()
        //            //                {
        //            //                    EntityMoniker = new EntityReference
        //            //                    {
        //            //                        Id = entcoll1.Entities[0].Id,
        //            //                        LogicalName = entcoll1.Entities[0].LogicalName,
        //            //                    },
        //            //                    State = new OptionSetValue(1),
        //            //                    Status = new OptionSetValue(2)
        //            //                };
        //            //                _service.Execute(setStateRequest);
        //            //            }
        //            //        }
        //            //    }
        //            //}
        //            //#endregion
        //        }
        //    }
        //}
        public static void GenerateFixedCompensationLines(IOrganizationService _service, Guid _performaInvoiceId)
        {
            QueryExpression queryExp;
            EntityCollection entcoll;

            LinkEntity lnkEntCH = new LinkEntity
            {
                LinkFromEntityName = "hil_claimline",
                LinkToEntityName = "hil_claimheader",
                LinkFromAttributeName = "hil_claimheader",
                LinkToAttributeName = "hil_claimheaderid",
                Columns = new ColumnSet("hil_performastatus"),
                EntityAlias = "ch",
                JoinOperator = JoinOperator.Inner
            };
            LinkEntity lnkEntSO = new LinkEntity
            {
                LinkFromEntityName = "account",
                LinkToEntityName = "hil_salesoffice",
                LinkFromAttributeName = "hil_salesoffice",
                LinkToAttributeName = "hil_salesofficeid",
                Columns = new ColumnSet("hil_state"),
                EntityAlias = "so",
                JoinOperator = JoinOperator.Inner
            };

            LinkEntity lnkEntCP = new LinkEntity
            {
                LinkFromEntityName = "hil_claimheader",
                LinkToEntityName = "account",
                LinkFromAttributeName = "hil_franchisee",
                LinkToAttributeName = "accountid",
                Columns = new ColumnSet("hil_state", "hil_salesoffice", "ownerid"),
                EntityAlias = "cp",
                JoinOperator = JoinOperator.Inner
            };
            lnkEntCP.LinkEntities.Add(lnkEntSO);
            lnkEntCH.LinkEntities.Add(lnkEntCP);
            queryExp = new QueryExpression("hil_claimline");
            queryExp.ColumnSet = new ColumnSet("hil_franchisee", "hil_claimperiod", "hil_claimheader", "hil_claimamount");
            queryExp.Criteria = new FilterExpression(LogicalOperator.And);
            queryExp.Criteria.AddCondition("hil_claimheader", ConditionOperator.Equal, _performaInvoiceId);
            queryExp.Criteria.AddCondition("hil_claimcategory", ConditionOperator.Equal,new Guid("824A1BCC-6EE5-EA11-A817-000D3AF0501C")); //Fixed Compensation
            queryExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
            queryExp.LinkEntities.Add(lnkEntCH);

            entcoll = _service.RetrieveMultiple(queryExp);
            if (entcoll.Entities.Count > 0)
            {
                int _gstCatg;
                decimal _amount = entcoll.Entities[0].GetAttributeValue<Money>("hil_claimamount").Value;
                EntityReference _erCP = entcoll.Entities[0].GetAttributeValue<EntityReference>("hil_franchisee");
                EntityReference _erCM = entcoll.Entities[0].GetAttributeValue<EntityReference>("hil_claimperiod");
                EntityReference _erCH = entcoll.Entities[0].GetAttributeValue<EntityReference>("hil_claimheader");
                Guid _ownerId = ((EntityReference)entcoll.Entities[0].GetAttributeValue<AliasedValue>("cp.ownerid").Value).Id;
                int jobCount = GetJobCountClaimMonth(_service, _erCP.Id, _erCM.Id);
                if (jobCount > 0)
                {
                    EntityCollection entColProdCatg = GetProdCatgWiseJobCountClaimMonth(_service, _erCP.Id, _erCM.Id);

                    if (entcoll.Entities[0].Contains("cp.hil_state") && entcoll.Entities[0].Contains("so.hil_state"))
                    {
                        EntityReference _erCPState = ((EntityReference)entcoll.Entities[0].GetAttributeValue<AliasedValue>("cp.hil_state").Value);
                        EntityReference _erSOState = ((EntityReference)entcoll.Entities[0].GetAttributeValue<AliasedValue>("so.hil_state").Value);
                        _gstCatg = _erCPState.Id == _erSOState.Id ? 1 : 2;
                        QueryExpression queryExpTemp;
                        EntityCollection entcollTemp;
                        string _activityCode = string.Empty;
                        string _claimCategory = string.Empty;
                        int _prodCatgJobs = 0;
                        decimal _prodCatgAmount = 0;
                        decimal _prodCatgTotalAmount = 0;
                        int _rowCount = 1;
                        foreach (Entity ent in entColProdCatg.Entities)
                        {
                            _activityCode = "";
                            _claimCategory = "";
                            _prodCatgJobs = 0;
                            _prodCatgAmount = 0;

                            _prodCatgJobs = Convert.ToInt32(ent.GetAttributeValue<AliasedValue>("rowcount").Value);
                            
                            _prodCatgAmount = Math.Round((_amount * _prodCatgJobs) / jobCount, 0);

                            _prodCatgTotalAmount += _prodCatgAmount;
                            if (_rowCount == entColProdCatg.Entities.Count) {
                                _prodCatgAmount = _prodCatgAmount + (_amount - _prodCatgTotalAmount);
                            }

                            EntityReference _erProdCatg = (EntityReference)ent.GetAttributeValue<AliasedValue>("prodcatg").Value;

                            queryExpTemp = new QueryExpression("hil_claimpostingsetup");
                            queryExpTemp.ColumnSet = new ColumnSet("hil_activitycode");
                            queryExpTemp.Criteria = new FilterExpression(LogicalOperator.And);
                            queryExpTemp.Criteria.AddCondition("hil_callsubtype", ConditionOperator.Equal, new Guid("6560565A-3C0B-E911-A94E-000D3AF06CD4")); //Breakdown
                            queryExpTemp.Criteria.AddCondition("hil_productcategory", ConditionOperator.Equal, _erProdCatg.Id);
                            queryExpTemp.Criteria.AddCondition("hil_activitygstslab", ConditionOperator.Equal, _gstCatg);
                            queryExpTemp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                            entcollTemp = _service.RetrieveMultiple(queryExpTemp);
                            if (entcollTemp.Entities.Count > 0)
                            {
                                _activityCode = entcollTemp.Entities[0].GetAttributeValue<string>("hil_activitycode");
                            }
                            _claimCategory = "824a1bcc-6ee5-ea11-a817-000d3af0501c"; //Fixed Compensation
                            ClaimLineDTO _claimLine = new ClaimLineDTO()
                            {
                                CallSubtype = new EntityReference("hil_callsubtype", new Guid("6560565A-3C0B-E911-A94E-000D3AF06CD4")), //Breakdown
                                ChannelPartner = _erCP,
                                ClaimAmount = _prodCatgAmount,
                                ClaimCatwgory = _claimCategory,
                                ClaimPeriod = _erCM,
                                OwnerId = _ownerId,
                                PerformInvoiceId = _erCH,
                                ProdCatg = _erProdCatg,
                                ActivityCode = _activityCode
                            };
                            CreateOverheadClaimLine(_service, _claimLine);
                            _rowCount += 1;
                        }

                        #region Inactivate Fixed Compensation Line
                        SetStateRequest setStateRequest = new SetStateRequest()
                        {
                            EntityMoniker = new EntityReference
                            {
                                Id = entcoll.Entities[0].Id,
                                LogicalName = entcoll.Entities[0].LogicalName,
                            },
                            State = new OptionSetValue(1),
                            Status = new OptionSetValue(2)
                        };
                        _service.Execute(setStateRequest);
                        #endregion
                    }
                }
            }
        }
        private static Guid ClaimLineExistsOverhead(IOrganizationService _service, ClaimLineDTO _claimLineData)
        {
            Guid _returnVal = Guid.Empty;
            QueryExpression queryExp;
            EntityCollection entcoll;

            try
            {
                queryExp = new QueryExpression("hil_claimline");
                queryExp.ColumnSet = new ColumnSet("hil_claimperiod");
                queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                queryExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                queryExp.Criteria.AddCondition("hil_franchisee", ConditionOperator.Equal, _claimLineData.ChannelPartner.Id);
                queryExp.Criteria.AddCondition("hil_claimperiod", ConditionOperator.Equal, _claimLineData.ClaimPeriod.Id);
                queryExp.Criteria.AddCondition("hil_productcategory", ConditionOperator.Equal, _claimLineData.ProdCatg.Id);
                queryExp.Criteria.AddCondition("hil_callsubtype", ConditionOperator.Equal, _claimLineData.CallSubtype.Id);

                if (_claimLineData.ClaimCatwgory != null)
                    queryExp.Criteria.AddCondition("hil_claimcategory", ConditionOperator.Equal, new Guid(_claimLineData.ClaimCatwgory));

                queryExp.Criteria.AddCondition("hil_claimheader", ConditionOperator.Equal, _claimLineData.PerformInvoiceId.Id);
                queryExp.Criteria.AddCondition("hil_claimamount", ConditionOperator.Equal, _claimLineData.ClaimAmount);
                entcoll = _service.RetrieveMultiple(queryExp);
                if (entcoll.Entities.Count > 0)
                {
                    _returnVal = entcoll.Entities[0].Id;
                }
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
            return _returnVal;
        }
        private static Guid CreateOverheadClaimLine(IOrganizationService _service, ClaimLineDTO _claimLineData)
        {
            Guid _claimLineId = Guid.Empty;
            try
            {
                Entity entObj = new Entity("hil_claimline");
                _claimLineId = ClaimLineExistsOverhead(_service, _claimLineData);
                if (_claimLineId == Guid.Empty)
                {
                    Entity _entClaimCatg = _service.Retrieve("hil_claimcategory", new Guid(_claimLineData.ClaimCatwgory), new ColumnSet("hil_claimtype"));
                    entObj["hil_franchisee"] = _claimLineData.ChannelPartner;
                    entObj["hil_claimperiod"] = _claimLineData.ClaimPeriod;
                    entObj["hil_productcategory"] = _claimLineData.ProdCatg;
                    entObj["hil_callsubtype"] = _claimLineData.CallSubtype;
                    entObj["hil_claimcategory"] = new EntityReference("hil_claimcategory", new Guid(_claimLineData.ClaimCatwgory));
                    entObj["hil_claimheader"] = _claimLineData.PerformInvoiceId;
                    if (_entClaimCatg.GetAttributeValue<EntityReference>("hil_claimtype").Name == "Penalty")
                        entObj["hil_claimamount"] = new Money(_claimLineData.ClaimAmount * -1);
                    else
                        entObj["hil_claimamount"] = new Money(_claimLineData.ClaimAmount);

                    entObj["hil_activitycode"] = _claimLineData.ActivityCode;
                    _claimLineId = _service.Create(entObj);

                    AssignRecord("hil_claimline", _claimLineData.OwnerId, _claimLineId, _service);
                }
                else
                {
                    entObj.Id = _claimLineId;
                    entObj["hil_activitycode"] = _claimLineData.ActivityCode;
                    entObj["hil_claimamount"] = new Money(_claimLineData.ClaimAmount);
                    _service.Update(entObj);
                }
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
            return _claimLineId;
        }

        private static void CreateClaimLine(IOrganizationService _service, Guid _performInvoiceId, decimal _claimAmount, string _claimCatwgory, EntityReference _erClaimPeriod, EntityReference _erChannelPartner, EntityReference _erJob)
        {
            try
            {
                if (!ClaimLineExists(_service, _erJob.Id, _erChannelPartner.Id, _erClaimPeriod.Id, _claimCatwgory))
                {
                    Entity entObj = new Entity("hil_claimline");
                    Guid _claimLineId = Guid.Empty;

                    entObj["hil_jobid"] = _erJob;
                    entObj["hil_franchisee"] = _erChannelPartner;
                    entObj["hil_claimperiod"] = _erClaimPeriod;
                    entObj["hil_claimcategory"] = new EntityReference("hil_claimcategory", new Guid(_claimCatwgory));
                    entObj["hil_claimheader"] = new EntityReference("hil_claimheader", _performInvoiceId);
                    entObj["hil_claimamount"] = new Money(_claimAmount);
                    _claimLineId = _service.Create(entObj);
                    Entity ent = _service.Retrieve(Account.EntityLogicalName, _erChannelPartner.Id, new ColumnSet("ownerid"));
                    if (ent != null)
                    {
                        AssignRecord("hil_claimline", ent.GetAttributeValue<EntityReference>("ownerid").Id, _claimLineId, _service);
                    }
                }
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
        }
        private static void UpsertClaimLine(IOrganizationService _service, Guid _performInvoiceId, decimal _claimAmount, string _claimCatwgory, EntityReference _erClaimPeriod, EntityReference _erChannelPartner, EntityReference _erJob)
        {
            Guid _claimLineId = Guid.Empty;
            try
            {
                _claimLineId = ClaimLineExistsReturnGuId(_service, _erJob.Id, _erChannelPartner.Id, _erClaimPeriod.Id, _claimCatwgory);
                if (_claimLineId != Guid.Empty)
                {
                    _service.Delete("hil_claimline", _claimLineId);
                }
                Entity entObj = new Entity("hil_claimline");
                entObj["hil_jobid"] = _erJob;
                entObj["hil_franchisee"] = _erChannelPartner;
                entObj["hil_claimperiod"] = _erClaimPeriod;
                entObj["hil_claimcategory"] = new EntityReference("hil_claimcategory", new Guid(_claimCatwgory));
                entObj["hil_claimheader"] = new EntityReference("hil_claimheader", _performInvoiceId);
                entObj["hil_claimamount"] = new Money(_claimAmount);
                _claimLineId = _service.Create(entObj);
                Entity ent = _service.Retrieve(Account.EntityLogicalName, _erChannelPartner.Id, new ColumnSet("ownerid"));
                if (ent != null)
                {
                    AssignRecord("hil_claimline", ent.GetAttributeValue<EntityReference>("ownerid").Id, _claimLineId, _service);
                }
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
        }
        private static bool ClaimLineExists(IOrganizationService _service, Guid? _jobId, Guid _franchiseId, Guid _fiscalMonthId, string _claimCategory)
        {
            bool _returnVal = false;
            QueryExpression queryExp;
            EntityCollection entcoll;

            try
            {
                queryExp = new QueryExpression("hil_claimline");
                queryExp.ColumnSet = new ColumnSet("hil_claimperiod");
                queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                queryExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                queryExp.Criteria.AddCondition("hil_claimcategory", ConditionOperator.Equal, new Guid(_claimCategory));
                queryExp.Criteria.AddCondition("hil_franchisee", ConditionOperator.Equal, _franchiseId);
                queryExp.Criteria.AddCondition("hil_claimperiod", ConditionOperator.Equal, _fiscalMonthId);
                if (_jobId != null)
                {
                    queryExp.Criteria.AddCondition("hil_jobid", ConditionOperator.Equal, _jobId);
                }
                entcoll = _service.RetrieveMultiple(queryExp);
                if (entcoll.Entities.Count > 0)
                {
                    _returnVal = true;
                }
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
            return _returnVal;
        }
        private static Guid ClaimLineExistsReturnGuId(IOrganizationService _service, Guid? _jobId, Guid _franchiseId, Guid _fiscalMonthId, string _claimCategory)
        {
            Guid _returnVal = Guid.Empty;
            QueryExpression queryExp;
            EntityCollection entcoll;

            try
            {
                queryExp = new QueryExpression("hil_claimline");
                queryExp.ColumnSet = new ColumnSet("hil_claimperiod");
                queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                queryExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                queryExp.Criteria.AddCondition("hil_claimcategory", ConditionOperator.Equal, new Guid(_claimCategory));
                queryExp.Criteria.AddCondition("hil_franchisee", ConditionOperator.Equal, _franchiseId);
                queryExp.Criteria.AddCondition("hil_claimperiod", ConditionOperator.Equal, _fiscalMonthId);
                if (_jobId != null)
                {
                    queryExp.Criteria.AddCondition("hil_jobid", ConditionOperator.Equal, _jobId);
                }
                entcoll = _service.RetrieveMultiple(queryExp);
                if (entcoll.Entities.Count > 0)
                {
                    _returnVal = entcoll.Entities[0].Id;
                }
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
            return _returnVal;
        }
        public static void GeneratePerformaInvoiceBackup(IOrganizationService _service)
        {
            Guid _performaInvoice = new Guid("7208b2ca-b9f5-ea11-a815-000d3af0563c");
            QueryExpression queryExp;
            EntityCollection entcoll;
            QueryExpression queryExpTAT;
            EntityCollection entcollTAT;
            Int32 _havellsClosedJobs = 0;
            Int32 _wpClosedJobs = 0;
            Int32 _havellsJobsWithinTAT = 0;
            Int32 _wpJobsWithinTAT = 0;
            decimal _havellsJobsTATAchvPer = 0;
            decimal _wpJobsTATAchvPer = 0;
            Guid _12HourTAT = new Guid("4FA75DC1-4FD2-EA11-A813-000D3AF05D7B");
            Guid _24HourTAT = new Guid("BFEB9826-3C0F-EA11-A811-000D3AF0563C");

            Guid _guidAMC = new Guid("55A71A52-3C0B-E911-A94E-000D3AF06CD4");
            Guid _guidPMS = new Guid("E2129D79-3C0B-E911-A94E-000D3AF06CD4");

            Guid _tempTATCatg = Guid.Empty;
            Guid _tempTATAchvSlab = Guid.Empty;
            OptionSetValue opSourceOfJob = null;
            EntityReference erCallSubType = null;
            int i = 0;
            int j = 0;
            Entity entObj = _service.Retrieve("hil_claimheader", _performaInvoice, new ColumnSet("hil_fiscalmonth", "hil_franchisee"));
            if (entObj != null)
            {
                if (entObj.Attributes.Contains("hil_fiscalmonth"))
                {
                    Entity entFiscalMonth = _service.Retrieve("hil_claimperiod", entObj.GetAttributeValue<EntityReference>("hil_fiscalmonth").Id, new ColumnSet("hil_fromdate", "hil_todate"));
                    if (entFiscalMonth != null)
                    {
                        DateTime _fromDate = entFiscalMonth.GetAttributeValue<DateTime>("hil_fromdate").AddMinutes(330);
                        DateTime _toDate = entFiscalMonth.GetAttributeValue<DateTime>("hil_todate").AddMinutes(330);
                        queryExp = new QueryExpression(msdyn_workorder.EntityLogicalName);
                        queryExp.ColumnSet = new ColumnSet("hil_claimstatus", "msdyn_workorderid", "hil_tatachievementslab");
                        queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                        queryExp.Criteria.AddCondition("hil_claimstatus", ConditionOperator.NotIn, new object[] { 4, 3, 8, 7 });
                        queryExp.Criteria.AddCondition("msdyn_substatus", ConditionOperator.Equal, new Guid("1727FA6C-FA0F-E911-A94E-000D3AF060A1"));
                        queryExp.Criteria.AddCondition("hil_owneraccount", ConditionOperator.Equal, entObj.GetAttributeValue<EntityReference>("hil_franchisee").Id);
                        queryExp.Criteria.AddCondition("hil_isocr", ConditionOperator.NotEqual, true);
                        queryExp.Criteria.AddCondition(new ConditionExpression("msdyn_timeclosed", ConditionOperator.OnOrAfter, _fromDate)); //
                        queryExp.Criteria.AddCondition(new ConditionExpression("msdyn_timeclosed", ConditionOperator.OnOrBefore, _toDate)); //
                        entcoll = _service.RetrieveMultiple(queryExp);
                        if (entcoll.Entities.Count > 0)
                        {
                            Console.WriteLine("Error!!!");
                        }
                        else
                        {
                            queryExp = new QueryExpression(msdyn_workorder.EntityLogicalName);
                            queryExp.ColumnSet = new ColumnSet("hil_sourceofjob", "msdyn_name", "hil_tatcategory", "hil_owneraccount", "hil_fiscalmonth", "hil_laborinwarranty", "hil_brand", "hil_callsubtype", "hil_sparepartuse", "msdyn_timeclosed", "hil_channelpartnercategory", "hil_productsubcategory", "msdyn_workorderid", "hil_tatachievementslab");
                            queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                            queryExp.Criteria.AddCondition("hil_owneraccount", ConditionOperator.Equal, entObj.GetAttributeValue<EntityReference>("hil_franchisee").Id);
                            queryExp.Criteria.AddCondition("hil_isocr", ConditionOperator.NotEqual, true);
                            queryExp.Criteria.AddCondition("hil_claimstatus", ConditionOperator.Equal, 4);
                            queryExp.Criteria.AddCondition("msdyn_substatus", ConditionOperator.Equal, new Guid("1727FA6C-FA0F-E911-A94E-000D3AF060A1"));
                            queryExp.Criteria.AddCondition(new ConditionExpression("msdyn_timeclosed", ConditionOperator.OnOrAfter, _fromDate)); //
                            queryExp.Criteria.AddCondition(new ConditionExpression("msdyn_timeclosed", ConditionOperator.OnOrBefore, _toDate)); //
                            entcoll = _service.RetrieveMultiple(queryExp);

                            if (entcoll.Entities.Count > 0)
                            {
                                foreach (Entity entTemp in entcoll.Entities)
                                {
                                    Console.WriteLine(i++.ToString() + "/" + entcoll.Entities.Count.ToString());
                                    if (entTemp.Attributes.Contains("hil_callsubtype"))
                                    {
                                        erCallSubType = entTemp.GetAttributeValue<EntityReference>("hil_callsubtype");
                                    }
                                    else { opSourceOfJob = null; }
                                    if (entTemp.Attributes.Contains("hil_sourceofjob"))
                                    {
                                        opSourceOfJob = entTemp.GetAttributeValue<OptionSetValue>("hil_sourceofjob");
                                    }
                                    if ((erCallSubType.Id == _guidAMC) || (erCallSubType.Id == _guidPMS && opSourceOfJob.Value == 10))
                                    {
                                        continue; // Don't consider in slab calculation 
                                    }
                                    if (entTemp.GetAttributeValue<OptionSetValue>("hil_brand").Value == 1)
                                    {
                                        _havellsClosedJobs += 1;
                                        _tempTATCatg = entTemp.GetAttributeValue<EntityReference>("hil_tatcategory").Id;
                                        if (_tempTATCatg == _12HourTAT || _tempTATCatg == _24HourTAT)
                                        {
                                            _havellsJobsWithinTAT += 1;
                                        }
                                    }
                                    else if (entTemp.GetAttributeValue<OptionSetValue>("hil_brand").Value == 3)
                                    {
                                        _wpClosedJobs += 1;
                                        _tempTATCatg = entTemp.GetAttributeValue<EntityReference>("hil_tatcategory").Id;
                                        if (_tempTATCatg == _12HourTAT || _tempTATCatg == _24HourTAT)
                                        {
                                            _wpJobsWithinTAT += 1;
                                        }
                                    }

                                }
                                if (_havellsJobsWithinTAT > 0)
                                {
                                    _havellsJobsTATAchvPer = Math.Round(Convert.ToDecimal(_havellsJobsWithinTAT * 1.0 / _havellsClosedJobs) * 100, 2);
                                }
                                if (_wpJobsWithinTAT > 0)
                                {
                                    _wpJobsTATAchvPer = Math.Round(Convert.ToDecimal(_wpJobsWithinTAT * 1.0 / _wpClosedJobs) * 100, 2);
                                }
                                EntityReference erHavellsTATAchvSlab = null;
                                EntityReference erWPTATAchvSlab = null;
                                if (_havellsJobsTATAchvPer > 0)
                                {
                                    queryExpTAT = new QueryExpression("hil_tatachievementslabmaster");
                                    queryExpTAT.ColumnSet = new ColumnSet("hil_tatachievementslabmasterid", "hil_name");
                                    queryExpTAT.Criteria = new FilterExpression(LogicalOperator.And);
                                    queryExpTAT.Criteria.AddCondition("hil_brand", ConditionOperator.Equal, 1); // Havells
                                    queryExpTAT.Criteria.AddCondition(new ConditionExpression("hil_rangefrom", ConditionOperator.LessEqual, _havellsJobsTATAchvPer)); //
                                    queryExpTAT.Criteria.AddCondition(new ConditionExpression("hil_rangeto", ConditionOperator.GreaterEqual, _havellsJobsTATAchvPer)); //
                                    entcollTAT = _service.RetrieveMultiple(queryExpTAT);
                                    if (entcollTAT.Entities.Count > 0)
                                    {
                                        erHavellsTATAchvSlab = entcollTAT.Entities[0].ToEntityReference();
                                    }
                                }
                                if (_wpJobsTATAchvPer > 0)
                                {
                                    queryExpTAT = new QueryExpression("hil_tatachievementslabmaster");
                                    queryExpTAT.ColumnSet = new ColumnSet("hil_tatachievementslabmasterid", "hil_name");
                                    queryExpTAT.Criteria = new FilterExpression(LogicalOperator.And);
                                    queryExpTAT.Criteria.AddCondition("hil_brand", ConditionOperator.Equal, 3); // WP
                                    queryExpTAT.Criteria.AddCondition(new ConditionExpression("hil_rangefrom", ConditionOperator.LessEqual, _wpJobsTATAchvPer)); //
                                    queryExpTAT.Criteria.AddCondition(new ConditionExpression("hil_rangeto", ConditionOperator.GreaterEqual, _wpJobsTATAchvPer)); //
                                    entcollTAT = _service.RetrieveMultiple(queryExpTAT);
                                    if (entcollTAT.Entities.Count > 0)
                                    {
                                        erWPTATAchvSlab = entcollTAT.Entities[0].ToEntityReference();
                                    }
                                }
                                i = 0;
                                foreach (Entity entTemp in entcoll.Entities)
                                {
                                    Console.WriteLine(i++.ToString() + "/" + entcoll.Entities.Count.ToString());
                                    if (entTemp.Attributes.Contains("hil_callsubtype"))
                                    {
                                        erCallSubType = entTemp.GetAttributeValue<EntityReference>("hil_callsubtype");
                                    }
                                    else { opSourceOfJob = null; }
                                    if (entTemp.Attributes.Contains("hil_sourceofjob"))
                                    {
                                        opSourceOfJob = entTemp.GetAttributeValue<OptionSetValue>("hil_sourceofjob");
                                    }

                                    if ((erCallSubType.Id == _guidAMC) || (erCallSubType.Id == _guidPMS && opSourceOfJob.Value == 10))
                                    {
                                        continue; // Don't create TAT incentive lines
                                    }

                                    bool _laborInWty = entTemp.GetAttributeValue<bool>("hil_laborinwarranty");
                                    OptionSetValue opBrand = entTemp.GetAttributeValue<OptionSetValue>("hil_brand");
                                    bool _sparepartuse = entTemp.GetAttributeValue<bool>("hil_sparepartuse");
                                    EntityReference erProdSubCatg = entTemp.GetAttributeValue<EntityReference>("hil_productsubcategory");
                                    EntityCollection _entColl = null;
                                    OptionSetValue _opRelatedFranchiseCategory = entTemp.GetAttributeValue<OptionSetValue>("hil_channelpartnercategory");
                                    //erCallSubType = entTemp.GetAttributeValue<EntityReference>("hil_callsubtype");
                                    DateTime _closedOn = entTemp.GetAttributeValue<DateTime>("msdyn_timeclosed").AddMinutes(330);
                                    EntityReference _erClaimPeriod = entTemp.GetAttributeValue<EntityReference>("hil_fiscalmonth");
                                    EntityReference _erRelatedFranchise = entTemp.GetAttributeValue<EntityReference>("hil_owneraccount");

                                    _tempTATCatg = entTemp.GetAttributeValue<EntityReference>("hil_tatcategory").Id;
                                    if ((_tempTATCatg == _12HourTAT || _tempTATCatg == _24HourTAT) && _laborInWty)
                                    {
                                        if (entTemp.GetAttributeValue<OptionSetValue>("hil_brand").Value == 1 && erHavellsTATAchvSlab != null)
                                        {
                                            Entity entJobUpdate = new Entity(msdyn_workorder.EntityLogicalName, entTemp.Id);
                                            entJobUpdate["hil_tatachievementslab"] = erHavellsTATAchvSlab;
                                            _service.Update(entJobUpdate);
                                            _entColl = RetrieveTATIncentiveNonLLOYD(_service, _sparepartuse, opBrand, _opRelatedFranchiseCategory, erCallSubType, erProdSubCatg, erHavellsTATAchvSlab, _closedOn);
                                            if (_entColl.Entities.Count > 0)
                                            {
                                                if (_entColl.Entities[0].GetAttributeValue<Money>("hil_amount").Value > 0)
                                                {
                                                    //CreateClaimLine(_service, (_entColl.Entities[0].GetAttributeValue<Money>("hil_amount").Value), ClaimCategoryConst._TATIncentive.ToString(), _erClaimPeriod, _erRelatedFranchise, new EntityReference(msdyn_workorder.EntityLogicalName, entTemp.Id));
                                                }
                                            }
                                        }
                                        else if (entTemp.GetAttributeValue<OptionSetValue>("hil_brand").Value == 3 && erWPTATAchvSlab != null)
                                        {
                                            Entity entJobUpdate = new Entity(msdyn_workorder.EntityLogicalName, entTemp.Id);
                                            entJobUpdate["hil_tatachievementslab"] = erWPTATAchvSlab;
                                            _service.Update(entJobUpdate);
                                            _entColl = RetrieveTATIncentiveNonLLOYD(_service, _sparepartuse, opBrand, _opRelatedFranchiseCategory, erCallSubType, erProdSubCatg, erWPTATAchvSlab, _closedOn);
                                            if (_entColl.Entities.Count > 0)
                                            {
                                                if (_entColl.Entities[0].GetAttributeValue<Money>("hil_amount").Value > 0)
                                                {
                                                    //CreateClaimLine(_service, (_entColl.Entities[0].GetAttributeValue<Money>("hil_amount").Value), ClaimCategoryConst._TATIncentive.ToString(), _erClaimPeriod, _erRelatedFranchise, new EntityReference(msdyn_workorder.EntityLogicalName, entTemp.Id));
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
        public static void GenerateTATClaimLines(IOrganizationService _service, Guid _performaInvoiceId, Guid _franchiseId, DateTime _fromDate, DateTime _toDate)
        {
            #region Variable Declaration
            QueryExpression queryExp;
            EntityCollection entcoll;
            QueryExpression queryExpTAT;
            EntityCollection entcollTAT;
            //Int32 _totalClosedJobs = 0;
            //Int32 _totalJobsWithinTAT = 0;
            Int32 _havellsClosedJobs = 0;
            Int32 _wpClosedJobs = 0;
            Int32 _lloydClosedJobs = 0;
            Int32 _havellsJobsWithinTAT = 0;
            Int32 _wpJobsWithinTAT = 0;
            Int32 _lloydJobsWithinTAT = 0;
            decimal _havellsJobsTATAchvPer = 0;
            decimal _wpJobsTATAchvPer = 0;
            decimal _lloydJobsTATAchvPer = 0;
            Guid _prodCategoryId_LLOYDAirCon = new Guid("D51EDD9D-16FA-E811-A94C-000D3AF0694E");
            Guid _callSubTypeId_Installation = new Guid("E3129D79-3C0B-E911-A94E-000D3AF06CD4");

            Guid _12HourTAT = new Guid("4FA75DC1-4FD2-EA11-A813-000D3AF05D7B");
            Guid _24HourTAT = new Guid("BFEB9826-3C0F-EA11-A811-000D3AF0563C");
            Guid _guidAMC = new Guid("55A71A52-3C0B-E911-A94E-000D3AF06CD4");
            Guid _guidPMS = new Guid("E2129D79-3C0B-E911-A94E-000D3AF06CD4");
            Guid _tempTATCatg = Guid.Empty;
            Guid _tempTATAchvSlab = Guid.Empty;
            EntityReference erProdCatg = null;

            OptionSetValue opSourceOfJob = null;
            EntityReference erCallSubType = null;
            bool _isGascharged = false;
            int i = 0;
            int j = 0;
            string[] _jobColumns = {"hil_isgascharged", "hil_sourceofjob", "msdyn_name", "hil_tatcategory", "hil_owneraccount", "hil_fiscalmonth", "hil_laborinwarranty", "hil_brand", "hil_callsubtype", "hil_sparepartuse", "msdyn_timeclosed", "hil_channelpartnercategory", "hil_productsubcategory", "hil_productcategory", "msdyn_workorderid", "hil_tatachievementslab" };
            #endregion

            try
            {
                queryExp = new QueryExpression(msdyn_workorder.EntityLogicalName);
                queryExp.ColumnSet = new ColumnSet(_jobColumns);
                queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                queryExp.Criteria.AddCondition("hil_owneraccount", ConditionOperator.Equal, _franchiseId);
                queryExp.Criteria.AddCondition("hil_tatcategory", ConditionOperator.NotNull);
                queryExp.Criteria.AddCondition("hil_isocr", ConditionOperator.NotEqual, true);
                queryExp.Criteria.AddCondition("hil_claimstatus", ConditionOperator.Equal, 4);
                queryExp.Criteria.AddCondition("msdyn_substatus", ConditionOperator.Equal, new Guid("1727FA6C-FA0F-E911-A94E-000D3AF060A1"));
                queryExp.Criteria.AddCondition(new ConditionExpression("msdyn_timeclosed", ConditionOperator.OnOrAfter, _fromDate)); //
                queryExp.Criteria.AddCondition(new ConditionExpression("msdyn_timeclosed", ConditionOperator.OnOrBefore, _toDate)); //
                entcoll = _service.RetrieveMultiple(queryExp);

                if (entcoll.Entities.Count > 0)
                {
                    Console.WriteLine("Calculating Count");
                    foreach (Entity entTemp in entcoll.Entities)
                    {
                        _isGascharged = false;
                        if (entTemp.Attributes.Contains("hil_productcategory"))
                        {
                            erProdCatg = entTemp.GetAttributeValue<EntityReference>("hil_productcategory");
                        }
                        if (entTemp.Attributes.Contains("hil_isgascharged"))
                        {
                            _isGascharged = entTemp.GetAttributeValue<bool>("hil_isgascharged");
                        }
                        
                        if (entTemp.Attributes.Contains("hil_callsubtype"))
                        {
                            erCallSubType = entTemp.GetAttributeValue<EntityReference>("hil_callsubtype");
                        }
                        
                        if (entTemp.Attributes.Contains("hil_sourceofjob"))
                        {
                            opSourceOfJob = entTemp.GetAttributeValue<OptionSetValue>("hil_sourceofjob");
                        }
                        else { opSourceOfJob = null; }
                        if ((erCallSubType.Id == _guidAMC) || (erCallSubType.Id == _guidPMS && opSourceOfJob.Value == 10))
                        {
                            continue;
                        }
                        
                        if (entTemp.GetAttributeValue<OptionSetValue>("hil_brand").Value == 1) // Havells
                        {
                            _havellsClosedJobs += 1;
                            _tempTATCatg = entTemp.GetAttributeValue<EntityReference>("hil_tatcategory").Id;
                            if (_tempTATCatg == _12HourTAT || _tempTATCatg == _24HourTAT)
                            {
                                _havellsJobsWithinTAT += 1;
                            }
                        }
                        else if (entTemp.GetAttributeValue<OptionSetValue>("hil_brand").Value == 3) // Water Purifier 
                        {
                            _wpClosedJobs += 1;
                            _tempTATCatg = entTemp.GetAttributeValue<EntityReference>("hil_tatcategory").Id;
                            if (_tempTATCatg == _12HourTAT || _tempTATCatg == _24HourTAT)
                            {
                                _wpJobsWithinTAT += 1;
                            }
                        }
                        else if (entTemp.GetAttributeValue<OptionSetValue>("hil_brand").Value == 2) // Lloyd 
                        {
                            _lloydClosedJobs += 1;
                            _tempTATCatg = entTemp.GetAttributeValue<EntityReference>("hil_tatcategory").Id;
                            if (_tempTATCatg == _12HourTAT || _tempTATCatg == _24HourTAT)
                            {
                                _lloydJobsWithinTAT += 1;
                            }
                        }
                    }


                    if (_havellsJobsWithinTAT > 0)
                    {
                        _havellsJobsTATAchvPer = Math.Round(Convert.ToDecimal(_havellsJobsWithinTAT * 1.0 / _havellsClosedJobs) * 100, 2);
                    }
                    if (_wpJobsWithinTAT > 0)
                    {
                        _wpJobsTATAchvPer = Math.Round(Convert.ToDecimal(_wpJobsWithinTAT * 1.0 / _wpClosedJobs) * 100, 2);
                    }
                    if (_lloydJobsWithinTAT > 0)
                    {
                        _lloydJobsTATAchvPer = Math.Round(Convert.ToDecimal(_lloydJobsWithinTAT * 1.0 / _lloydClosedJobs) * 100, 2);
                    }
                    EntityReference erHavellsTATAchvSlab = null;
                    EntityReference erWPTATAchvSlab = null;
                    EntityReference erLloydTATAchvSlab = null;
                    if (_havellsJobsTATAchvPer > 0)
                    {
                        queryExpTAT = new QueryExpression("hil_tatachievementslabmaster");
                        queryExpTAT.ColumnSet = new ColumnSet("hil_tatachievementslabmasterid", "hil_name");
                        queryExpTAT.Criteria = new FilterExpression(LogicalOperator.And);
                        queryExpTAT.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                        queryExpTAT.Criteria.AddCondition("hil_brand", ConditionOperator.Equal, 1); // Havells
                        queryExpTAT.Criteria.AddCondition(new ConditionExpression("hil_rangefrom", ConditionOperator.LessEqual, _havellsJobsTATAchvPer)); //
                        queryExpTAT.Criteria.AddCondition(new ConditionExpression("hil_rangeto", ConditionOperator.GreaterEqual, _havellsJobsTATAchvPer)); //
                        entcollTAT = _service.RetrieveMultiple(queryExpTAT);
                        if (entcollTAT.Entities.Count > 0)
                        {
                            erHavellsTATAchvSlab = entcollTAT.Entities[0].ToEntityReference();
                        }
                    }
                    if (_wpJobsTATAchvPer > 0)
                    {
                        queryExpTAT = new QueryExpression("hil_tatachievementslabmaster");
                        queryExpTAT.ColumnSet = new ColumnSet("hil_tatachievementslabmasterid", "hil_name");
                        queryExpTAT.Criteria = new FilterExpression(LogicalOperator.And);
                        queryExpTAT.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                        queryExpTAT.Criteria.AddCondition("hil_brand", ConditionOperator.Equal, 3); // WP
                        queryExpTAT.Criteria.AddCondition(new ConditionExpression("hil_rangefrom", ConditionOperator.LessEqual, _wpJobsTATAchvPer)); //
                        queryExpTAT.Criteria.AddCondition(new ConditionExpression("hil_rangeto", ConditionOperator.GreaterEqual, _wpJobsTATAchvPer)); //
                        entcollTAT = _service.RetrieveMultiple(queryExpTAT);
                        if (entcollTAT.Entities.Count > 0)
                        {
                            erWPTATAchvSlab = entcollTAT.Entities[0].ToEntityReference();
                        }
                    }
                    if (_lloydJobsTATAchvPer > 0)
                    {
                        queryExpTAT = new QueryExpression("hil_tatachievementslabmaster");
                        queryExpTAT.ColumnSet = new ColumnSet("hil_tatachievementslabmasterid", "hil_name");
                        queryExpTAT.Criteria = new FilterExpression(LogicalOperator.And);
                        queryExpTAT.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                        queryExpTAT.Criteria.AddCondition("hil_brand", ConditionOperator.Equal, 2); // LLOYD
                        queryExpTAT.Criteria.AddCondition(new ConditionExpression("hil_rangefrom", ConditionOperator.LessEqual, _lloydJobsTATAchvPer)); //
                        queryExpTAT.Criteria.AddCondition(new ConditionExpression("hil_rangeto", ConditionOperator.GreaterEqual, _lloydJobsTATAchvPer)); //
                        entcollTAT = _service.RetrieveMultiple(queryExpTAT);
                        if (entcollTAT.Entities.Count > 0)
                        {
                            erLloydTATAchvSlab = entcollTAT.Entities[0].ToEntityReference();
                        }
                    }
                    Console.WriteLine("Calculating TAT Slab");
                    i = 0;
                    foreach (Entity entTemp in entcoll.Entities)
                    {
                        if (entTemp.Attributes.Contains("hil_productcategory"))
                        {
                            erProdCatg = entTemp.GetAttributeValue<EntityReference>("hil_productcategory");
                        }
                        if (entTemp.Attributes.Contains("hil_isgascharged"))
                        {
                            _isGascharged = entTemp.GetAttributeValue<bool>("hil_isgascharged");
                        }
                        if (entTemp.Attributes.Contains("hil_callsubtype"))
                        {
                            erCallSubType = entTemp.GetAttributeValue<EntityReference>("hil_callsubtype");
                        }
                        else { opSourceOfJob = null; }
                        if (entTemp.Attributes.Contains("hil_sourceofjob"))
                        {
                            opSourceOfJob = entTemp.GetAttributeValue<OptionSetValue>("hil_sourceofjob");
                        }
                        if ((erCallSubType.Id == _guidAMC) || (erCallSubType.Id == _guidPMS && opSourceOfJob.Value == 10))
                        {
                            continue;
                        }
                        if (entTemp.GetAttributeValue<OptionSetValue>("hil_brand").Value == 2 && ((_prodCategoryId_LLOYDAirCon == erProdCatg.Id && _callSubTypeId_Installation == erCallSubType.Id) || _isGascharged))
                        {
                            continue;
                        }
                        bool _laborInWty = entTemp.GetAttributeValue<bool>("hil_laborinwarranty");
                        OptionSetValue opBrand = entTemp.GetAttributeValue<OptionSetValue>("hil_brand");
                        bool _sparepartuse = entTemp.GetAttributeValue<bool>("hil_sparepartuse");
                        EntityReference erProdSubCatg = entTemp.GetAttributeValue<EntityReference>("hil_productsubcategory");
                        EntityCollection _entColl = null;
                        OptionSetValue _opRelatedFranchiseCategory = null;

                        if (entTemp.Attributes.Contains("hil_channelpartnercategory"))
                        {
                            _opRelatedFranchiseCategory = entTemp.GetAttributeValue<OptionSetValue>("hil_channelpartnercategory");
                        }
                        else
                        {
                            Entity entCust = _service.Retrieve(Account.EntityLogicalName, _franchiseId, new ColumnSet("hil_category"));
                            if (entCust != null)
                            {
                                _opRelatedFranchiseCategory = entCust.GetAttributeValue<OptionSetValue>("hil_category");
                                if (_opRelatedFranchiseCategory == null) {
                                    _opRelatedFranchiseCategory = new OptionSetValue(910590000); //Default Category "A"
                                }
                            }
                            else {
                                _opRelatedFranchiseCategory = new OptionSetValue(910590000); //Default Category "A"
                            }
                        }
                        DateTime _closedOn = entTemp.GetAttributeValue<DateTime>("msdyn_timeclosed").AddMinutes(330);
                        EntityReference _erClaimPeriod = entTemp.GetAttributeValue<EntityReference>("hil_fiscalmonth");
                        EntityReference _erRelatedFranchise = entTemp.GetAttributeValue<EntityReference>("hil_owneraccount");

                        _tempTATCatg = entTemp.GetAttributeValue<EntityReference>("hil_tatcategory").Id;
                        if ((_tempTATCatg == _12HourTAT || _tempTATCatg == _24HourTAT) && _laborInWty)
                        {
                            if (entTemp.GetAttributeValue<OptionSetValue>("hil_brand").Value == 1 && erHavellsTATAchvSlab != null)
                            {
                                Entity entJobUpdate = new Entity(msdyn_workorder.EntityLogicalName, entTemp.Id);
                                entJobUpdate["hil_tatachievementslab"] = erHavellsTATAchvSlab;
                                _service.Update(entJobUpdate);
                                _entColl = RetrieveTATIncentiveNonLLOYD(_service, _sparepartuse, opBrand, _opRelatedFranchiseCategory, erCallSubType, erProdSubCatg, erHavellsTATAchvSlab, _closedOn);
                                if (_entColl.Entities.Count > 0)
                                {
                                    if (_entColl.Entities[0].GetAttributeValue<Money>("hil_amount").Value > 0)
                                    {
                                        UpsertClaimLine(_service, _performaInvoiceId, (_entColl.Entities[0].GetAttributeValue<Money>("hil_amount").Value), ClaimCategoryConst._TATIncentive.ToString(), _erClaimPeriod, _erRelatedFranchise, new EntityReference(msdyn_workorder.EntityLogicalName, entTemp.Id));
                                    }
                                }
                            }
                            else if (entTemp.GetAttributeValue<OptionSetValue>("hil_brand").Value == 3 && erWPTATAchvSlab != null)
                            {
                                Entity entJobUpdate = new Entity(msdyn_workorder.EntityLogicalName, entTemp.Id);
                                entJobUpdate["hil_tatachievementslab"] = erWPTATAchvSlab;
                                _service.Update(entJobUpdate);
                                _entColl = RetrieveTATIncentiveNonLLOYD(_service, _sparepartuse, opBrand, _opRelatedFranchiseCategory, erCallSubType, erProdSubCatg, erWPTATAchvSlab, _closedOn);
                                if (_entColl.Entities.Count > 0)
                                {
                                    if (_entColl.Entities[0].GetAttributeValue<Money>("hil_amount").Value > 0)
                                    {
                                        CreateClaimLine(_service, _performaInvoiceId, (_entColl.Entities[0].GetAttributeValue<Money>("hil_amount").Value), ClaimCategoryConst._TATIncentive.ToString(), _erClaimPeriod, _erRelatedFranchise, new EntityReference(msdyn_workorder.EntityLogicalName, entTemp.Id));
                                    }
                                }
                            }
                            if (entTemp.GetAttributeValue<OptionSetValue>("hil_brand").Value == 2 && erLloydTATAchvSlab != null)
                            {
                                Entity entJobUpdate = new Entity(msdyn_workorder.EntityLogicalName, entTemp.Id);
                                entJobUpdate["hil_tatachievementslab"] = erLloydTATAchvSlab;
                                _service.Update(entJobUpdate);
                                _entColl = RetrieveTATIncentiveNonLLOYD(_service, _sparepartuse, opBrand, _opRelatedFranchiseCategory, erCallSubType, erProdSubCatg, erLloydTATAchvSlab, _closedOn);
                                if (_entColl.Entities.Count > 0)
                                {
                                    if (_entColl.Entities[0].GetAttributeValue<Money>("hil_amount").Value > 0)
                                    {
                                        UpsertClaimLine(_service, _performaInvoiceId, (_entColl.Entities[0].GetAttributeValue<Money>("hil_amount").Value), ClaimCategoryConst._TATIncentive.ToString(), _erClaimPeriod, _erRelatedFranchise, new EntityReference(msdyn_workorder.EntityLogicalName, entTemp.Id));
                                    }
                                }
                            }
                        }
                        Console.WriteLine("Calculating TAT Incentive " + i++.ToString() + "/" + entcoll.Entities.Count.ToString());
                    }
                }
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
        }
        private static Guid GeneratePerformaInvoice(IOrganizationService _service, Guid _franchiseeAcId, Guid _fiscalMonthId)
        {
            QueryExpression queryExp;
            EntityCollection entcoll;
            Guid _performaInvoiceId = Guid.Empty;
            try
            {
                queryExp = new QueryExpression("hil_claimheader");
                queryExp.ColumnSet = new ColumnSet("hil_claimheaderid");
                queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                queryExp.Criteria.AddCondition("hil_fiscalmonth", ConditionOperator.Equal, _fiscalMonthId);
                queryExp.Criteria.AddCondition("hil_franchisee", ConditionOperator.Equal, _franchiseeAcId);
                entcoll = _service.RetrieveMultiple(queryExp);

                if (entcoll.Entities.Count == 0)
                {
                    Entity entObj = new Entity("hil_claimheader");
                    entObj["hil_fiscalmonth"] = new EntityReference("hil_claimperiod", _fiscalMonthId);
                    entObj["hil_franchisee"] = new EntityReference("account", _franchiseeAcId);
                    _performaInvoiceId = _service.Create(entObj);

                    Entity ent = _service.Retrieve(Account.EntityLogicalName, _franchiseeAcId, new ColumnSet("ownerid"));
                    if (ent != null)
                    {
                        AssignRecord("hil_claimheader", ent.GetAttributeValue<EntityReference>("ownerid").Id, _performaInvoiceId, _service);
                    }
                }
                else
                {
                    _performaInvoiceId = entcoll.Entities[0].Id;
                }
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
            return _performaInvoiceId;
        }
        private static void AssignRecord(String TargetLogicalName, Guid AssigneeId, Guid TargetId, IOrganizationService service)
        {
            try
            {
                AssignRequest assign = new AssignRequest();
                assign.Assignee = new EntityReference(SystemUser.EntityLogicalName, AssigneeId); //User or team
                assign.Target = new EntityReference(TargetLogicalName, TargetId); //Record to be assigned
                service.Execute(assign);
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
        }
        public static OptionSetValue RefreshClaimParameters(IOrganizationService _service, Guid _fiscalMonthId, Entity ent)
        {
            #region Varialbe Declaration
            EntityReference _unitWarranty = null;
            Guid _warrantyTemplateId = Guid.Empty;
            bool SparePartUsed = false;
            QueryExpression qrExp;
            bool laborInWarranty = false;
            int _jobWarrantyStatus = 2; //OutWarranty
            int _jobWarrantySubStatus = 0;
            int _warrantyTempType = 0;
            DateTime _unitWarrStartDate = new DateTime(1900, 1, 1);
            double _jobMonth = 0;
            EntityReference _claimPeriod = null;
            bool _isOCR = false;
            bool _reviewForCountryClassification = false;
            EntityReference erCustomerAsset = null;
            OptionSetValue _claimStatusVal = null;
            EntityReference erDistJob = null;
            EntityReference erDistASP = null;
            OptionSetValue _brand = null;
            Guid _jobGuId = ent.Id;
            int upCountry = 1;
            #endregion

            if (ent != null)
            {
                if (ent.Attributes.Contains("hil_brand"))
                {
                    _brand = ent.GetAttributeValue<OptionSetValue>("hil_brand");
                }

                if (_brand.Value == 2) //LLOYD
                {
                    QueryExpression Query1;
                    EntityCollection entcoll1;
                    Query1 = new QueryExpression("hil_assignmentmatrix");
                    Query1.ColumnSet = new ColumnSet("hil_upcountry");
                    Query1.Criteria = new FilterExpression(LogicalOperator.And);
                    Query1.Criteria.AddCondition("hil_division", ConditionOperator.Equal, ent.GetAttributeValue<EntityReference>("hil_productcategory").Id);
                    Query1.Criteria.AddCondition("hil_pincode", ConditionOperator.Equal, ent.GetAttributeValue<EntityReference>("hil_pincode").Id);
                    Query1.Criteria.AddCondition("hil_franchiseedirectengineer", ConditionOperator.Equal, ent.GetAttributeValue<EntityReference>("hil_owneraccount").Id);
                    Query1.Criteria.AddCondition("hil_callsubtype", ConditionOperator.Equal, ent.GetAttributeValue<EntityReference>("hil_callsubtype").Id);
                    Query1.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                    entcoll1 = _service.RetrieveMultiple(Query1);
                    if (entcoll1.Entities.Count > 0)
                    {
                        if (entcoll1.Entities[0].Attributes.Contains("hil_upcountry"))
                        {
                            bool flag = entcoll1.Entities[0].GetAttributeValue<bool>("hil_upcountry");
                            upCountry = (flag == true ? 2 : 1);
                        }
                        else
                        {
                            upCountry = 1; // Local
                        }
                    }
                    else
                    {
                        upCountry = 1; // Local
                        _reviewForCountryClassification = true;
                        // Update "Review for Country Classification" flag om Job
                        //1) If Assignment Matrix does not exist then flag job as "BSH Review for upcountry" and move it in separate bucket.--- ON CLOSURE OF JOB
                        //2) If any job flagged as "BSH Review for upcountry" and Login User Position is BSH and NSH then show button "Mark as Upcountry" and once user clicks on Button change the coutry classification to "OTHER DISTRICT" and disable the Button.
                        //3) Bypass such Jobs where BSH Review for upcountry is True && Country classification is "OTHER DISTRICT" from Param Refresh
                    }
                }
                else
                {
                    if (ent.Attributes.Contains("hil_district"))
                    {
                        erDistJob = ent.GetAttributeValue<EntityReference>("hil_district");
                    }
                    if (ent.Attributes.Contains("DistASP.hil_district"))
                    {
                        erDistASP = (EntityReference)ent.GetAttributeValue<AliasedValue>("DistASP.hil_district").Value;
                    }

                    if (erDistJob != null && erDistASP != null)
                    {
                        if (erDistJob.Id != erDistASP.Id)
                        {
                            //New Logic for Havells need to be implement
                            //Check if Customer district exists in Channel Partner Local Districts Setup then Mark as Local else upcountry
                            QueryExpression Query1;
                            EntityCollection entcoll1;
                            Query1 = new QueryExpression("hil_channelpartnercountryclassification");
                            Query1.ColumnSet = new ColumnSet("hil_channelpartnercountryclassificationid");
                            Query1.Criteria = new FilterExpression(LogicalOperator.And);
                            Query1.Criteria.AddCondition("hil_brand", ConditionOperator.Equal, _brand.Value);
                            Query1.Criteria.AddCondition("hil_countryclassification", ConditionOperator.Equal, 1);
                            Query1.Criteria.AddCondition("hil_channelpartner", ConditionOperator.Equal, ent.GetAttributeValue<EntityReference>("hil_owneraccount").Id);
                            Query1.Criteria.AddCondition("hil_district", ConditionOperator.Equal, erDistJob.Id);
                            Query1.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                            entcoll1 = _service.RetrieveMultiple(Query1);
                            if (entcoll1.Entities.Count > 0) {
                                upCountry = 1; // Local
                            }
                            else 
                            {
                                upCountry = 2; // UpCountry
                            }
                        }
                    }
                }

                if (ent.Attributes.Contains("msdyn_customerasset"))
                {
                    erCustomerAsset = ent.GetAttributeValue<EntityReference>("msdyn_customerasset");
                }

                if (ent.Attributes.Contains("hil_isocr"))
                {
                    _isOCR = ent.GetAttributeValue<bool>("hil_isocr");
                }

                if (!_isOCR && erCustomerAsset != null)
                {
                    DateTime _jobClosedOn = ent.GetAttributeValue<DateTime>("msdyn_timeclosed").AddMinutes(330);
                    DateTime _jobCreatedOn = ent.GetAttributeValue<DateTime>("createdon").AddMinutes(330);
                    DateTime _assetPurchaseDate = new DateTime(1900, 1, 1);

                    Entity entCustAsset = _service.Retrieve("msdyn_customerasset", ent.GetAttributeValue<EntityReference>("msdyn_customerasset").Id, new ColumnSet("hil_invoicedate", "hil_invoiceavailable"));
                    if (entCustAsset != null)
                    {
                        if (entCustAsset.Attributes.Contains("hil_invoicedate"))
                        {
                            _assetPurchaseDate = entCustAsset.GetAttributeValue<DateTime>("hil_invoicedate");
                        }
                    }

                    if (_jobCreatedOn < _assetPurchaseDate)
                    {
                        _jobCreatedOn = _assetPurchaseDate;
                    }

                    qrExp = new QueryExpression("msdyn_workorderincident");
                    qrExp.ColumnSet = new ColumnSet("msdyn_workorderincidentid");
                    qrExp.Criteria = new FilterExpression(LogicalOperator.And);
                    qrExp.Criteria.AddCondition("msdyn_workorder", ConditionOperator.Equal, _jobGuId);
                    qrExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                    qrExp.Criteria.AddCondition("hil_warrantystatus", ConditionOperator.Equal, 3); // Warranty Void
                    EntityCollection entCol1 = _service.RetrieveMultiple(qrExp);
                    if (entCol1.Entities.Count == 0)
                    {
                        qrExp = new QueryExpression("hil_unitwarranty");
                        qrExp.ColumnSet = new ColumnSet("hil_name", "hil_warrantytemplate", "hil_warrantystartdate", "hil_warrantyenddate");
                        qrExp.Criteria = new FilterExpression(LogicalOperator.And);
                        qrExp.Criteria.AddCondition("hil_customerasset", ConditionOperator.Equal, ent.GetAttributeValue<EntityReference>("msdyn_customerasset").Id);
                        qrExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                        EntityCollection entCol2 = _service.RetrieveMultiple(qrExp);
                        if (entCol2.Entities.Count > 0)
                        {
                            foreach (Entity Wt in entCol2.Entities)
                            {
                                DateTime iValidTo = ((DateTime)Wt["hil_warrantyenddate"]).AddMinutes(330);
                                DateTime iValidFrom = ((DateTime)Wt["hil_warrantystartdate"]).AddMinutes(330);
                                if (_jobCreatedOn >= iValidFrom && _jobCreatedOn <= iValidTo)
                                {
                                    _jobWarrantyStatus = 1;
                                    _unitWarranty = Wt.ToEntityReference();
                                    _warrantyTemplateId = Wt.GetAttributeValue<EntityReference>("hil_warrantytemplate").Id;
                                    Entity _entTemp = _service.Retrieve(hil_warrantytemplate.EntityLogicalName, _warrantyTemplateId, new ColumnSet("hil_name", "hil_type"));
                                    if (_entTemp != null)
                                    {
                                        _warrantyTempType = _entTemp.GetAttributeValue<OptionSetValue>("hil_type").Value;
                                        if (_warrantyTempType == 1) { _jobWarrantySubStatus = 1; }
                                        else if (_warrantyTempType == 2) { _jobWarrantySubStatus = 2; }
                                        else if (_warrantyTempType == 7) { _jobWarrantySubStatus = 3; }
                                        else if (_warrantyTempType == 3) { _jobWarrantySubStatus = 4; }
                                    }
                                    _unitWarrStartDate = Wt.GetAttributeValue<DateTime>("hil_warrantystartdate").AddMinutes(330);
                                    TimeSpan difference = (_jobCreatedOn - _unitWarrStartDate);
                                    _jobMonth = Math.Ceiling((difference.Days * 1.0 / 30.42));
                                    qrExp = new QueryExpression("hil_labor");
                                    qrExp.ColumnSet = new ColumnSet("hil_laborid", "hil_includedinwarranty", "hil_validtomonths", "hil_validfrommonths");
                                    qrExp.Criteria = new FilterExpression(LogicalOperator.And);
                                    qrExp.Criteria.AddCondition("hil_warrantytemplateid", ConditionOperator.Equal, _warrantyTemplateId);
                                    qrExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                                    EntityCollection entCol3 = _service.RetrieveMultiple(qrExp);
                                    if (entCol3.Entities.Count == 0) { laborInWarranty = true; }
                                    else
                                    {
                                        if (_jobMonth >= entCol3.Entities[0].GetAttributeValue<int>("hil_validfrommonths") && _jobMonth <= entCol3.Entities[0].GetAttributeValue<int>("hil_validtomonths"))
                                        {
                                            OptionSetValue _laborType = entCol3.Entities[0].GetAttributeValue<OptionSetValue>("hil_includedinwarranty");
                                            laborInWarranty = _laborType.Value == 1 ? true : false;
                                        }
                                        else
                                        {
                                            OptionSetValue _laborType = entCol3.Entities[0].GetAttributeValue<OptionSetValue>("hil_includedinwarranty");
                                            laborInWarranty = !(_laborType.Value == 1 ? true : false);
                                        }
                                    }
                                    break;
                                }
                            }
                        }
                    }
                    else
                    {
                        _jobWarrantyStatus = 3;
                    }

                    qrExp = new QueryExpression("msdyn_workorderproduct");
                    qrExp.ColumnSet = new ColumnSet("msdyn_workorderproductid");
                    qrExp.Criteria = new FilterExpression(LogicalOperator.And);
                    qrExp.Criteria.AddCondition("msdyn_workorder", ConditionOperator.Equal, _jobGuId);
                    qrExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                    qrExp.Criteria.AddCondition("hil_markused", ConditionOperator.Equal, true);
                    EntityCollection entCol4 = _service.RetrieveMultiple(qrExp);
                    if (entCol4.Entities.Count > 0) { SparePartUsed = true; } else { SparePartUsed = false; }

                    if (ent.GetAttributeValue<EntityReference>("msdyn_substatus").Id == new Guid("6C8F2123-5106-EA11-A811-000D3AF057DD"))
                    {
                        qrExp = new QueryExpression("hil_sawactivity");
                        qrExp.ColumnSet = new ColumnSet("hil_sawactivityid", "hil_approvalstatus");
                        qrExp.Criteria = new FilterExpression(LogicalOperator.And);
                        qrExp.Criteria.AddCondition("hil_jobid", ConditionOperator.Equal, _jobGuId);
                        qrExp.Criteria.AddCondition("hil_sawcategory", ConditionOperator.Equal, new Guid("E123BD08-8ADD-EA11-A813-000D3AF05A4B"));
                        EntityCollection entCol50 = _service.RetrieveMultiple(qrExp);
                        if (entCol50.Entities.Count > 0)
                        {
                            OptionSetValue sawStatus = entCol50.Entities[0].GetAttributeValue<OptionSetValue>("hil_approvalstatus");
                            if (sawStatus.Value == 4) // Rejected
                            {
                                _claimStatusVal = new OptionSetValue(7); // Abnormal Penalty Imposed
                            }
                            else if (sawStatus.Value == 3) // Approved
                            {
                                _claimStatusVal = new OptionSetValue(8); // Abnormal Penalty Waived
                            }
                            else {
                                _claimStatusVal = new OptionSetValue(1); //Under Review
                            }
                        }
                        else {
                            _claimStatusVal = new OptionSetValue(1); //Under Review
                        }
                    }
                    else
                    {
                        qrExp = new QueryExpression("hil_sawactivity");
                        qrExp.ColumnSet = new ColumnSet("hil_sawactivityid");
                        qrExp.Criteria = new FilterExpression(LogicalOperator.And);
                        qrExp.Criteria.AddCondition("hil_jobid", ConditionOperator.Equal, _jobGuId);
                        EntityCollection entCol5 = _service.RetrieveMultiple(qrExp);
                        if (entCol5.Entities.Count == 0)
                        {
                            _claimStatusVal = new OptionSetValue(4); // Claim Approved
                        }
                        else
                        {
                            qrExp = new QueryExpression("hil_sawactivityapproval");
                            qrExp.ColumnSet = new ColumnSet("hil_sawactivityapprovalid");
                            qrExp.Criteria = new FilterExpression(LogicalOperator.And);
                            qrExp.Criteria.AddCondition("hil_jobid", ConditionOperator.Equal, _jobGuId);
                            qrExp.Criteria.AddCondition("hil_approvalstatus", ConditionOperator.Equal, 4);// Rejected
                            EntityCollection entCol6 = _service.RetrieveMultiple(qrExp);
                            if (entCol6.Entities.Count > 0)
                            {
                                _claimStatusVal = new OptionSetValue(3); //Rejected
                            }
                            else
                            {
                                qrExp = new QueryExpression("hil_sawactivityapproval");
                                qrExp.ColumnSet = new ColumnSet("hil_sawactivityapprovalid");
                                qrExp.Criteria = new FilterExpression(LogicalOperator.And);
                                qrExp.Criteria.AddCondition("hil_jobid", ConditionOperator.Equal, _jobGuId);
                                qrExp.Criteria.AddCondition("hil_approvalstatus", ConditionOperator.NotEqual, 3); //Approved
                                EntityCollection entCol7 = _service.RetrieveMultiple(qrExp);
                                if (entCol7.Entities.Count == 0)
                                {
                                    _claimStatusVal = new OptionSetValue(4); //Approved
                                }
                                else
                                {
                                    _claimStatusVal = new OptionSetValue(1); //Under Review
                                }
                            }
                        }
                    }

                    Entity entJobUpdate = new Entity(msdyn_workorder.EntityLogicalName, _jobGuId);

                    entJobUpdate["hil_warrantystatus"] = new OptionSetValue(_jobWarrantyStatus);
                    if (_jobWarrantyStatus == 1)
                    {
                        entJobUpdate["hil_warrantysubstatus"] = new OptionSetValue(_jobWarrantySubStatus);
                    }
                    else
                    {
                        entJobUpdate["hil_warrantysubstatus"] = null;
                    }
                    if (_unitWarranty != null)
                    {
                        entJobUpdate["hil_unitwarranty"] = _unitWarranty;
                    }
                    else { entJobUpdate["hil_unitwarranty"] = null; }

                    entJobUpdate["hil_laborinwarranty"] = laborInWarranty;

                    if (_assetPurchaseDate.Year != 1900 && _assetPurchaseDate.Year != 1)
                    {
                        entJobUpdate["hil_purchasedate"] = _assetPurchaseDate;
                    }
                    else
                    {
                        entJobUpdate["hil_purchasedate"] = null;
                    }

                    entJobUpdate["hil_sparepartuse"] = SparePartUsed;

                    if (_fiscalMonthId != Guid.Empty)
                    {
                        entJobUpdate["hil_fiscalmonth"] = new EntityReference("hil_claimperiod", _fiscalMonthId);
                    }

                    if (ent.Attributes.Contains("custCatg.hil_category"))
                    {
                        OptionSetValue opVal = (OptionSetValue)((AliasedValue)ent.Attributes["custCatg.hil_category"]).Value;
                        entJobUpdate["hil_channelpartnercategory"] = new OptionSetValue(opVal.Value);
                    }
                    entJobUpdate["hil_claimstatus"] = _claimStatusVal;

                    //if (_brand.Value != 2)
                    //{
                        entJobUpdate["hil_countryclassification"] = new OptionSetValue(upCountry);
                    //}
                    if (_reviewForCountryClassification) {
                        entJobUpdate["hil_reviewforcountryclassification"] = _reviewForCountryClassification;
                    }
                    try
                    {
                        _service.Update(entJobUpdate);
                    }
                    catch (Exception ex) { Console.WriteLine(ex.Message); }
                }
            }

            return _claimStatusVal;
        }
        public static void ProcessFranchiseClaims(IOrganizationService _service, Guid _franchiseIdRefresh, Guid _jobIdRefresh)
        {

            #region Variable declaration
            Guid _fiscalMonthId = new Guid("DC417E88-E62B-EB11-A813-0022486E5C9D");
            //Guid _fiscalMonthId = Guid.Empty;
            DateTime? _startDate = null;
            DateTime? _endDate = null;
            Guid _performaInvoiceId = Guid.Empty;
            string _fetchXML = string.Empty;
            QueryExpression queryExp;
            EntityCollection entcoll;
            EntityCollection entcollTATSlabJobs;
            EntityReference _erFiscalMonth = null;
            EntityReference _erFranchise = null;
            int _totalClosedJobs = 0, _fixedCharges = 0;
            string[] _jobColumns = { "hil_fiscalmonth", "msdyn_name", "createdon", "hil_owneraccount", "hil_channelpartnercategory", "hil_warrantystatus", "hil_countryclassification", "hil_tatachievementslab", "hil_tatcategory", "hil_schemecode", "hil_salesoffice", "hil_productcategory", "hil_kkgcode", "hil_typeofassignee", "hil_brand", "hil_productsubcategory", "hil_callsubtype", "hil_claimstatus", "hil_isocr", "hil_laborinwarranty", "hil_isgascharged", "hil_sparepartuse", "hil_jobclosemobile", "msdyn_timeclosed", "msdyn_customerasset", "hil_quantity", "hil_purchasedate" };
            string _performaRemarks = string.Empty;
            #endregion

            #region Fetching Fiscal Month Date Range
            if (_fiscalMonthId != Guid.Empty)
            {
                Entity ent = _service.Retrieve("hil_claimperiod", _fiscalMonthId, new ColumnSet("hil_fromdate", "hil_todate"));
                if (ent != null)
                {
                    _startDate = ent.GetAttributeValue<DateTime>("hil_fromdate").AddMinutes(330);
                    _endDate = ent.GetAttributeValue<DateTime>("hil_todate").AddMinutes(330);
                    _erFiscalMonth = ent.ToEntityReference();
                }
            }
            else
            {
                DateTime _processDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);

                queryExp = new QueryExpression("hil_claimperiod");
                queryExp.ColumnSet = new ColumnSet("hil_fromdate", "hil_todate");
                queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                queryExp.Criteria.AddCondition("hil_fromdate", ConditionOperator.OnOrBefore, _processDate);
                queryExp.Criteria.AddCondition("hil_todate", ConditionOperator.OnOrAfter, _processDate);
                entcoll = _service.RetrieveMultiple(queryExp);
                if (entcoll.Entities.Count > 0)
                {
                    _startDate = entcoll.Entities[0].GetAttributeValue<DateTime>("hil_fromdate").AddMinutes(330);
                    _endDate = entcoll.Entities[0].GetAttributeValue<DateTime>("hil_todate").AddMinutes(330);
                    _erFiscalMonth = entcoll.Entities[0].ToEntityReference();
                }
            }
            #endregion

            //queryExp = new QueryExpression("msdyn_workorder");
            //queryExp.ColumnSet = new ColumnSet("hil_owneraccount");
            //queryExp.Criteria = new FilterExpression(LogicalOperator.And);
            //queryExp.Criteria.AddCondition("hil_isocr", ConditionOperator.NotEqual, true);
            //queryExp.Criteria.AddCondition("msdyn_substatus", ConditionOperator.Equal, new Guid("1727FA6C-FA0F-E911-A94E-000D3AF060A1")); //{Sub-Status:"Closed"}
            //queryExp.Criteria.AddCondition("msdyn_timeclosed", ConditionOperator.OnOrAfter, _startDate);
            //queryExp.Criteria.AddCondition("msdyn_timeclosed", ConditionOperator.OnOrBefore, _endDate);
            //queryExp.Criteria.AddCondition("hil_owneraccount", ConditionOperator.Equal, _franchiseIdRefresh);
            //queryExp.Criteria.AddCondition("hil_typeofassignee", ConditionOperator.In, new object[] {new Guid("4A1AA189-1208-E911-A94D-000D3AF0694E"), new Guid("0197EA9B-1208-E911-A94D-000D3AF0694E") }); //{Position:"Franchise","Franchise Technician"}
            //queryExp.Distinct = true;

            _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
            <entity name='msdyn_workorder'>
            <attribute name='hil_owneraccount' />
            <filter type='and'>
                <condition attribute='hil_isocr' operator='ne' value='1' />
                <filter type='or'>
                    <condition attribute='msdyn_substatus' operator='eq' value='{1727FA6C-FA0F-E911-A94E-000D3AF060A1}' />
                    <condition attribute='hil_claimstatus' operator='eq' value='4' />
                </filter>
                <condition attribute='msdyn_timeclosed' operator='on-or-after' value='" + (_startDate.Value.Year.ToString() + "-" + _startDate.Value.Month.ToString().PadLeft(2, '0') + "-" + _startDate.Value.Day.ToString().PadLeft(2, '0')) + @"' />
                <condition attribute='msdyn_timeclosed' operator='on-or-before' value='" + (_endDate.Value.Year.ToString() + "-" + _endDate.Value.Month.ToString().PadLeft(2, '0') + "-" + _endDate.Value.Day.ToString().PadLeft(2, '0')) + @"' />
                <condition attribute='hil_owneraccount' operator='eq' value='{" + _franchiseIdRefresh + @"}' />
                <condition attribute='hil_typeofassignee' operator='in'>
                <value uiname='Franchise' uitype='position'>{4A1AA189-1208-E911-A94D-000D3AF0694E}</value>
                <value uiname='Franchise Technician' uitype='position'>{0197EA9B-1208-E911-A94D-000D3AF0694E}</value>
                </condition>
            </filter>
            </entity>
            </fetch>";
            entcoll = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
            if (entcoll.Entities.Count > 0)
            {
                foreach (Entity entFranchise in entcoll.Entities)
                {
                    try
                    {
                        _performaInvoiceId = Guid.Empty;
                        _totalClosedJobs = 0;
                        _fixedCharges = 0;
                        _erFranchise = entFranchise.GetAttributeValue<EntityReference>("hil_owneraccount");
                        #region Performa Invoice Generation
                        _performaInvoiceId = GeneratePerformaInvoice(_service, _erFranchise.Id, _erFiscalMonth.Id);
                        #endregion
                        if (_performaInvoiceId != null)
                        {
                            #region Calculating TAT Incentive Slab
                            Console.WriteLine("Calculating TAT Incentive Slab");
                            queryExp = new QueryExpression(msdyn_workorder.EntityLogicalName);
                            queryExp.ColumnSet = new ColumnSet("hil_claimstatus", "msdyn_workorderid", "hil_tatachievementslab");
                            queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                            queryExp.Criteria.AddCondition("hil_claimstatus", ConditionOperator.NotIn, new object[] { 4, 3, 8, 7 });
                            queryExp.Criteria.AddCondition("msdyn_substatus", ConditionOperator.Equal, new Guid("1727FA6C-FA0F-E911-A94E-000D3AF060A1"));
                            queryExp.Criteria.AddCondition("hil_owneraccount", ConditionOperator.Equal, _erFranchise.Id);
                            queryExp.Criteria.AddCondition("hil_isocr", ConditionOperator.NotEqual, true);
                            queryExp.Criteria.AddCondition(new ConditionExpression("msdyn_timeclosed", ConditionOperator.OnOrAfter, _startDate));
                            queryExp.Criteria.AddCondition(new ConditionExpression("msdyn_timeclosed", ConditionOperator.OnOrBefore, _endDate));
                            entcollTATSlabJobs = _service.RetrieveMultiple(queryExp);
                            if (entcollTATSlabJobs.Entities.Count == 0)
                            {
                                GenerateTATClaimLines(_service, _performaInvoiceId, _erFranchise.Id, Convert.ToDateTime(_startDate), Convert.ToDateTime(_endDate));
                            }
                            else
                            {
                                _performaRemarks = "Some Jobs Claim Status is Under Review !!! Unable to process TAT Slab Incentive";
                            }
                            #endregion
                            //13cd24d2-c1f4-ea11-a815-000d3af05d7b

                            #region Calculating Claim Amount
                            //Console.WriteLine("Calculating Claim amount");
                            //_fetchXML = @"<fetch distinct='false' mapping='logical' aggregate='true'>
                            //<entity name='hil_claimline'>
                            //<attribute name='hil_claimamount' alias='claimamount' aggregate='sum'/> 
                            //<attribute name='hil_jobid' alias='jobid' groupby='true' /> 
                            //<filter type='and'>
                            //    <condition attribute='hil_franchisee' operator='eq' value='{" + _erFranchise.Id + @"}' />
                            //    <condition attribute='hil_claimperiod' operator='eq' value='{" + _erFiscalMonth.Id + @"}' />
                            //    <condition attribute='hil_jobid' operator='eq' value='{2c10af69-10ec-ea11-a815-000d3af055b6}' />
                            //</filter>
                            //</entity>
                            //</fetch>";
                            //EntityCollection entcollCA = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                            //Decimal _JobClaimAmount = 0;
                            //foreach (Entity entCA in entcollCA.Entities)
                            //{
                            //_JobClaimAmount = ((Money)((AliasedValue)entCA["claimamount"]).Value).Value;
                            //_claimAmount += _JobClaimAmount;
                            //if (entCA.Attributes.Contains("jobid"))
                            //{
                            //    Guid _claimJobId = ((EntityReference)((AliasedValue)entCA["jobid"]).Value).Id;
                            //    if (_claimJobId != null)
                            //    {
                            //        entUpdate = new Entity("msdyn_workorder", _claimJobId);
                            //        entUpdate["hil_generateclaim"] = true;
                            //        entUpdate["hil_claimamount"] = new Money(_JobClaimAmount);
                            //        _service.Update(entUpdate);
                            //    }
                            //}
                            //}

                            #endregion

                            #region Update Perform Invoice information
                            //entUpdate = new Entity("hil_claimheader", _performaInvoiceId);
                            //entUpdate["hil_description"] = _performaRemarks;
                            //entUpdate["hil_fixedcharges"] = _fixedCharges;
                            //entUpdate["hil_totalclaimamount"] = new Money(_claimAmount);
                            //entUpdate["hil_totaljobsclosed"] = _totalClosedJobs;
                            //_service.Update(entUpdate);
                            #endregion

                            //Console.WriteLine("Marking UnProcessed Jobs");
                            #region Update Job Claim Status
                            //_fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                            //<entity name='msdyn_workorder'>
                            //<attribute name='msdyn_name' />
                            //<attribute name='hil_claimstatus' />
                            //<attribute name='hil_generateclaim' />
                            //<attribute name='hil_laborinwarranty' />
                            //<filter type='and'>
                            //    <condition attribute='hil_generateclaim' operator='ne' value='1' />
                            //    <condition attribute='hil_isocr' operator='ne' value='1' />
                            //    <filter type='or'>
                            //        <condition attribute='msdyn_substatus' operator='eq' value='{1727FA6C-FA0F-E911-A94E-000D3AF060A1}' />
                            //        <condition attribute='hil_claimstatus' operator='eq' value='7' />
                            //    </filter>
                            //    <condition attribute='msdyn_timeclosed' operator='on-or-after' value='" + (_startDate.Value.Year.ToString() + "-" + _startDate.Value.Month.ToString().PadLeft(2, '0') + "-" + _startDate.Value.Day.ToString().PadLeft(2, '0')) + @"' />
                            //    <condition attribute='msdyn_timeclosed' operator='on-or-before' value='" + (_endDate.Value.Year.ToString() + "-" + _endDate.Value.Month.ToString().PadLeft(2, '0') + "-" + _endDate.Value.Day.ToString().PadLeft(2, '0')) + @"' />
                            //    <condition attribute='hil_owneraccount' operator='eq' value='{" + _franchiseIdRefresh + @"}' />
                            //</filter>
                            //</entity>
                            //</fetch>";
                            //entcollJobs = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                            //i = 1;
                            //foreach (Entity entJobupdate in entcollJobs.Entities)
                            //{
                            //    bool _laborInWarranty = entJobupdate.GetAttributeValue<bool>("hil_laborinwarranty");
                            //    OptionSetValue _clmStatus = entJobupdate.GetAttributeValue<OptionSetValue>("hil_claimstatus");
                            //    entUpdate = new Entity("msdyn_workorder", entJobupdate.Id);
                            //    if (_clmStatus.Value == 4 && _laborInWarranty)
                            //    {
                            //        entUpdate["hil_generateclaim"] = false;
                            //    }
                            //    else
                            //    {
                            //        entUpdate["hil_generateclaim"] = true;
                            //    }
                            //    _service.Update(entUpdate);
                            //    Console.WriteLine(i++.ToString() + "/" + entcollJobs.Entities.Count.ToString());
                            //}
                            #endregion
                        }
                    }
                    catch (Exception ex) { Console.WriteLine(ex.Message); }
                }
            }
        }
        public static void ProcessKKGAuditFailedImpose(IOrganizationService _service)
        {

            #region Variable declaration
            Guid _fiscalMonthId = new Guid("DC417E88-E62B-EB11-A813-0022486E5C9D");
            //Guid _fiscalMonthId = Guid.Empty;
            DateTime? _startDate = null;
            DateTime? _endDate = null;
            Guid _performaInvoiceId = Guid.Empty;
            int _totalClosedJobs = 0, _fixedCharges = 0;
            string _fetchXML = string.Empty;
            QueryExpression queryExp;
            EntityCollection entcoll;
            EntityCollection entcollJobs;
            EntityReference _erFiscalMonth = null;
            EntityReference _erFranchise = null;
            string[] _jobColumns = { "hil_fiscalmonth", "msdyn_name", "createdon", "hil_owneraccount", "hil_channelpartnercategory", "hil_warrantystatus", "hil_countryclassification", "hil_tatachievementslab", "hil_tatcategory", "hil_schemecode", "hil_salesoffice", "hil_productcategory", "hil_kkgcode", "hil_typeofassignee", "hil_brand", "hil_productsubcategory", "hil_callsubtype", "hil_claimstatus", "hil_isocr", "hil_laborinwarranty", "hil_isgascharged", "hil_sparepartuse", "hil_jobclosemobile", "msdyn_timeclosed", "msdyn_customerasset", "hil_quantity", "hil_purchasedate" };
            string _performaRemarks = string.Empty;
            #endregion

            #region Fetching Fiscal Month Date Range
            if (_fiscalMonthId != Guid.Empty)
            {
                Entity ent = _service.Retrieve("hil_claimperiod", _fiscalMonthId, new ColumnSet("hil_fromdate", "hil_todate"));
                if (ent != null)
                {
                    _startDate = ent.GetAttributeValue<DateTime>("hil_fromdate").AddMinutes(330);
                    _endDate = ent.GetAttributeValue<DateTime>("hil_todate").AddMinutes(330);
                    _erFiscalMonth = ent.ToEntityReference();
                }
            }
            else
            {

                DateTime _processDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);

                queryExp = new QueryExpression("hil_claimperiod");
                queryExp.ColumnSet = new ColumnSet("hil_fromdate", "hil_todate");
                queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                queryExp.Criteria.AddCondition("hil_fromdate", ConditionOperator.OnOrBefore, _processDate);
                queryExp.Criteria.AddCondition("hil_todate", ConditionOperator.OnOrAfter, _processDate);
                entcoll = _service.RetrieveMultiple(queryExp);
                if (entcoll.Entities.Count > 0)
                {
                    _startDate = entcoll.Entities[0].GetAttributeValue<DateTime>("hil_fromdate").AddMinutes(330);
                    _endDate = entcoll.Entities[0].GetAttributeValue<DateTime>("hil_todate").AddMinutes(330);
                    _erFiscalMonth = entcoll.Entities[0].ToEntityReference();
                }
            }
            #endregion

            #region Refreshing Claim Parameters & Generating applicable Claim Lines

            _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                <entity name='msdyn_workorder'>
                <attribute name='hil_fiscalmonth' />
                <attribute name='msdyn_name' />
                <attribute name='createdon' />
                <attribute name='hil_owneraccount' />
                <attribute name='hil_channelpartnercategory' />
                <attribute name='hil_warrantystatus' />
                <attribute name='hil_countryclassification' />
                <attribute name='hil_tatachievementslab' />
                <attribute name='hil_tatcategory' />
                <attribute name='hil_schemecode' />
                <attribute name='hil_salesoffice' />
                <attribute name='hil_productcategory' />
                <attribute name='hil_kkgcode' />
                <attribute name='hil_typeofassignee' />
                <attribute name='hil_brand' />
                <attribute name='hil_productsubcategory' />
                <attribute name='hil_callsubtype' />
                <attribute name='hil_claimstatus' />
                <attribute name='hil_isocr' />
                <attribute name='hil_laborinwarranty' />
                <attribute name='hil_isgascharged' />
                <attribute name='hil_sparepartuse' />
                <attribute name='hil_jobclosemobile' />
                <attribute name='msdyn_timeclosed' />
                <attribute name='msdyn_customerasset' />
                <attribute name='hil_quantity' />
                <attribute name='hil_purchasedate' />
                <attribute name='hil_sourceofjob' />
                <filter type='and'>
                    <condition attribute='msdyn_name' operator='eq' value='2609206500949' />
                </filter>
                <link-entity name='account' from='accountid' to='hil_owneraccount' visible='false' link-type='outer' alias='custCatg'>
                <attribute name='hil_category' />
                </link-entity>
                </entity>
                </fetch>";

            entcollJobs = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
            _totalClosedJobs = entcollJobs.Entities.Count;
            int i = 0;
            foreach (Entity entJob in entcollJobs.Entities)
            {
                _performaInvoiceId = Guid.Empty;
                _totalClosedJobs = 0;
                _fixedCharges = 0;
                _erFranchise = entJob.GetAttributeValue<EntityReference>("hil_owneraccount");
                _performaInvoiceId = GeneratePerformaInvoice(_service, _erFranchise.Id, _erFiscalMonth.Id);

                GenerateClaimLines(entJob, _performaInvoiceId, _service, _erFiscalMonth);
                Console.WriteLine("Claim Cline generated");
            }
            #endregion
        }
        public static void ProcessFranchiseClaimsJobWise(IOrganizationService _service)
        {

            #region Variable declaration
            Guid _fiscalMonthId = new Guid("ad387612-46e8-ea11-a817-000d3af05a4b");
            //Guid _fiscalMonthId = Guid.Empty;
            DateTime? _startDate = null;
            DateTime? _endDate = null;
            Guid _performaInvoiceId = Guid.Empty;
            int _totalClosedJobs = 0;
            Decimal _fixedCharges = 0;
            string _fetchXML = string.Empty;
            QueryExpression queryExp;
            EntityCollection entcoll;
            EntityCollection entcollJobs;
            EntityReference _erFiscalMonth = null;
            EntityReference _erFranchise = null;
            OptionSetValue _claimStatus = null;
            string[] _jobColumns = { "hil_fiscalmonth", "msdyn_name", "createdon", "hil_owneraccount", "hil_channelpartnercategory", "hil_warrantystatus", "hil_countryclassification", "hil_tatachievementslab", "hil_tatcategory", "hil_schemecode", "hil_salesoffice", "hil_productcategory", "hil_kkgcode", "hil_typeofassignee", "hil_brand", "hil_productsubcategory", "hil_callsubtype", "hil_claimstatus", "hil_isocr", "hil_laborinwarranty", "hil_isgascharged", "hil_sparepartuse", "hil_jobclosemobile", "msdyn_timeclosed", "msdyn_customerasset", "hil_quantity", "hil_purchasedate" };
            Entity entUpdate;
            string _performaRemarks = string.Empty;
            #endregion

            #region Fetching Fiscal Month Date Range
            if (_fiscalMonthId != Guid.Empty)
            {
                Entity ent = _service.Retrieve("hil_claimperiod", _fiscalMonthId, new ColumnSet("hil_fromdate", "hil_todate"));
                if (ent != null)
                {
                    _startDate = ent.GetAttributeValue<DateTime>("hil_fromdate").AddMinutes(330);
                    _endDate = ent.GetAttributeValue<DateTime>("hil_todate").AddMinutes(330);
                    _erFiscalMonth = ent.ToEntityReference();
                }
            }
            else
            {

                DateTime _processDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);

                queryExp = new QueryExpression("hil_claimperiod");
                queryExp.ColumnSet = new ColumnSet("hil_fromdate", "hil_todate");
                queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                queryExp.Criteria.AddCondition("hil_fromdate", ConditionOperator.OnOrBefore, _processDate);
                queryExp.Criteria.AddCondition("hil_todate", ConditionOperator.OnOrAfter, _processDate);
                entcoll = _service.RetrieveMultiple(queryExp);
                if (entcoll.Entities.Count > 0)
                {
                    _startDate = entcoll.Entities[0].GetAttributeValue<DateTime>("hil_fromdate").AddMinutes(330);
                    _endDate = entcoll.Entities[0].GetAttributeValue<DateTime>("hil_todate").AddMinutes(330);
                    _erFiscalMonth = entcoll.Entities[0].ToEntityReference();
                }
            }
            #endregion

            #region Refreshing Claim Parameters & Generating applicable Claim Lines

            _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                <entity name='msdyn_workorder'>
                <attribute name='hil_fiscalmonth' />
                <attribute name='msdyn_name' />
                <attribute name='createdon' />
                <attribute name='hil_owneraccount' />
                <attribute name='hil_channelpartnercategory' />
                <attribute name='hil_warrantystatus' />
                <attribute name='hil_countryclassification' />
                <attribute name='hil_tatachievementslab' />
                <attribute name='hil_tatcategory' />
                <attribute name='hil_schemecode' />
                <attribute name='hil_salesoffice' />
                <attribute name='hil_productcategory' />
                <attribute name='hil_kkgcode' />
                <attribute name='hil_typeofassignee' />
                <attribute name='hil_brand' />
                <attribute name='hil_productsubcategory' />
                <attribute name='hil_callsubtype' />
                <attribute name='hil_claimstatus' />
                <attribute name='hil_isocr' />
                <attribute name='hil_laborinwarranty' />
                <attribute name='hil_isgascharged' />
                <attribute name='hil_sparepartuse' />
                <attribute name='hil_jobclosemobile' />
                <attribute name='msdyn_timeclosed' />
                <attribute name='msdyn_customerasset' />
                <attribute name='hil_quantity' />
                <attribute name='hil_purchasedate' />
                <attribute name='hil_sourceofjob' />
                <filter type='and'>
                    <condition attribute='hil_isocr' operator='ne' value='1' />
                    <condition attribute='msdyn_name' operator='eq' value='2511207261873' />
                    <condition attribute='msdyn_substatus' operator='eq' value='{1727FA6C-FA0F-E911-A94E-000D3AF060A1}' />
                    <condition attribute='hil_typeofassignee' operator='in'>
                        <value uiname='Franchise' uitype='position'>{4A1AA189-1208-E911-A94D-000D3AF0694E}</value>
                        <value uiname='Franchise Technician' uitype='position'>{0197EA9B-1208-E911-A94D-000D3AF0694E}</value>
                    </condition>
                    <condition attribute='msdyn_timeclosed' operator='on-or-after' value='2020-12-01' />
                    <condition attribute='msdyn_timeclosed' operator='on-or-before' value='2020-12-05' />
                </filter>
                <link-entity name='account' from='accountid' to='hil_owneraccount' visible='false' link-type='outer' alias='custCatg'>
                <attribute name='hil_category' />
                </link-entity>
                </entity>
                </fetch>";
            //<condition attribute='hil_generateclaim' operator='ne' value='1' />
            /*                    <condition attribute='msdyn_timeclosed' operator='on-or-after' value='2020-09-16' />
                    <condition attribute='msdyn_timeclosed' operator='on-or-before' value='2020-09-17' />
            */

            entcollJobs = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
            _totalClosedJobs = entcollJobs.Entities.Count;
            int i = 0;
            foreach (Entity entJob in entcollJobs.Entities)
            {
                _performaInvoiceId = Guid.Empty;
                _totalClosedJobs = 0;
                _fixedCharges = 0;
                _erFranchise = entJob.GetAttributeValue<EntityReference>("hil_owneraccount");
                _performaInvoiceId = GeneratePerformaInvoice(_service, _erFranchise.Id, _erFiscalMonth.Id);

                Console.WriteLine("Job # " + entJob.GetAttributeValue<string>("msdyn_name"));
                #region Fixed Charge Calculation
                _fixedCharges = CreateFixedChargeClaimLine(_service, _performaInvoiceId, ClaimCategoryConst._FixedCompensation, _erFiscalMonth, _erFranchise, Convert.ToDateTime(_endDate));
                #endregion

                _claimStatus = RefreshClaimParameters(_service, _erFiscalMonth.Id, entJob);

                if (_performaInvoiceId != null && _claimStatus.Value == 4)
                {
                    GenerateClaimLines(entJob, _performaInvoiceId, _service, _erFiscalMonth);
                    Console.WriteLine("Claim Cline generated");
                }

                Console.WriteLine("Calculating Claim amount");
                _fetchXML = @"<fetch distinct='false' mapping='logical' aggregate='true'>
        <entity name='hil_claimline'>
        <attribute name='hil_claimamount' alias='claimamount' aggregate='sum'/> 
        <attribute name='hil_jobid' alias='jobid' groupby='true' /> 
        <filter type='and'>
        <filter type='or'>
            <condition attribute='hil_jobid' operator='eq' value='{" + entJob.Id + @"}' />
        </filter>
        </filter>
        </entity>
        </fetch>";
                Decimal _JobClaimAmount = 0;
                EntityCollection entcollCA = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                if (entcollCA.Entities.Count > 0)
                {
                    Guid _claimJobId = ((EntityReference)((AliasedValue)entcollCA.Entities[0].Attributes["jobid"]).Value).Id;
                    _JobClaimAmount = ((Money)((AliasedValue)entcollCA.Entities[0]["claimamount"]).Value).Value;
                    if (_claimJobId != null)
                    {
                        entUpdate = new Entity("msdyn_workorder", _claimJobId);
                        entUpdate["hil_generateclaim"] = true;
                        entUpdate["hil_claimamount"] = new Money(_JobClaimAmount);
                        _service.Update(entUpdate);
                    }
                }
                else
                {
                    entUpdate = new Entity("msdyn_workorder", entJob.Id);
                    entUpdate["hil_generateclaim"] = true;
                    entUpdate["hil_claimamount"] = new Money(0);
                    _service.Update(entUpdate);
                }
                Console.WriteLine(i++.ToString() + "/" + _totalClosedJobs.ToString());
            }
            #endregion
        }
        public static void ProcessFranchiseClaimsJobWiseBackLog(IOrganizationService _service)
        {

            #region Variable declaration
            Guid _fiscalMonthId = new Guid("DC417E88-E62B-EB11-A813-0022486E5C9D");
            //Guid _fiscalMonthId = Guid.Empty;
            DateTime? _startDate = null;
            DateTime? _endDate = null;
            Guid _performaInvoiceId = Guid.Empty;
            int _totalClosedJobs = 0;
            Decimal _fixedCharges = 0;
            string _fetchXML = string.Empty;
            QueryExpression queryExp;
            EntityCollection entcoll;
            EntityCollection entcollJobs;
            EntityCollection entcollTATSlabJobs;
            EntityReference _erFiscalMonth = null;
            EntityReference _erFranchise = null;
            OptionSetValue _claimStatus = null;
            string[] _jobColumns = { "hil_fiscalmonth", "msdyn_name", "createdon", "hil_owneraccount", "hil_channelpartnercategory", "hil_warrantystatus", "hil_countryclassification", "hil_tatachievementslab", "hil_tatcategory", "hil_schemecode", "hil_salesoffice", "hil_productcategory", "hil_kkgcode", "hil_typeofassignee", "hil_brand", "hil_productsubcategory", "hil_callsubtype", "hil_claimstatus", "hil_isocr", "hil_laborinwarranty", "hil_isgascharged", "hil_sparepartuse", "hil_jobclosemobile", "msdyn_timeclosed", "msdyn_customerasset", "hil_quantity", "hil_purchasedate" };
            Entity entUpdate;
            string _performaRemarks = string.Empty;
            #endregion

            #region Fetching Fiscal Month Date Range
            if (_fiscalMonthId != Guid.Empty)
            {
                Entity ent = _service.Retrieve("hil_claimperiod", _fiscalMonthId, new ColumnSet("hil_fromdate", "hil_todate"));
                if (ent != null)
                {
                    _startDate = ent.GetAttributeValue<DateTime>("hil_fromdate").AddMinutes(330);
                    _endDate = ent.GetAttributeValue<DateTime>("hil_todate").AddMinutes(330);
                    _erFiscalMonth = ent.ToEntityReference();
                }
            }
            else
            {

                DateTime _processDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);

                queryExp = new QueryExpression("hil_claimperiod");
                queryExp.ColumnSet = new ColumnSet("hil_fromdate", "hil_todate");
                queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                queryExp.Criteria.AddCondition("hil_fromdate", ConditionOperator.OnOrBefore, _processDate);
                queryExp.Criteria.AddCondition("hil_todate", ConditionOperator.OnOrAfter, _processDate);
                entcoll = _service.RetrieveMultiple(queryExp);
                if (entcoll.Entities.Count > 0)
                {
                    _startDate = entcoll.Entities[0].GetAttributeValue<DateTime>("hil_fromdate").AddMinutes(330);
                    _endDate = entcoll.Entities[0].GetAttributeValue<DateTime>("hil_todate").AddMinutes(330);
                    _erFiscalMonth = entcoll.Entities[0].ToEntityReference();
                }
            }
            #endregion

            #region Refreshing Claim Parameters & Generating applicable Claim Lines
            try
            {
                /*
                 *<link-entity name='account' from='accountid' to='hil_owneraccount' link-type='inner' alias='ad'>
                    <filter type='and'>
                    <condition attribute='hil_salesoffice' operator='eq' value='{a1210df3-bcf7-e811-a94c-000d3af06091}' />
                    </filter>
                </link-entity>
                 * */

                //_fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                //<entity name='msdyn_workorder'>
                //<attribute name='hil_owneraccount' />
                //<filter type='and'>
                //    <condition attribute='hil_owneraccount' operator='eq' value='{5683FA3C-7FB6-EA11-A812-000D3AF05D7B}' />
                //    <condition attribute='hil_isocr' operator='ne' value='1' />
                //    <condition attribute='hil_generateclaim' operator='ne' value='1' />
                //    <filter type='or'>
                //        <condition attribute='msdyn_substatus' operator='eq' value='{1727FA6C-FA0F-E911-A94E-000D3AF060A1}' />
                //        <condition attribute='hil_claimstatus' operator='eq' value='4' />
                //    </filter>
                //    <condition attribute='msdyn_timeclosed' operator='on-or-after' value='" + (_startDate.Value.Year.ToString() + "-" + _startDate.Value.Month.ToString().PadLeft(2, '0') + "-" + _startDate.Value.Day.ToString().PadLeft(2, '0')) + @"' />
                //    <condition attribute='msdyn_timeclosed' operator='on-or-before' value='" + (_endDate.Value.Year.ToString() + "-" + _endDate.Value.Month.ToString().PadLeft(2, '0') + "-" + _endDate.Value.Day.ToString().PadLeft(2, '0')) + @"' />
                //    <condition attribute='hil_typeofassignee' operator='in'>
                //    <value uiname='Franchise' uitype='position'>{4A1AA189-1208-E911-A94D-000D3AF0694E}</value>
                //    <value uiname='Franchise Technician' uitype='position'>{0197EA9B-1208-E911-A94D-000D3AF0694E}</value>
                //    </condition>
                //</filter>
                //<link-entity name='account' from='accountid' to='hil_owneraccount' link-type='inner' alias='ad'>
                //    <filter type='and'>
                //    <condition attribute='hil_salesoffice' operator='eq' value='{a1210df3-bcf7-e811-a94c-000d3af06091}' />
                //    </filter>
                //</link-entity>
                //</entity>
                //</fetch>";

                _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                <entity name='msdyn_workorder'>
                <attribute name='hil_owneraccount' />
                <filter type='and'>
                    <condition attribute='hil_owneraccount' operator='eq' value='{5683FA3C-7FB6-EA11-A812-000D3AF05D7B}' />
                    <condition attribute='hil_isocr' operator='ne' value='1' />
                    <filter type='or'>
                        <condition attribute='msdyn_substatus' operator='eq' value='{1727FA6C-FA0F-E911-A94E-000D3AF060A1}' />
                        <condition attribute='hil_claimstatus' operator='eq' value='4' />
                    </filter>
                    <condition attribute='msdyn_timeclosed' operator='on-or-after' value='" + (_startDate.Value.Year.ToString() + "-" + _startDate.Value.Month.ToString().PadLeft(2, '0') + "-" + _startDate.Value.Day.ToString().PadLeft(2, '0')) + @"' />
                    <condition attribute='msdyn_timeclosed' operator='on-or-before' value='" + (_endDate.Value.Year.ToString() + "-" + _endDate.Value.Month.ToString().PadLeft(2, '0') + "-" + _endDate.Value.Day.ToString().PadLeft(2, '0')) + @"' />
                    <condition attribute='hil_typeofassignee' operator='in'>
                    <value uiname='Franchise' uitype='position'>{4A1AA189-1208-E911-A94D-000D3AF0694E}</value>
                    <value uiname='Franchise Technician' uitype='position'>{0197EA9B-1208-E911-A94D-000D3AF0694E}</value>
                    </condition>
                </filter>
                </entity>
                </fetch>";

                entcoll = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                foreach (Entity entASP in entcoll.Entities)
                {
                    _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                     <entity name='msdyn_workorder'>
                        <attribute name='hil_fiscalmonth' />
                        <attribute name='msdyn_name' />
                        <attribute name='createdon' />
                        <attribute name='hil_owneraccount' />
                        <attribute name='hil_channelpartnercategory' />
                        <attribute name='hil_warrantystatus' />
                        <attribute name='hil_countryclassification' />
                        <attribute name='hil_tatachievementslab' />
                        <attribute name='hil_tatcategory' />
                        <attribute name='hil_schemecode' />
                        <attribute name='hil_salesoffice' />
                        <attribute name='hil_productcategory' />
                        <attribute name='hil_kkgcode' />
                        <attribute name='hil_typeofassignee' />
                        <attribute name='hil_brand' />
                        <attribute name='hil_productsubcategory' />
                        <attribute name='hil_callsubtype' />
                        <attribute name='hil_claimstatus' />
                        <attribute name='hil_isocr' />
                        <attribute name='hil_laborinwarranty' />
                        <attribute name='hil_isgascharged' />
                        <attribute name='hil_sparepartuse' />
                        <attribute name='hil_jobclosemobile' />
                        <attribute name='msdyn_timeclosed' />
                        <attribute name='msdyn_customerasset' />
                        <attribute name='hil_quantity' />
                        <attribute name='hil_purchasedate' />
                        <attribute name='hil_sourceofjob' />
                        <attribute name='hil_pincode' />
                        <filter type='and'>
                            <condition attribute='hil_owneraccount' operator='eq'  value='{5683FA3C-7FB6-EA11-A812-000D3AF05D7B}' />
                            <condition attribute='hil_isocr' operator='ne' value='1' />
                            <condition attribute='hil_typeofassignee' operator='in'>
                                <value uiname='Franchise' uitype='position'>{4A1AA189-1208-E911-A94D-000D3AF0694E}</value>
                                <value uiname='Franchise Technician' uitype='position'>{0197EA9B-1208-E911-A94D-000D3AF0694E}</value>
                            </condition>
                            <condition attribute='msdyn_timeclosed' operator='on-or-after' value='2020-10-21' />
                            <condition attribute='msdyn_timeclosed' operator='on-or-before' value='2020-11-20' />
                            <filter type='or'>
                                <condition attribute='hil_claimstatus' operator='eq' value='4' />
                                <condition attribute='msdyn_substatus' operator='eq' value='{1727FA6C-FA0F-E911-A94E-000D3AF060A1}' />
                            </filter>
                        </filter>
                        <link-entity name='account' from='accountid' to='hil_owneraccount' visible='false' link-type='outer' alias='custCatg'>
                        <attribute name='hil_category' />
                        </link-entity>
                    </entity>
                    </fetch>";
                    try
                    {
                        entcollJobs = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                        _totalClosedJobs = entcollJobs.Entities.Count;
                        int i = 0;
                        foreach (Entity entJob in entcollJobs.Entities)
                        {
                            _performaInvoiceId = Guid.Empty;
                            _totalClosedJobs = 0;
                            _fixedCharges = 0;
                            _erFranchise = entJob.GetAttributeValue<EntityReference>("hil_owneraccount");
                            _performaInvoiceId = GeneratePerformaInvoice(_service, _erFranchise.Id, _erFiscalMonth.Id);

                            Console.WriteLine("Job # " + entJob.GetAttributeValue<string>("msdyn_name"));
                            #region Fixed Charge Calculation
                            _fixedCharges = CreateFixedChargeClaimLine(_service, _performaInvoiceId, ClaimCategoryConst._FixedCompensation, _erFiscalMonth, _erFranchise, Convert.ToDateTime(_endDate));
                            #endregion

                            //_claimStatus = RefreshClaimParameters(_service, _erFiscalMonth.Id, entJob);

                            _claimStatus = entJob.GetAttributeValue<OptionSetValue>("hil_claimstatus");

                            if (_performaInvoiceId != null && _claimStatus.Value == 4)
                            {
                                GenerateClaimLines(entJob, _performaInvoiceId, _service, _erFiscalMonth);
                                Console.WriteLine("Claim Cline generated");
                            }

                            //GenerateClaimLines(entJob, _performaInvoiceId, _service, _erFiscalMonth);
                            //Console.WriteLine("Claim Cline generated");

                            entUpdate = new Entity("msdyn_workorder", entJob.Id);
                            entUpdate["hil_generateclaim"] = true;
                            entUpdate["hil_claimamount"] = new Money(0);
                            _service.Update(entUpdate);

                            //Console.WriteLine("Calculating Claim amount");
                            //_fetchXML = @"<fetch distinct='false' mapping='logical' aggregate='true'>
                            //<entity name='hil_claimline'>
                            //<attribute name='hil_claimamount' alias='claimamount' aggregate='sum'/> 
                            //<attribute name='hil_jobid' alias='jobid' groupby='true' /> 
                            //<filter type='and'>
                            //<filter type='or'>
                            //    <condition attribute='hil_jobid' operator='eq' value='{" + entJob.Id + @"}' />
                            //</filter>
                            //</filter>
                            //</entity>
                            //</fetch>";
                            //Decimal _JobClaimAmount = 0;
                            //EntityCollection entcollCA = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                            //if (entcollCA.Entities.Count > 0)
                            //{
                            //    Guid _claimJobId = ((EntityReference)((AliasedValue)entcollCA.Entities[0].Attributes["jobid"]).Value).Id;
                            //    _JobClaimAmount = ((Money)((AliasedValue)entcollCA.Entities[0]["claimamount"]).Value).Value;
                            //    if (_claimJobId != null)
                            //    {
                            //        entUpdate = new Entity("msdyn_workorder", _claimJobId);
                            //        entUpdate["hil_generateclaim"] = true;
                            //        entUpdate["hil_claimamount"] = new Money(_JobClaimAmount);
                            //        _service.Update(entUpdate);
                            //    }
                            //}
                            //else
                            //{
                            //    entUpdate = new Entity("msdyn_workorder", entJob.Id);
                            //    entUpdate["hil_generateclaim"] = true;
                            //    entUpdate["hil_claimamount"] = new Money(0);
                            //    _service.Update(entUpdate);
                            //}
                            Console.WriteLine(i++.ToString() + "/" + _totalClosedJobs.ToString());
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    #region Calculating TAT Incentive Slab
                    Console.WriteLine("Calculating TAT Incentive Slab");
                    // Generate TAT Slab if all Claim status of all Work Order of that Franchisee not in (CLaim approved, Claim Rejected, Penalty Imposed, Penalty Waived)
                    queryExp = new QueryExpression(msdyn_workorder.EntityLogicalName);
                    queryExp.ColumnSet = new ColumnSet("hil_claimstatus", "msdyn_workorderid", "hil_tatachievementslab");
                    queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                    queryExp.Criteria.AddCondition("hil_claimstatus", ConditionOperator.NotIn, new object[] {0, 4, 3, 8, 7 });
                    queryExp.Criteria.AddCondition("msdyn_substatus", ConditionOperator.Equal, new Guid("1727FA6C-FA0F-E911-A94E-000D3AF060A1"));
                    queryExp.Criteria.AddCondition("hil_owneraccount", ConditionOperator.Equal, entASP.GetAttributeValue<EntityReference>("hil_owneraccount").Id);
                    //queryExp.Criteria.AddCondition("hil_owneraccount", ConditionOperator.Equal,new Guid("5683FA3C-7FB6-EA11-A812-000D3AF05D7B"));
                    queryExp.Criteria.AddCondition("hil_isocr", ConditionOperator.NotEqual, true);
                    queryExp.Criteria.AddCondition(new ConditionExpression("msdyn_timeclosed", ConditionOperator.OnOrAfter, _startDate));
                    queryExp.Criteria.AddCondition(new ConditionExpression("msdyn_timeclosed", ConditionOperator.OnOrBefore, _endDate));
                    entcollTATSlabJobs = _service.RetrieveMultiple(queryExp);
                    if (entcollTATSlabJobs.Entities.Count == 0)
                    {
                        GenerateTATClaimLines(_service, _performaInvoiceId, _erFranchise.Id, Convert.ToDateTime(_startDate), Convert.ToDateTime(_endDate));
                    }
                    #endregion
                }
                #endregion

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        public static void RefreshClaimParameters(IOrganizationService _service)
        {
            #region Variable declaration
            Guid _fiscalMonthId = Guid.Empty;
            DateTime? _startDate = null;
            DateTime? _endDate = null;
            string _strStartDate = null;
            string _strEndDate = null;
            Guid _performaInvoiceId = Guid.Empty;
            int _totalClosedJobs = 0;
            string _fetchXML = string.Empty;
            QueryExpression queryExp;
            EntityCollection entcoll;
            EntityCollection entcollJobs;
            EntityReference _erFiscalMonth = null;
            OptionSetValue _claimStatus = null;
            string[] _jobColumns = { "hil_fiscalmonth", "msdyn_name", "createdon", "hil_owneraccount", "hil_channelpartnercategory", "hil_warrantystatus", "hil_countryclassification", "hil_tatachievementslab", "hil_tatcategory", "hil_schemecode", "hil_salesoffice", "hil_productcategory", "hil_kkgcode", "hil_typeofassignee", "hil_brand", "hil_productsubcategory", "hil_callsubtype", "hil_claimstatus", "hil_isocr", "hil_laborinwarranty", "hil_isgascharged", "hil_sparepartuse", "hil_jobclosemobile", "msdyn_timeclosed", "msdyn_customerasset", "hil_quantity", "hil_purchasedate" };
            string _performaRemarks = string.Empty;
            #endregion

            #region Fetching Fiscal Month Date Range
            
            //DateTime _processDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);
            
            var _salesOfficeId = ConfigurationManager.AppSettings["fromDate"].ToString();
            DateTime _processDate = Convert.ToDateTime(_salesOfficeId);

            queryExp = new QueryExpression("hil_claimperiod");
            queryExp.ColumnSet = new ColumnSet("hil_fromdate", "hil_todate");
            queryExp.Criteria = new FilterExpression(LogicalOperator.And);
            queryExp.Criteria.AddCondition("hil_fromdate", ConditionOperator.OnOrBefore, _processDate);
            queryExp.Criteria.AddCondition("hil_todate", ConditionOperator.OnOrAfter, _processDate);
            entcoll = _service.RetrieveMultiple(queryExp);
            if (entcoll.Entities.Count > 0)
            {
                _startDate = entcoll.Entities[0].GetAttributeValue<DateTime>("hil_fromdate").AddMinutes(330);
                _strStartDate = _startDate.Value.Year.ToString() + "-" + _startDate.Value.Month.ToString().PadLeft(2, '0') + "-" + _startDate.Value.Day.ToString().PadLeft(2, '0');
                _endDate = entcoll.Entities[0].GetAttributeValue<DateTime>("hil_todate").AddMinutes(330);
                _strEndDate = _endDate.Value.Year.ToString() + "-" + _endDate.Value.Month.ToString().PadLeft(2, '0') + "-" + _endDate.Value.Day.ToString().PadLeft(2, '0');
                _erFiscalMonth = entcoll.Entities[0].ToEntityReference();
            }
            if (_erFiscalMonth == null)
            {
                Console.WriteLine("Fiscal Month is not defined." + DateTime.Now.ToString());
                return;
            }
            #endregion

            #region Refreshing Claim Parameters & Generating applicable Claim Lines
            Console.WriteLine("Batch Started " + DateTime.Now.ToString());
            try
            {
                _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' top='50'>
                     <entity name='msdyn_workorder'>
                        <attribute name='hil_fiscalmonth' />
                        <attribute name='msdyn_name' />
                        <attribute name='createdon' />
                        <attribute name='hil_owneraccount' />
                        <attribute name='hil_channelpartnercategory' />
                        <attribute name='hil_warrantystatus' />
                        <attribute name='hil_countryclassification' />
                        <attribute name='hil_tatachievementslab' />
                        <attribute name='hil_tatcategory' />
                        <attribute name='hil_schemecode' />
                        <attribute name='hil_salesoffice' />
                        <attribute name='hil_productcategory' />
                        <attribute name='hil_kkgcode' />
                        <attribute name='hil_typeofassignee' />
                        <attribute name='hil_brand' />
                        <attribute name='hil_productsubcategory' />
                        <attribute name='hil_callsubtype' />
                        <attribute name='hil_claimstatus' />
                        <attribute name='hil_isocr' />
                        <attribute name='hil_district' />
                        <attribute name='hil_laborinwarranty' />
                        <attribute name='hil_isgascharged' />
                        <attribute name='hil_sparepartuse' />
                        <attribute name='hil_jobclosemobile' />
                        <attribute name='msdyn_timeclosed' />
                        <attribute name='msdyn_customerasset' />
                        <attribute name='hil_quantity' />
                        <attribute name='hil_purchasedate' />
                        <attribute name='hil_sourceofjob' />
                        <attribute name='hil_pincode' />
                        <attribute name='msdyn_substatus' />
                        <order attribute='msdyn_timeclosed' descending='false' />
                        <filter type='and'>
                            <condition attribute='hil_owneraccount' operator='not-null' />
                            <condition attribute='hil_salesoffice' operator='not-null' />
                            <condition attribute='hil_callsubtype' operator='not-null' />
                            <condition attribute='msdyn_customerasset' operator='not-null' />
                            <condition attribute='hil_productsubcategory' operator='not-null' />
                            <condition attribute='hil_productcategory' operator='not-null' />
                            <condition attribute='msdyn_substatus' operator='eq' value='{1727FA6C-FA0F-E911-A94E-000D3AF060A1}' />
                            <condition attribute='hil_isocr' operator='ne' value='1' />
                            <condition attribute='hil_claimstatus' operator='ne' value='3' />
                            <condition attribute='hil_fiscalmonth' operator='ne' value='" + _erFiscalMonth.Id + @"' />
                            <condition attribute='msdyn_timeclosed' operator='on-or-after' value='" + _strStartDate + @"' />
                            <condition attribute='msdyn_timeclosed' operator='on-or-before' value='" + _strEndDate + @"' />
                            <condition attribute='hil_typeofassignee' operator='in'>
                                <value uiname='Franchise' uitype='position'>{4A1AA189-1208-E911-A94D-000D3AF0694E}</value>
                                <value uiname='Franchise Technician' uitype='position'>{0197EA9B-1208-E911-A94D-000D3AF0694E}</value>
                            </condition>
                        </filter>
                        <link-entity name='account' from='accountid' to='hil_owneraccount' visible='false' link-type='outer' alias='custCatg'>
                        <attribute name='hil_category' />
                        </link-entity>
                        <link-entity name='account' from='accountid' to='hil_owneraccount' link-type='outer' alias='DistASP'>
                            <attribute name='hil_district' />
                        </link-entity>
                    </entity>
                    </fetch>";

                try
                {
                    while (true)
                    {
                        entcollJobs = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                        _totalClosedJobs = entcollJobs.Entities.Count;
                        int i = 1;
                        if (entcollJobs.Entities.Count == 0) { break; }
                        foreach (Entity entJob in entcollJobs.Entities)
                        {
                            try
                            {
                                _claimStatus = RefreshClaimParameters(_service, _erFiscalMonth.Id, entJob);
                                Console.WriteLine(i++.ToString() + "/" + _totalClosedJobs.ToString() + "/" + entJob.GetAttributeValue<string>("msdyn_name") + "/" + entJob.GetAttributeValue<DateTime>("msdyn_timeclosed").AddMinutes(330).ToString());
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.Message);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
            #endregion
            //}
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            Console.WriteLine("Batch Ended " + DateTime.Now.ToString());
        }

        public static void RefreshClaimParametersTest(IOrganizationService _service)
        {
            #region Variable declaration
            Guid _fiscalMonthId = Guid.Empty;
            DateTime? _startDate = null;
            DateTime? _endDate = null;
            string _strStartDate = null;
            string _strEndDate = null;
            Guid _performaInvoiceId = Guid.Empty;
            int _totalClosedJobs = 0;
            string _fetchXML = string.Empty;
            QueryExpression queryExp;
            EntityCollection entcoll;
            EntityCollection entcollJobs;
            EntityReference _erFiscalMonth = null;
            OptionSetValue _claimStatus = null;
            string[] _jobColumns = { "hil_fiscalmonth", "msdyn_name", "createdon", "hil_owneraccount", "hil_channelpartnercategory", "hil_warrantystatus", "hil_countryclassification", "hil_tatachievementslab", "hil_tatcategory", "hil_schemecode", "hil_salesoffice", "hil_productcategory", "hil_kkgcode", "hil_typeofassignee", "hil_brand", "hil_productsubcategory", "hil_callsubtype", "hil_claimstatus", "hil_isocr", "hil_laborinwarranty", "hil_isgascharged", "hil_sparepartuse", "hil_jobclosemobile", "msdyn_timeclosed", "msdyn_customerasset", "hil_quantity", "hil_purchasedate" };
            string _performaRemarks = string.Empty;
            #endregion

            #region Fetching Fiscal Month Date Range

            //DateTime _processDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);

            var _salesOfficeId = ConfigurationManager.AppSettings["fromDate"].ToString();
            DateTime _processDate = Convert.ToDateTime(_salesOfficeId);

            queryExp = new QueryExpression("hil_claimperiod");
            queryExp.ColumnSet = new ColumnSet("hil_fromdate", "hil_todate");
            queryExp.Criteria = new FilterExpression(LogicalOperator.And);
            queryExp.Criteria.AddCondition("hil_fromdate", ConditionOperator.OnOrBefore, _processDate);
            queryExp.Criteria.AddCondition("hil_todate", ConditionOperator.OnOrAfter, _processDate);
            entcoll = _service.RetrieveMultiple(queryExp);
            if (entcoll.Entities.Count > 0)
            {
                _startDate = entcoll.Entities[0].GetAttributeValue<DateTime>("hil_fromdate").AddMinutes(330);
                _strStartDate = _startDate.Value.Year.ToString() + "-" + _startDate.Value.Month.ToString().PadLeft(2, '0') + "-" + _startDate.Value.Day.ToString().PadLeft(2, '0');
                _endDate = entcoll.Entities[0].GetAttributeValue<DateTime>("hil_todate").AddMinutes(330);
                _strEndDate = _endDate.Value.Year.ToString() + "-" + _endDate.Value.Month.ToString().PadLeft(2, '0') + "-" + _endDate.Value.Day.ToString().PadLeft(2, '0');
                _erFiscalMonth = entcoll.Entities[0].ToEntityReference();
            }
            if (_erFiscalMonth == null)
            {
                Console.WriteLine("Fiscal Month is not defined." + DateTime.Now.ToString());
                return;
            }
            #endregion

            #region Refreshing Claim Parameters & Generating applicable Claim Lines
            Console.WriteLine("Batch Started " + DateTime.Now.ToString());
            try
            {
                //<condition attribute='hil_district' operator='not-null' /> -- Need to check
                //<condition attribute='hil_claimstatus' operator='ne' value='3' />

                _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                     <entity name='msdyn_workorder'>
                        <attribute name='hil_fiscalmonth' />
                        <attribute name='msdyn_name' />
                        <attribute name='createdon' />
                        <attribute name='hil_owneraccount' />
                        <attribute name='hil_channelpartnercategory' />
                        <attribute name='hil_warrantystatus' />
                        <attribute name='hil_countryclassification' />
                        <attribute name='hil_tatachievementslab' />
                        <attribute name='hil_tatcategory' />
                        <attribute name='hil_schemecode' />
                        <attribute name='hil_salesoffice' />
                        <attribute name='hil_productcategory' />
                        <attribute name='hil_kkgcode' />
                        <attribute name='hil_typeofassignee' />
                        <attribute name='hil_brand' />
                        <attribute name='hil_productsubcategory' />
                        <attribute name='hil_callsubtype' />
                        <attribute name='hil_claimstatus' />
                        <attribute name='hil_isocr' />
                        <attribute name='hil_district' />
                        <attribute name='hil_laborinwarranty' />
                        <attribute name='hil_isgascharged' />
                        <attribute name='hil_sparepartuse' />
                        <attribute name='hil_jobclosemobile' />
                        <attribute name='msdyn_timeclosed' />
                        <attribute name='msdyn_customerasset' />
                        <attribute name='hil_quantity' />
                        <attribute name='hil_purchasedate' />
                        <attribute name='hil_sourceofjob' />
                        <attribute name='hil_pincode' />
                        <attribute name='msdyn_substatus' />
                        <order attribute='msdyn_timeclosed' descending='false' />
                        <filter type='and'>
                            <condition attribute='hil_owneraccount' operator='not-null' />
                            <condition attribute='hil_salesoffice' operator='not-null' />
                            <condition attribute='hil_callsubtype' operator='not-null' />
                            <condition attribute='msdyn_customerasset' operator='not-null' />
                            <condition attribute='hil_productsubcategory' operator='not-null' />
                            <condition attribute='hil_productcategory' operator='not-null' />
                            <condition attribute='hil_fiscalmonth' operator='ne' value='" + _erFiscalMonth.Id + @"' />
                            <condition attribute='msdyn_substatus' operator='eq' value='{1727FA6C-FA0F-E911-A94E-000D3AF060A1}' />
                            <condition attribute='hil_isocr' operator='ne' value='1' />
                            <condition attribute='msdyn_timeclosed' operator='on-or-after' value='" + _strStartDate + @"' />
                            <condition attribute='msdyn_timeclosed' operator='on-or-before' value='" + _strEndDate + @"' />
                            <condition attribute='hil_typeofassignee' operator='in'>
                                <value uiname='Franchise' uitype='position'>{4A1AA189-1208-E911-A94D-000D3AF0694E}</value>
                                <value uiname='Franchise Technician' uitype='position'>{0197EA9B-1208-E911-A94D-000D3AF0694E}</value>
                            </condition>
                            <condition attribute='msdyn_name' operator='eq' value='26102217739436' />  
                        </filter>
                        <link-entity name='account' from='accountid' to='hil_owneraccount' visible='false' link-type='outer' alias='custCatg'>
                        <attribute name='hil_category' />
                        </link-entity>
                        <link-entity name='account' from='accountid' to='hil_owneraccount' link-type='outer' alias='DistASP'>
                            <attribute name='hil_district' />
                        </link-entity>
                    </entity>
                    </fetch>";

                try
                {
                    while (true)
                    {
                        entcollJobs = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                        _totalClosedJobs = entcollJobs.Entities.Count;
                        int i = 1;
                        if (entcollJobs.Entities.Count == 0) { break; }
                        foreach (Entity entJob in entcollJobs.Entities)
                        {
                            try
                            {
                                _claimStatus = RefreshClaimParameters(_service, _erFiscalMonth.Id, entJob);
                                Console.WriteLine(i++.ToString() + "/" + _totalClosedJobs.ToString() + "/" + entJob.GetAttributeValue<string>("msdyn_name") + "/" + entJob.GetAttributeValue<DateTime>("msdyn_timeclosed").AddMinutes(330).ToString());
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.Message);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
            #endregion
            //}
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            Console.WriteLine("Batch Ended " + DateTime.Now.ToString());
        }
        public static void RefreshClaimParametersForKKGAuditFailedJobs(IOrganizationService _service)
        {
            #region Variable declaration
            Guid _fiscalMonthId = new Guid("3d4de739-89e9-eb11-bacb-000d3af0c552");
            //Guid _fiscalMonthId = Guid.Empty;
            DateTime? _startDate = null;
            DateTime? _endDate = null;
            string _strStartDate = null;
            string _strEndDate = null;
            Guid _performaInvoiceId = Guid.Empty;
            int _totalClosedJobs = 0;
            string _fetchXML = string.Empty;
            QueryExpression queryExp;
            EntityCollection entcoll;
            EntityCollection entcollJobs;
            EntityReference _erFiscalMonth = null;
            EntityReference _erFranchise = null;
            OptionSetValue _claimStatus = null;
            Decimal _fixedCharges = 0;
            Entity entUpdate;
            string[] _jobColumns = { "hil_fiscalmonth", "msdyn_name", "createdon", "hil_owneraccount", "hil_channelpartnercategory", "hil_warrantystatus", "hil_countryclassification", "hil_tatachievementslab", "hil_tatcategory", "hil_schemecode", "hil_salesoffice", "hil_productcategory", "hil_kkgcode", "hil_typeofassignee", "hil_brand", "hil_productsubcategory", "hil_callsubtype", "hil_claimstatus", "hil_isocr", "hil_laborinwarranty", "hil_isgascharged", "hil_sparepartuse", "hil_jobclosemobile", "msdyn_timeclosed", "msdyn_customerasset", "hil_quantity", "hil_purchasedate" };
            string _performaRemarks = string.Empty;
            #endregion

            #region Fetching Fiscal Month Date Range
            //DateTime _processDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);
            queryExp = new QueryExpression("hil_claimperiod");
            queryExp.ColumnSet = new ColumnSet("hil_fromdate", "hil_todate");
            queryExp.Criteria = new FilterExpression(LogicalOperator.And);
            queryExp.Criteria.AddCondition("hil_isapplicable", ConditionOperator.Equal, true);
            queryExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
            entcoll = _service.RetrieveMultiple(queryExp);
            if (entcoll.Entities.Count > 0)
            {
                _startDate = entcoll.Entities[0].GetAttributeValue<DateTime>("hil_fromdate").AddMinutes(330);
                _strStartDate = _startDate.Value.Year.ToString() + "-" + _startDate.Value.Month.ToString().PadLeft(2, '0') + "-" + _startDate.Value.Day.ToString().PadLeft(2, '0');
                _endDate = entcoll.Entities[0].GetAttributeValue<DateTime>("hil_todate").AddMinutes(330);
                _strEndDate = _endDate.Value.Year.ToString() + "-" + _endDate.Value.Month.ToString().PadLeft(2, '0') + "-" + _endDate.Value.Day.ToString().PadLeft(2, '0');
                _erFiscalMonth = entcoll.Entities[0].ToEntityReference();
            }
            if (_erFiscalMonth == null)
            {
                Console.WriteLine("Fiscal Month is not defined." + DateTime.Now.ToString());
                return;
            }
            #endregion

            #region Refreshing Claim Parameters & Generating applicable Claim Lines
            Console.WriteLine("Batch Started " + DateTime.Now.ToString());
            try
            {
                _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                  <entity name='hil_sawactivity'>
                    <attribute name='hil_sawactivityid' />
                    <attribute name='hil_name' />
                    <attribute name='createdon' />
                    <attribute name='hil_jobid' />
                    <order attribute='hil_name' descending='false' />
                    <filter type='and'>
                      <condition attribute='createdon' operator='on-or-after' value='"+ _strStartDate  + @"' />
                      <condition attribute='createdon' operator='on-or-before' value='" + _strEndDate + @"' />
                      <condition attribute='hil_sawcategory' operator='eq' value='{E123BD08-8ADD-EA11-A813-000D3AF05A4B}' />
                      <condition attribute='hil_approvalstatus' operator='eq' value='4' />
                    </filter>
                    <link-entity name='msdyn_workorder' from='msdyn_workorderid' to='hil_jobid' link-type='inner' alias='wo'>
                        <filter type='and'>
                        <filter type='or'>
                            <condition attribute='hil_fiscalmonth' operator='ne' value='" + _erFiscalMonth.Id + @"' />
                            <condition attribute='hil_claimstatus' operator='ne' value='7' />
                        </filter>
                        <condition attribute='hil_typeofassignee' operator='in'>
                            <value uiname='Franchise' uitype='position'>{4A1AA189-1208-E911-A94D-000D3AF0694E}</value>
                            <value uiname='Franchise Technician' uitype='position'>{0197EA9B-1208-E911-A94D-000D3AF0694E}</value>
                        </condition>
                        <condition attribute='hil_generateclaim' operator='eq' value='0' />
                        </filter>
                    </link-entity>
                  </entity>
                </fetch>";
                int i = 0;
                entcoll = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                _totalClosedJobs = entcoll.Entities.Count;
                foreach (Entity entASP in entcoll.Entities)
                {
                    string _fetchAssetDetail = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                          <entity name='msdyn_workorderincident'>
                            <attribute name='msdyn_customerasset' />
                            <attribute name='msdyn_workorderincidentid' />
                            <order attribute='msdyn_customerasset' descending='false' />
                            <filter type='and'>
                              <condition attribute='msdyn_workorder' operator='eq' value='{entASP.GetAttributeValue<EntityReference>("hil_jobid").Id}' />
                              <condition attribute='statecode' operator='eq' value='0' />
                            </filter>
                          </entity>
                        </fetch>";
                    EntityCollection entAsset = _service.RetrieveMultiple(new FetchExpression(_fetchAssetDetail));
                    if (entAsset.Entities.Count > 0) {

                        Entity _woUpdate = new Entity("msdyn_workorder", entASP.GetAttributeValue<EntityReference>("hil_jobid").Id);
                        _woUpdate["msdyn_customerasset"] = entAsset.Entities[0].GetAttributeValue<EntityReference>("msdyn_customerasset");
                        _service.Update(_woUpdate);
                    }

                    _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                     <entity name='msdyn_workorder'>
                        <attribute name='hil_fiscalmonth' />
                        <attribute name='msdyn_name' />
                        <attribute name='createdon' />
                        <attribute name='hil_owneraccount' />
                        <attribute name='hil_channelpartnercategory' />
                        <attribute name='hil_warrantystatus' />
                        <attribute name='hil_countryclassification' />
                        <attribute name='hil_tatachievementslab' />
                        <attribute name='hil_tatcategory' />
                        <attribute name='hil_schemecode' />
                        <attribute name='hil_salesoffice' />
                        <attribute name='hil_productcategory' />
                        <attribute name='hil_kkgcode' />
                        <attribute name='hil_typeofassignee' />
                        <attribute name='hil_brand' />
                        <attribute name='hil_productsubcategory' />
                        <attribute name='hil_callsubtype' />
                        <attribute name='hil_claimstatus' />
                        <attribute name='hil_isocr' />
                        <attribute name='hil_laborinwarranty' />
                        <attribute name='hil_isgascharged' />
                        <attribute name='hil_sparepartuse' />
                        <attribute name='hil_jobclosemobile' />
                        <attribute name='msdyn_timeclosed' />
                        <attribute name='msdyn_customerasset' />
                        <attribute name='hil_quantity' />
                        <attribute name='hil_purchasedate' />
                        <attribute name='hil_sourceofjob' />
                        <attribute name='hil_pincode' />
                        <attribute name='msdyn_substatus' />
                        <filter type='and'>
                            <condition attribute='msdyn_workorderid' operator='eq'  value='" + entASP.GetAttributeValue<EntityReference>("hil_jobid").Id + @"' />
                            <condition attribute='hil_owneraccount' operator='not-null' />
                            <condition attribute='hil_salesoffice' operator='not-null' />
                            <condition attribute='hil_callsubtype' operator='not-null' />
                            <condition attribute='msdyn_customerasset' operator='not-null' />
                            <condition attribute='hil_productsubcategory' operator='not-null' />
                            <condition attribute='hil_productcategory' operator='not-null' />
                            <condition attribute='hil_tatcategory' operator='not-null' />
                            <condition attribute='hil_typeofassignee' operator='in'>
                                <value uiname='Franchise' uitype='position'>{4A1AA189-1208-E911-A94D-000D3AF0694E}</value>
                                <value uiname='Franchise Technician' uitype='position'>{0197EA9B-1208-E911-A94D-000D3AF0694E}</value>
                            </condition>
                        </filter>
                        <link-entity name='account' from='accountid' to='hil_owneraccount' visible='false' link-type='outer' alias='custCatg'>
                        <attribute name='hil_category' />
                        </link-entity>
                        <link-entity name='hil_address' from='hil_addressid' to='hil_address' link-type='outer' alias='DistJob'>
                            <attribute name='hil_district' />
                        </link-entity>
                        <link-entity name='account' from='accountid' to='hil_owneraccount' link-type='outer' alias='DistASP'>
                            <attribute name='hil_district' />
                        </link-entity>
                    </entity>
                    </fetch>";
                    try
                    {
                        entcollJobs = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                        
                        
                        foreach (Entity entJob in entcollJobs.Entities)
                        {
                            _claimStatus = RefreshClaimParameters(_service, _erFiscalMonth.Id, entJob);
                            _performaInvoiceId = Guid.Empty;
                            _fixedCharges = 0;
                            _erFranchise = entJob.GetAttributeValue<EntityReference>("hil_owneraccount");
                            _performaInvoiceId = GeneratePerformaInvoice(_service, _erFranchise.Id, _erFiscalMonth.Id);

                            Console.WriteLine("Job # " + entJob.GetAttributeValue<string>("msdyn_name"));
                            #region Fixed Charge Calculation
                            _fixedCharges = CreateFixedChargeClaimLine(_service, _performaInvoiceId, ClaimCategoryConst._FixedCompensation, _erFiscalMonth, _erFranchise, Convert.ToDateTime(_endDate));

                            if (_performaInvoiceId != null && (_claimStatus.Value == 4 || _claimStatus.Value == 7))
                            {
                                GenerateClaimLines(entJob, _performaInvoiceId, _service, _erFiscalMonth);
                                Console.WriteLine("Claim Cline generated");
                            }
                            entUpdate = new Entity("msdyn_workorder", entJob.Id);
                            entUpdate["hil_generateclaim"] = true;
                            entUpdate["hil_claimamount"] = new Money(0);
                            _service.Update(entUpdate);
                            #endregion
                            Console.WriteLine(i++.ToString() + "/" + _totalClosedJobs.ToString());
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }
                #endregion
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            Console.WriteLine("Batch Ended " + DateTime.Now.ToString());
        }
        public static void GenerateClaimLines(IOrganizationService _service)
        {
            #region Variable declaration
            Guid _fiscalMonthId = Guid.Empty;
            DateTime? _startDate = null;
            DateTime? _endDate = null;
            string _strStartDate = null;
            string _strEndDate = null;
            Guid _performaInvoiceId = Guid.Empty;
            int _totalClosedJobs = 0;
            string _fetchXML = string.Empty;
            Decimal _fixedCharges = 0;
            QueryExpression queryExp;
            EntityCollection entcoll;
            EntityCollection entcollJobs;
            EntityCollection entcollTATSlabJobs;
            EntityReference _erFiscalMonth = null;
            EntityReference _erFranchise = null;
            OptionSetValue _claimStatus = null;
            string[] _jobColumns = { "hil_fiscalmonth", "msdyn_name", "createdon", "hil_owneraccount", "hil_channelpartnercategory", "hil_warrantystatus", "hil_countryclassification", "hil_tatachievementslab", "hil_tatcategory", "hil_schemecode", "hil_salesoffice", "hil_productcategory", "hil_kkgcode", "hil_typeofassignee", "hil_brand", "hil_productsubcategory", "hil_callsubtype", "hil_claimstatus", "hil_isocr", "hil_laborinwarranty", "hil_isgascharged", "hil_sparepartuse", "hil_jobclosemobile", "msdyn_timeclosed", "msdyn_customerasset", "hil_quantity", "hil_purchasedate" };
            string _performaRemarks = string.Empty;
            Entity entUpdate;
            DateTime _closedOn;
            DateTime _asOnDate;
            #endregion

            #region Fetching Fiscal Month Date Range
            //DateTime _processDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);
            //DateTime _processDate = new DateTime(2022, 11, 20);

            //var fromDate = ConfigurationManager.AppSettings["fromDate"].ToString();
            //var toDate = ConfigurationManager.AppSettings["toDate"].ToString();

            var _salesOfficeId = ConfigurationManager.AppSettings["fromDate"].ToString();
            DateTime _processDate = Convert.ToDateTime(_salesOfficeId);

            queryExp = new QueryExpression("hil_claimperiod");
            queryExp.ColumnSet = new ColumnSet("hil_fromdate", "hil_todate");
            queryExp.Criteria = new FilterExpression(LogicalOperator.And);
            queryExp.Criteria.AddCondition("hil_fromdate", ConditionOperator.OnOrBefore, _processDate);
            queryExp.Criteria.AddCondition("hil_todate", ConditionOperator.OnOrAfter, _processDate);
            entcoll = _service.RetrieveMultiple(queryExp);
            if (entcoll.Entities.Count > 0)
            {
                _startDate = entcoll.Entities[0].GetAttributeValue<DateTime>("hil_fromdate").AddMinutes(330);
                _strStartDate = _startDate.Value.Year.ToString() + "-" + _startDate.Value.Month.ToString().PadLeft(2, '0') + "-" + _startDate.Value.Day.ToString().PadLeft(2, '0');
                _endDate = entcoll.Entities[0].GetAttributeValue<DateTime>("hil_todate").AddMinutes(330);
                _strEndDate = _endDate.Value.Year.ToString() + "-" + _endDate.Value.Month.ToString().PadLeft(2, '0') + "-" + _endDate.Value.Day.ToString().PadLeft(2, '0');
                _erFiscalMonth = entcoll.Entities[0].ToEntityReference();
            }
            if (_erFiscalMonth == null)
            {
                Console.WriteLine("Fiscal Month is not defined." + DateTime.Now.ToString());
                return;
            }
            #endregion

            #region Refreshing Claim Parameters & Generating applicable Claim Lines
            try
            {
                //_strStartDate = "2023-10-21";
                //_strEndDate = "2023-11-20";

                _fetchXML = @"<fetch top='50' version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                        <entity name='msdyn_workorder'>
                        <attribute name='hil_fiscalmonth' />
                        <attribute name='msdyn_name' />
                        <attribute name='createdon' />
                        <attribute name='hil_owneraccount' />
                        <attribute name='hil_channelpartnercategory' />
                        <attribute name='hil_warrantystatus' />
                        <attribute name='hil_countryclassification' />
                        <attribute name='hil_tatachievementslab' />
                        <attribute name='hil_tatcategory' />
                        <attribute name='hil_schemecode' />
                        <attribute name='hil_salesoffice' />
                        <attribute name='hil_productcategory' />
                        <attribute name='hil_kkgcode' />
                        <attribute name='hil_typeofassignee' />
                        <attribute name='hil_brand' />
                        <attribute name='hil_productsubcategory' />
                        <attribute name='hil_callsubtype' />
                        <attribute name='hil_claimstatus' />
                        <attribute name='hil_isocr' />
                        <attribute name='hil_laborinwarranty' />
                        <attribute name='hil_isgascharged' />
                        <attribute name='hil_sparepartuse' />
                        <attribute name='hil_jobclosemobile' />
                        <attribute name='msdyn_timeclosed' />
                        <attribute name='msdyn_customerasset' />
                        <attribute name='hil_quantity' />
                        <attribute name='hil_purchasedate' />
                        <attribute name='hil_sourceofjob' />
                        <attribute name='hil_pincode' />
                        <attribute name='hil_reviewforcountryclassification' />
                        <attribute name='hil_countryclassification' />
                        <order attribute='msdyn_timeclosed' descending='true' />
                        <filter type='and'>
                            <condition attribute='hil_fiscalmonth' operator='eq' value='" + _erFiscalMonth.Id + @"' />
                            <condition attribute='hil_generateclaim' operator='ne' value='1' />
                            <condition attribute='hil_claimstatus' operator='in'>
                            <value>4</value>
                            <value>7</value>
                            </condition>
                            <condition attribute='msdyn_timeclosed' operator='on-or-after' value='" + _strStartDate + @"' />
                            <condition attribute='msdyn_timeclosed' operator='on-or-before' value='" + _strEndDate + @"' />
                        </filter>
                        <link-entity name='account' from='accountid' to='hil_owneraccount' visible='false' link-type='outer' alias='custCatg'>
                            <attribute name='hil_category' />
                        </link-entity>
                    </entity>
                    </fetch>";
                try
                {
                    while (true)
                    {
                        entcollJobs = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                        _totalClosedJobs = entcollJobs.Entities.Count;
                        if (_totalClosedJobs == 0)
                        {
                            break;
                        }
                        int i = 0;
                        foreach (Entity entJob in entcollJobs.Entities)
                        {
                            Console.WriteLine("Job # " + entJob.GetAttributeValue<string>("msdyn_name"));
                            try
                            {
                                _closedOn = entJob.GetAttributeValue<DateTime>("msdyn_timeclosed").AddMinutes(330);
                                _asOnDate = DateTime.Now.AddHours(-48);
                                if (_closedOn >= _asOnDate && entJob.GetAttributeValue<bool>("hil_reviewforcountryclassification") && entJob.GetAttributeValue<OptionSetValue>("hil_countryclassification").Value == 1)
                                {
                                    //Do Nothing
                                    //Console.WriteLine("Under Review for Country Classification" + i.ToString() + " / " + _totalClosedJobs.ToString());
                                }
                                else
                                {
                                    _performaInvoiceId = Guid.Empty;
                                    _fixedCharges = 0;
                                    _erFranchise = entJob.GetAttributeValue<EntityReference>("hil_owneraccount");
                                    _performaInvoiceId = GeneratePerformaInvoice(_service, _erFranchise.Id, _erFiscalMonth.Id);

                                    
                                    #region Fixed Charge Calculation
                                    _fixedCharges = CreateFixedChargeClaimLine(_service, _performaInvoiceId, ClaimCategoryConst._FixedCompensation, _erFiscalMonth, _erFranchise, Convert.ToDateTime(_endDate));
                                    _claimStatus = entJob.GetAttributeValue<OptionSetValue>("hil_claimstatus");

                                    if (_performaInvoiceId != null && (_claimStatus.Value == 4 || _claimStatus.Value == 7))
                                    {
                                        GenerateClaimLines(entJob, _performaInvoiceId, _service, _erFiscalMonth);
                                        Console.WriteLine("Claim Cline generated");
                                    }
                                    entUpdate = new Entity("msdyn_workorder", entJob.Id);
                                    entUpdate["hil_generateclaim"] = true;
                                    entUpdate["hil_claimamount"] = new Money(0);
                                    _service.Update(entUpdate);
                                    #endregion
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(i++.ToString() + "/" + _totalClosedJobs.ToString() + "::ERROR !!! " + ex.Message);
                            }
                            Console.WriteLine(i++.ToString() + "/" + _totalClosedJobs.ToString());
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
                //}
                #endregion
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        public static void GenerateClaimLinesTest(IOrganizationService _service)
        {
            #region Variable declaration
            Guid _fiscalMonthId = Guid.Empty;
            DateTime? _startDate = null;
            DateTime? _endDate = null;
            string _strStartDate = null;
            string _strEndDate = null;
            Guid _performaInvoiceId = Guid.Empty;
            int _totalClosedJobs = 0;
            string _fetchXML = string.Empty;
            Decimal _fixedCharges = 0;
            QueryExpression queryExp;
            EntityCollection entcoll;
            EntityCollection entcollJobs;
            EntityCollection entcollTATSlabJobs;
            EntityReference _erFiscalMonth = null;
            EntityReference _erFranchise = null;
            OptionSetValue _claimStatus = null;
            string[] _jobColumns = { "hil_fiscalmonth", "msdyn_name", "createdon", "hil_owneraccount", "hil_channelpartnercategory", "hil_warrantystatus", "hil_countryclassification", "hil_tatachievementslab", "hil_tatcategory", "hil_schemecode", "hil_salesoffice", "hil_productcategory", "hil_kkgcode", "hil_typeofassignee", "hil_brand", "hil_productsubcategory", "hil_callsubtype", "hil_claimstatus", "hil_isocr", "hil_laborinwarranty", "hil_isgascharged", "hil_sparepartuse", "hil_jobclosemobile", "msdyn_timeclosed", "msdyn_customerasset", "hil_quantity", "hil_purchasedate" };
            string _performaRemarks = string.Empty;
            Entity entUpdate;
            DateTime _closedOn;
            DateTime _asOnDate;
            #endregion

            #region Fetching Fiscal Month Date Range
            //DateTime _processDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);
            DateTime _processDate = new DateTime(2022, 11, 20);

            //var fromDate = ConfigurationManager.AppSettings["fromDate"].ToString();
            //var toDate = ConfigurationManager.AppSettings["toDate"].ToString();

            queryExp = new QueryExpression("hil_claimperiod");
            queryExp.ColumnSet = new ColumnSet("hil_fromdate", "hil_todate");
            queryExp.Criteria = new FilterExpression(LogicalOperator.And);
            queryExp.Criteria.AddCondition("hil_fromdate", ConditionOperator.OnOrBefore, _processDate);
            queryExp.Criteria.AddCondition("hil_todate", ConditionOperator.OnOrAfter, _processDate);
            entcoll = _service.RetrieveMultiple(queryExp);
            if (entcoll.Entities.Count > 0)
            {
                _startDate = entcoll.Entities[0].GetAttributeValue<DateTime>("hil_fromdate").AddMinutes(330);
                _strStartDate = _startDate.Value.Year.ToString() + "-" + _startDate.Value.Month.ToString().PadLeft(2, '0') + "-" + _startDate.Value.Day.ToString().PadLeft(2, '0');
                _endDate = entcoll.Entities[0].GetAttributeValue<DateTime>("hil_todate").AddMinutes(330);
                _strEndDate = _endDate.Value.Year.ToString() + "-" + _endDate.Value.Month.ToString().PadLeft(2, '0') + "-" + _endDate.Value.Day.ToString().PadLeft(2, '0');
                _erFiscalMonth = entcoll.Entities[0].ToEntityReference();
            }
            if (_erFiscalMonth == null)
            {
                Console.WriteLine("Fiscal Month is not defined." + DateTime.Now.ToString());
                return;
            }
            #endregion

            #region Refreshing Claim Parameters & Generating applicable Claim Lines
            try
            {
                _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                        <entity name='msdyn_workorder'>
                        <attribute name='hil_fiscalmonth' />
                        <attribute name='msdyn_name' />
                        <attribute name='createdon' />
                        <attribute name='hil_owneraccount' />
                        <attribute name='hil_channelpartnercategory' />
                        <attribute name='hil_warrantystatus' />
                        <attribute name='hil_countryclassification' />
                        <attribute name='hil_tatachievementslab' />
                        <attribute name='hil_tatcategory' />
                        <attribute name='hil_schemecode' />
                        <attribute name='hil_salesoffice' />
                        <attribute name='hil_productcategory' />
                        <attribute name='hil_kkgcode' />
                        <attribute name='hil_typeofassignee' />
                        <attribute name='hil_brand' />
                        <attribute name='hil_productsubcategory' />
                        <attribute name='hil_callsubtype' />
                        <attribute name='hil_claimstatus' />
                        <attribute name='hil_isocr' />
                        <attribute name='hil_laborinwarranty' />
                        <attribute name='hil_isgascharged' />
                        <attribute name='hil_sparepartuse' />
                        <attribute name='hil_jobclosemobile' />
                        <attribute name='msdyn_timeclosed' />
                        <attribute name='msdyn_customerasset' />
                        <attribute name='hil_quantity' />
                        <attribute name='hil_purchasedate' />
                        <attribute name='hil_sourceofjob' />
                        <attribute name='hil_pincode' />
                        <attribute name='hil_reviewforcountryclassification' />
                        <attribute name='hil_countryclassification' />
                        <order attribute='msdyn_timeclosed' descending='false' />
                        <filter type='and'>
                            <condition attribute='hil_generateclaim' operator='ne' value='1' />
                            <condition attribute='hil_claimstatus' operator='in'>
                                <value>4</value>
                                <value>7</value>
                            </condition>
                            <condition attribute='hil_fiscalmonth' operator='eq' value='" + _erFiscalMonth.Id + @"' />
                            <condition attribute='msdyn_timeclosed' operator='on-or-after' value='" + _strStartDate + @"' />
                            <condition attribute='msdyn_timeclosed' operator='on-or-before' value='" + _strEndDate + @"' />
                        </filter>
                        <link-entity name='account' from='accountid' to='hil_owneraccount' visible='false' link-type='outer' alias='custCatg'>
                            <attribute name='hil_category' />
                        </link-entity>
                    </entity>
                    </fetch>";
                
                try
                {
                    while (true)
                    {
                        entcollJobs = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                        _totalClosedJobs = entcollJobs.Entities.Count;
                        if (_totalClosedJobs == 0)
                        {
                            break;
                        }
                        int i = 0;
                        foreach (Entity entJob in entcollJobs.Entities)
                        {
                            Console.WriteLine("Job # " + entJob.GetAttributeValue<string>("msdyn_name"));

                            try
                            {
                                _closedOn = entJob.GetAttributeValue<DateTime>("msdyn_timeclosed").AddMinutes(330);
                                _asOnDate = DateTime.Now.AddHours(-48);
                                if (_closedOn >= _asOnDate && entJob.GetAttributeValue<bool>("hil_reviewforcountryclassification") && entJob.GetAttributeValue<OptionSetValue>("hil_countryclassification").Value == 1)
                                {
                                    // Do Nothing
                                    Console.WriteLine("Under Review for Country Classification" + i.ToString() + " / " + _totalClosedJobs.ToString());
                                }
                                else
                                {
                                    _performaInvoiceId = Guid.Empty;
                                    _fixedCharges = 0;
                                    _erFranchise = entJob.GetAttributeValue<EntityReference>("hil_owneraccount");
                                    _performaInvoiceId = GeneratePerformaInvoice(_service, _erFranchise.Id, _erFiscalMonth.Id);

                                    
                                    #region Fixed Charge Calculation
                                    _fixedCharges = CreateFixedChargeClaimLine(_service, _performaInvoiceId, ClaimCategoryConst._FixedCompensation, _erFiscalMonth, _erFranchise, Convert.ToDateTime(_endDate));
                                    _claimStatus = entJob.GetAttributeValue<OptionSetValue>("hil_claimstatus");

                                    if (_performaInvoiceId != null && (_claimStatus.Value == 4 || _claimStatus.Value == 7))
                                    {
                                        GenerateClaimLines(entJob, _performaInvoiceId, _service, _erFiscalMonth);
                                        Console.WriteLine("Claim Cline generated");
                                    }
                                    entUpdate = new Entity("msdyn_workorder", entJob.Id);
                                    entUpdate["hil_generateclaim"] = true;
                                    entUpdate["hil_claimamount"] = new Money(0);
                                    _service.Update(entUpdate);
                                    #endregion
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(i++.ToString() + "/" + _totalClosedJobs.ToString() + "::ERROR !!! " + ex.Message);
                            }
                            Console.WriteLine(i++.ToString() + "/" + _totalClosedJobs.ToString());
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
                //}
                #endregion
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        public static void RejectUnderReviewSAW(IOrganizationService _service)
        {
            #region Variable declaration
            Guid _fiscalMonthId = Guid.Empty;
            DateTime? _startDate = null;
            DateTime? _endDate = null;
            string _strStartDate = null;
            string _strEndDate = null;
            Guid _performaInvoiceId = Guid.Empty;
            int _totalClosedJobs = 0;
            string _fetchXML = string.Empty;
            QueryExpression queryExp;
            EntityCollection entcoll;
            EntityCollection entcollJobs;
            EntityReference _erFiscalMonth = null;
            string[] _jobColumns = { "hil_fiscalmonth", "msdyn_name", "createdon", "hil_owneraccount", "hil_channelpartnercategory", "hil_warrantystatus", "hil_countryclassification", "hil_tatachievementslab", "hil_tatcategory", "hil_schemecode", "hil_salesoffice", "hil_productcategory", "hil_kkgcode", "hil_typeofassignee", "hil_brand", "hil_productsubcategory", "hil_callsubtype", "hil_claimstatus", "hil_isocr", "hil_laborinwarranty", "hil_isgascharged", "hil_sparepartuse", "hil_jobclosemobile", "msdyn_timeclosed", "msdyn_customerasset", "hil_quantity", "hil_purchasedate" };
            string _performaRemarks = string.Empty;
            #endregion

            #region Fetching Fiscal Month Date Range
            queryExp = new QueryExpression("hil_claimperiod");
            queryExp.ColumnSet = new ColumnSet("hil_fromdate", "hil_todate");
            queryExp.Criteria = new FilterExpression(LogicalOperator.And);
            queryExp.Criteria.AddCondition("hil_isapplicable", ConditionOperator.Equal, true);
            queryExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
            entcoll = _service.RetrieveMultiple(queryExp);
            if (entcoll.Entities.Count > 0)
            {
                _startDate = entcoll.Entities[0].GetAttributeValue<DateTime>("hil_fromdate").AddMinutes(330);
                _strStartDate = _startDate.Value.Year.ToString() + "-" + _startDate.Value.Month.ToString().PadLeft(2, '0') + "-" + _startDate.Value.Day.ToString().PadLeft(2, '0');
                _endDate = entcoll.Entities[0].GetAttributeValue<DateTime>("hil_todate").AddMinutes(330);
                _strEndDate = _endDate.Value.Year.ToString() + "-" + _endDate.Value.Month.ToString().PadLeft(2, '0') + "-" + _endDate.Value.Day.ToString().PadLeft(2, '0');
                _erFiscalMonth = entcoll.Entities[0].ToEntityReference();
            }
            if (_erFiscalMonth == null)
            {
                Console.WriteLine("Fiscal Month is not defined." + DateTime.Now.ToString());
                return;
            }
            #endregion

            #region Refreshing Claim Parameters & Generating applicable Claim Lines
            try
            {
                _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                  <entity name='hil_sawactivityapproval'>
                    <attribute name='hil_approver' />
                    <attribute name='hil_approvalstatus' />
                    <order attribute='createdon' descending='false' />
                    <filter type='and'>
                      <condition attribute='hil_approvalstatus' operator='in'>
                        <value>1</value>
                        <value>2</value>
                        <value>5</value>
                        <value>6</value>
                      </condition>
                    </filter>
                    <link-entity name='msdyn_workorder' from='msdyn_workorderid' to='hil_jobid' link-type='inner' alias='a_eecd118653ddea11a813000d3af0563c'>
                      <filter type='and'>
                        <condition attribute='hil_fiscalmonth' operator='eq' uiname='202204' uitype='hil_claimperiod' value='" + _erFiscalMonth.Id + @"' />
                      </filter>
                    </link-entity>
                  </entity>
                </fetch>";

                try
                {
                    while (true)
                    {
                        entcollJobs = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                        _totalClosedJobs = entcollJobs.Entities.Count;
                        if (_totalClosedJobs == 0)
                        {
                            break;
                        }
                        int i = 1;
                        foreach (Entity entJob in entcollJobs.Entities)
                        {
                            entJob["hil_approver"] = new EntityReference("systemuser",new Guid("08074320-fcee-e811-a949-000d3af03089"));
                            _service.Update(entJob);

                            entJob["hil_approvalstatus"] = new OptionSetValue(4); // Rejected
                            _service.Update(entJob);

                            Console.WriteLine(i++.ToString() + "/" + _totalClosedJobs.ToString());
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
                #endregion
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        public static void GenerateClaimLinesForJobFixedCompensation(IOrganizationService _service)
        {
            #region Variable declaration
            Guid _fiscalMonthId = Guid.Empty;
            DateTime? _startDate = null;
            DateTime? _endDate = null;
            string _strStartDate = null;
            string _strEndDate = null;
            Guid _performaInvoiceId = Guid.Empty;
            int _totalClosedJobs = 0;
            string _fetchXML = string.Empty;
            Decimal _fixedCharges = 0;
            QueryExpression queryExp;
            EntityCollection entcoll;
            EntityCollection entcollJobs;
            EntityCollection entcollTATSlabJobs;
            EntityReference _erFiscalMonth = null;
            EntityReference _erFranchise = null;
            OptionSetValue _claimStatus = null;
            string[] _jobColumns = { "hil_fiscalmonth", "msdyn_name", "createdon", "hil_owneraccount", "hil_channelpartnercategory", "hil_warrantystatus", "hil_countryclassification", "hil_tatachievementslab", "hil_tatcategory", "hil_schemecode", "hil_salesoffice", "hil_productcategory", "hil_kkgcode", "hil_typeofassignee", "hil_brand", "hil_productsubcategory", "hil_callsubtype", "hil_claimstatus", "hil_isocr", "hil_laborinwarranty", "hil_isgascharged", "hil_sparepartuse", "hil_jobclosemobile", "msdyn_timeclosed", "msdyn_customerasset", "hil_quantity", "hil_purchasedate" };
            string _performaRemarks = string.Empty;
            Entity entUpdate;
            DateTime _closedOn;
            DateTime _asOnDate;
            #endregion

            #region Fetching Fiscal Month Date Range
            //DateTime _processDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);
            DateTime _processDate = new DateTime(2021, 08, 20);
            queryExp = new QueryExpression("hil_claimperiod");
            queryExp.ColumnSet = new ColumnSet("hil_fromdate", "hil_todate");
            queryExp.Criteria = new FilterExpression(LogicalOperator.And);
            queryExp.Criteria.AddCondition("hil_fromdate", ConditionOperator.OnOrBefore, _processDate);
            queryExp.Criteria.AddCondition("hil_todate", ConditionOperator.OnOrAfter, _processDate);
            entcoll = _service.RetrieveMultiple(queryExp);
            if (entcoll.Entities.Count > 0)
            {
                _startDate = entcoll.Entities[0].GetAttributeValue<DateTime>("hil_fromdate").AddMinutes(330);
                _strStartDate = _startDate.Value.Year.ToString() + "-" + _startDate.Value.Month.ToString().PadLeft(2, '0') + "-" + _startDate.Value.Day.ToString().PadLeft(2, '0');
                _endDate = entcoll.Entities[0].GetAttributeValue<DateTime>("hil_todate").AddMinutes(330);
                _strEndDate = _endDate.Value.Year.ToString() + "-" + _endDate.Value.Month.ToString().PadLeft(2, '0') + "-" + _endDate.Value.Day.ToString().PadLeft(2, '0');
                _erFiscalMonth = entcoll.Entities[0].ToEntityReference();
            }
            if (_erFiscalMonth == null)
            {
                Console.WriteLine("Fiscal Month is not defined." + DateTime.Now.ToString());
                return;
            }
            #endregion

            #region Refreshing Claim Parameters & Generating applicable Claim Lines
            try
            {

                _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                              <entity name='msdyn_workorder'>
                                <attribute name='msdyn_name' />
                                <attribute name='createdon' />
                                <attribute name='hil_productsubcategory' />
                                <attribute name='hil_owneraccount' />
                                <attribute name='hil_callsubtype' />
                                <attribute name='msdyn_workorderid' />
                                <order attribute='msdyn_name' descending='false' />
                                <filter type='and'>
                                  <condition attribute='msdyn_workorderid' operator='in'>
                                    <value uiname='01082110475218' uitype='msdyn_workorder'>{646A3503-93F2-EB11-94EF-000D3A3E3121}</value>
                                    <value uiname='31072110471971' uitype='msdyn_workorder'>{7B356292-08F2-EB11-94EF-000D3A3E22AA}</value>
                                    <value uiname='27072110417444' uitype='msdyn_workorder'>{DDD0C472-08EF-EB11-94EF-000D3A3E22AA}</value>
                                    <value uiname='01082110478021' uitype='msdyn_workorder'>{D48BA234-AFF2-EB11-94EF-000D3A3E39D6}</value>
                                    <value uiname='18082110706150' uitype='msdyn_workorder'>{1ADF3D66-0E00-EC11-94EF-6045BD72E3F7}</value>
                                    <value uiname='12082110623745' uitype='msdyn_workorder'>{C4BA2E8E-34FB-EB11-94EF-6045BD72EA7B}</value>
                                    <value uiname='18082110699439' uitype='msdyn_workorder'>{4901E1AF-E8FF-EB11-94EF-6045BD72E02A}</value>
                                    <value uiname='16082110673610' uitype='msdyn_workorder'>{D5E92D54-6DFE-EB11-94EF-6045BD72E5B7}</value>
                                    <value uiname='03082110510269' uitype='msdyn_workorder'>{1355DE30-4DF4-EB11-94EF-000D3A3E3121}</value>
                                    <value uiname='18082110697091' uitype='msdyn_workorder'>{0D468533-DDFF-EB11-94EF-6045BD72EAD7}</value>
                                    <value uiname='12082110631490' uitype='msdyn_workorder'>{E20CABCB-5EFB-EB11-94EF-6045BD72E14A}</value>
                                    <value uiname='16082110676958' uitype='msdyn_workorder'>{51452C8C-7CFE-EB11-94EF-6045BD72E516}</value>
                                    <value uiname='10082110593204' uitype='msdyn_workorder'>{8A733948-96F9-EB11-94EF-6045BD72E14A}</value>
                                    <value uiname='17082110694284' uitype='msdyn_workorder'>{D1C4B2FF-53FF-EB11-94EF-6045BD72E290}</value>
                                    <value uiname='17082110685276' uitype='msdyn_workorder'>{B771D3DE-22FF-EB11-94EF-6045BD72E5B7}</value>
                                    <value uiname='18082110699565' uitype='msdyn_workorder'>{88E10C65-E9FF-EB11-94EF-6045BD72EA7B}</value>
                                    <value uiname='16082110677107' uitype='msdyn_workorder'>{6E7E4273-7DFE-EB11-94EF-6045BD72E14A}</value>
                                    <value uiname='11082110615922' uitype='msdyn_workorder'>{B064F5C3-8DFA-EB11-94EF-6045BD72E5B7}</value>
                                    <value uiname='16082110677976' uitype='msdyn_workorder'>{96D10EB7-82FE-EB11-94EF-6045BD72E82D}</value>
                                    <value uiname='18082110706792' uitype='msdyn_workorder'>{AA592E11-1200-EC11-94EF-6045BD72E14A}</value>
                                    <value uiname='02082110491060' uitype='msdyn_workorder'>{0461245B-79F3-EB11-94EF-000D3A3E3C39}</value>
                                    <value uiname='10082110596878' uitype='msdyn_workorder'>{5EBE06B2-A8F9-EB11-94EF-6045BD72E14A}</value>
                                    <value uiname='07082110559914' uitype='msdyn_workorder'>{622B8ECF-48F7-EB11-94EF-000D3A3E326A}</value>
                                    <value uiname='17082110689835' uitype='msdyn_workorder'>{E71D584B-39FF-EB11-94EF-6045BD72E02A}</value>
                                    <value uiname='16082110677925' uitype='msdyn_workorder'>{6D43516D-82FE-EB11-94EF-6045BD72E5B7}</value>
                                    <value uiname='09082110580611' uitype='msdyn_workorder'>{5E3CE0A0-DBF8-EB11-94EF-6045BD72E290}</value>
                                    <value uiname='18082110703330' uitype='msdyn_workorder'>{D009D592-FCFF-EB11-94EF-6045BD72E02A}</value>
                                    <value uiname='18082110703018' uitype='msdyn_workorder'>{B78452BF-FAFF-EB11-94EF-6045BD72E82D}</value>
                                    <value uiname='16082110677355' uitype='msdyn_workorder'>{4740F816-7FFE-EB11-94EF-6045BD72E14A}</value>
                                    <value uiname='09082110578619' uitype='msdyn_workorder'>{2E3E5ADB-D2F8-EB11-94EF-6045BD72E5B7}</value>
                                    <value uiname='02082110489322' uitype='msdyn_workorder'>{FD199F0E-70F3-EB11-94EF-000D3A3E3953}</value>
                                    <value uiname='16082110666673' uitype='msdyn_workorder'>{AF8319C3-4FFE-EB11-94EF-6045BD72E5B7}</value>
                                    <value uiname='17082110695354' uitype='msdyn_workorder'>{5BB3CA44-71FF-EB11-94EF-6045BD72E5B7}</value>
                                    <value uiname='17082110682077' uitype='msdyn_workorder'>{E9F0095D-13FF-EB11-94EF-6045BD72E14A}</value>
                                    <value uiname='24072110370424' uitype='msdyn_workorder'>{991B3BD3-50EC-EB11-BACB-000D3AF0CB64}</value>
                                    <value uiname='16082110663881' uitype='msdyn_workorder'>{56422717-41FE-EB11-94EF-6045BD72E7EE}</value>
                                    <value uiname='11082110609337' uitype='msdyn_workorder'>{D0A7C440-6AFA-EB11-94EF-6045BD72E490}</value>
                                    <value uiname='12082110629301' uitype='msdyn_workorder'>{62583A80-51FB-EB11-94EF-6045BD72E82D}</value>
                                    <value uiname='12082110632493' uitype='msdyn_workorder'>{7F106F43-65FB-EB11-94EF-6045BD72E5B7}</value>
                                    <value uiname='18082110707145' uitype='msdyn_workorder'>{B91CAD14-1400-EC11-94EF-6045BD72E02A}</value>
                                    <value uiname='18082110708119' uitype='msdyn_workorder'>{FF1E15B1-1900-EC11-94EF-6045BD72EAD7}</value>
                                    <value uiname='17082110692879' uitype='msdyn_workorder'>{67BFF0CC-4BFF-EB11-94EF-6045BD72E290}</value>
                                    <value uiname='06082110542679' uitype='msdyn_workorder'>{B7E19BC7-6CF6-EB11-94EF-000D3A3E39D6}</value>
                                    <value uiname='13082110646962' uitype='msdyn_workorder'>{16D0602C-37FC-EB11-94EF-6045BD72E5B7}</value>
                                    <value uiname='31072110470189' uitype='msdyn_workorder'>{3E1D51E5-F0F1-EB11-94EF-000D3A3E58E8}</value>
                                    <value uiname='16082110675412' uitype='msdyn_workorder'>{CA654A91-73FE-EB11-94EE-6045BD72A5A6}</value>
                                    <value uiname='18082110697924' uitype='msdyn_workorder'>{ADE9D119-E1FF-EB11-94EF-6045BD72E02A}</value>
                                    <value uiname='16082110673535' uitype='msdyn_workorder'>{0947E4E4-6CFE-EB11-94EF-6045BD72EA7B}</value>
                                    <value uiname='16082110672859' uitype='msdyn_workorder'>{39D1E238-69FE-EB11-94EF-6045BD72E516}</value>
                                    <value uiname='21072110326216' uitype='msdyn_workorder'>{E0E3092C-DFE9-EB11-BACB-000D3AF0C443}</value>
                                    <value uiname='07072110130293' uitype='msdyn_workorder'>{9360D76C-FFDE-EB11-BACB-0022486EA35C}</value>
                                    <value uiname='02082110491022' uitype='msdyn_workorder'>{51F7D918-79F3-EB11-94EF-000D3A3E3D78}</value>
                                    <value uiname='10072110174761' uitype='msdyn_workorder'>{74300884-41E1-EB11-BACB-0022486EA016}</value>
                                    <value uiname='06082110548158' uitype='msdyn_workorder'>{E62BF4F4-8AF6-EB11-94EF-000D3A3E3F77}</value>
                                    <value uiname='03082110508585' uitype='msdyn_workorder'>{0C2DED92-42F4-EB11-94EF-000D3A3E3121}</value>
                                    <value uiname='24072110375251' uitype='msdyn_workorder'>{511F3F51-6FEC-EB11-BACB-000D3AF0CD16}</value>
                                    <value uiname='28072110429981' uitype='msdyn_workorder'>{5A5DC2BF-99EF-EB11-94EF-000D3A3E1ED7}</value>
                                    <value uiname='19072110307859' uitype='msdyn_workorder'>{D86CC072-85E8-EB11-BACB-000D3AF0B5D7}</value>
                                    <value uiname='01082110477614' uitype='msdyn_workorder'>{54BE8EE3-A9F2-EB11-94EF-000D3A3E326A}</value>
                                    <value uiname='24072110373648' uitype='msdyn_workorder'>{DC5907D9-63EC-EB11-BACB-000D3AF0CA07}</value>
                                    <value uiname='05082110534163' uitype='msdyn_workorder'>{44B47AE6-C0F5-EB11-94EF-000D3A3E3D78}</value>
                                  </condition>
                                </filter>
                              </entity>
                            </fetch>";
                    try
                    {
                        entcollJobs = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                        _totalClosedJobs = entcollJobs.Entities.Count;
                        int i = 0;
                        foreach (Entity entJob in entcollJobs.Entities)
                        {
                            _performaInvoiceId = Guid.Empty;
                            _fixedCharges = 0;
                            _erFranchise = entJob.GetAttributeValue<EntityReference>("hil_owneraccount");
                            _performaInvoiceId = GeneratePerformaInvoice(_service, _erFranchise.Id, _erFiscalMonth.Id);
                            #region Fixed Charge Calculation
                            _fixedCharges = CreateFixedChargeClaimLine(_service, _performaInvoiceId, ClaimCategoryConst._FixedCompensation, _erFiscalMonth, _erFranchise, Convert.ToDateTime(_endDate));
                            #endregion
                            Console.WriteLine(i++.ToString() + "/" + _totalClosedJobs.ToString());
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                #endregion
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        public static void GenerateClaimLinesForJob(IOrganizationService _service)
        {
            #region Variable declaration
            Guid _fiscalMonthId = Guid.Empty;
            DateTime? _startDate = null;
            DateTime? _endDate = null;
            string _strStartDate = null;
            string _strEndDate = null;
            Guid _performaInvoiceId = Guid.Empty;
            int _totalClosedJobs = 0;
            string _fetchXML = string.Empty;
            Decimal _fixedCharges = 0;
            QueryExpression queryExp;
            EntityCollection entcoll;
            EntityCollection entcollJobs;
            EntityCollection entcollTATSlabJobs;
            EntityReference _erFiscalMonth = null;
            EntityReference _erFranchise = null;
            OptionSetValue _claimStatus = null;
            string[] _jobColumns = { "hil_fiscalmonth", "msdyn_name", "createdon", "hil_owneraccount", "hil_channelpartnercategory", "hil_warrantystatus", "hil_countryclassification", "hil_tatachievementslab", "hil_tatcategory", "hil_schemecode", "hil_salesoffice", "hil_productcategory", "hil_kkgcode", "hil_typeofassignee", "hil_brand", "hil_productsubcategory", "hil_callsubtype", "hil_claimstatus", "hil_isocr", "hil_laborinwarranty", "hil_isgascharged", "hil_sparepartuse", "hil_jobclosemobile", "msdyn_timeclosed", "msdyn_customerasset", "hil_quantity", "hil_purchasedate" };
            string _performaRemarks = string.Empty;
            Entity entUpdate;
            DateTime _closedOn;
            DateTime _asOnDate;
            #endregion

            #region Fetching Fiscal Month Date Range
            //DateTime _processDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);
            DateTime _processDate = new DateTime(2022, 03, 31);
            queryExp = new QueryExpression("hil_claimperiod");
            queryExp.ColumnSet = new ColumnSet("hil_fromdate", "hil_todate");
            queryExp.Criteria = new FilterExpression(LogicalOperator.And);
            queryExp.Criteria.AddCondition("hil_fromdate", ConditionOperator.OnOrBefore, _processDate);
            queryExp.Criteria.AddCondition("hil_todate", ConditionOperator.OnOrAfter, _processDate);
            entcoll = _service.RetrieveMultiple(queryExp);
            if (entcoll.Entities.Count > 0)
            {
                _startDate = entcoll.Entities[0].GetAttributeValue<DateTime>("hil_fromdate").AddMinutes(330);
                _strStartDate = _startDate.Value.Year.ToString() + "-" + _startDate.Value.Month.ToString().PadLeft(2, '0') + "-" + _startDate.Value.Day.ToString().PadLeft(2, '0');
                _endDate = entcoll.Entities[0].GetAttributeValue<DateTime>("hil_todate").AddMinutes(330);
                _strEndDate = _endDate.Value.Year.ToString() + "-" + _endDate.Value.Month.ToString().PadLeft(2, '0') + "-" + _endDate.Value.Day.ToString().PadLeft(2, '0');
                _erFiscalMonth = entcoll.Entities[0].ToEntityReference();
            }
            if (_erFiscalMonth == null)
            {
                Console.WriteLine("Fiscal Month is not defined." + DateTime.Now.ToString());
                return;
            }
            #endregion

            #region Refreshing Claim Parameters & Generating applicable Claim Lines
            try
            {
                /*
                 <condition attribute='hil_reviewforcountryclassification' operator='eq' value='1' />
                 <condition attribute='hil_countryclassification' operator='eq' value='1' />
                <condition attribute='msdyn_name' operator='eq' value='20112111948898' />
                 */

                _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                <entity name='msdyn_workorder'>
                <order attribute='msdyn_timeclosed' descending='true' />
                <attribute name='hil_owneraccount' />
                    <filter type='and'>
                    <condition attribute='hil_fiscalmonth' operator='eq' value='" + _erFiscalMonth.Id + @"' />
                    <condition attribute='hil_generateclaim' operator='ne' value='1' />
                    <condition attribute='hil_claimstatus' operator='in'>
                    <value>4</value>
                    <value>7</value>
                    </condition>
                    
                    </filter>
                </entity>
                </fetch>";

                entcoll = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                foreach (Entity entASP in entcoll.Entities)
                {
                    if (entASP.GetAttributeValue<EntityReference>("hil_owneraccount") != null)
                    {
                        _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                        <entity name='msdyn_workorder'>
                        <attribute name='hil_fiscalmonth' />
                        <attribute name='msdyn_name' />
                        <attribute name='createdon' />
                        <attribute name='hil_owneraccount' />
                        <attribute name='hil_channelpartnercategory' />
                        <attribute name='hil_warrantystatus' />
                        <attribute name='hil_countryclassification' />
                        <attribute name='hil_tatachievementslab' />
                        <attribute name='hil_tatcategory' />
                        <attribute name='hil_schemecode' />
                        <attribute name='hil_salesoffice' />
                        <attribute name='hil_productcategory' />
                        <attribute name='hil_kkgcode' />
                        <attribute name='hil_typeofassignee' />
                        <attribute name='hil_brand' />
                        <attribute name='hil_productsubcategory' />
                        <attribute name='hil_callsubtype' />
                        <attribute name='hil_claimstatus' />
                        <attribute name='hil_isocr' />
                        <attribute name='hil_laborinwarranty' />
                        <attribute name='hil_isgascharged' />
                        <attribute name='hil_sparepartuse' />
                        <attribute name='hil_jobclosemobile' />
                        <attribute name='msdyn_timeclosed' />
                        <attribute name='msdyn_customerasset' />
                        <attribute name='hil_quantity' />
                        <attribute name='hil_purchasedate' />
                        <attribute name='hil_sourceofjob' />
                        <attribute name='hil_pincode' />
                        <attribute name='hil_reviewforcountryclassification' />
                        <attribute name='hil_countryclassification' />
                        <order attribute='msdyn_timeclosed' descending='true' />
                        <filter type='and'>
                            <condition attribute='hil_owneraccount' operator='eq'  value='" + entASP.GetAttributeValue<EntityReference>("hil_owneraccount").Id + @"' />
                            <condition attribute='hil_fiscalmonth' operator='eq' value='" + _erFiscalMonth.Id + @"' />
                            <condition attribute='hil_claimstatus' operator='in'>
                            <value>4</value>
                            <value>7</value>
                            </condition>
                            <condition attribute='hil_generateclaim' operator='ne' value='1' />
                        </filter>
                        <link-entity name='account' from='accountid' to='hil_owneraccount' visible='false' link-type='outer' alias='custCatg'>
                            <attribute name='hil_category' />
                        </link-entity>
                    </entity>
                    </fetch>";
                        try
                        {
                            entcollJobs = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                            _totalClosedJobs = entcollJobs.Entities.Count;
                            if (_totalClosedJobs == 0)
                            {
                                //break;
                                continue;
                            }
                            int i = 0;
                            foreach (Entity entJob in entcollJobs.Entities)
                            {
                                _closedOn = entJob.GetAttributeValue<DateTime>("msdyn_timeclosed").AddMinutes(330);
                                _asOnDate = DateTime.Now.AddHours(-48);
                                if (_closedOn >= _asOnDate && entJob.GetAttributeValue<bool>("hil_reviewforcountryclassification") && entJob.GetAttributeValue<OptionSetValue>("hil_countryclassification").Value == 1)
                                {
                                    continue;
                                }
                                _performaInvoiceId = Guid.Empty;
                                _fixedCharges = 0;
                                _erFranchise = entJob.GetAttributeValue<EntityReference>("hil_owneraccount");
                                _performaInvoiceId = GeneratePerformaInvoice(_service, _erFranchise.Id, _erFiscalMonth.Id);

                                Console.WriteLine("Job # " + entJob.GetAttributeValue<string>("msdyn_name"));
                                #region Fixed Charge Calculation
                                _fixedCharges = CreateFixedChargeClaimLine(_service, _performaInvoiceId, ClaimCategoryConst._FixedCompensation, _erFiscalMonth, _erFranchise, Convert.ToDateTime(_endDate));
                                _claimStatus = entJob.GetAttributeValue<OptionSetValue>("hil_claimstatus");

                                if (_performaInvoiceId != null && (_claimStatus.Value == 4 || _claimStatus.Value == 7))
                                {
                                    GenerateClaimLines(entJob, _performaInvoiceId, _service, _erFiscalMonth);
                                    Console.WriteLine("Claim Cline generated");
                                }
                                entUpdate = new Entity("msdyn_workorder", entJob.Id);
                                entUpdate["hil_generateclaim"] = true;
                                entUpdate["hil_claimamount"] = new Money(0);
                                _service.Update(entUpdate);
                                #endregion

                                Console.WriteLine(i++.ToString() + "/" + _totalClosedJobs.ToString());
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                    }
                    #region Calculating TAT Incentive Slab
                    //Console.WriteLine("Calculating TAT Incentive Slab");
                    //Generate TAT Slab if all Claim status of all Work Order of that Franchisee not in (CLaim approved, Claim Rejected, Penalty Imposed, Penalty Waived)
                    //queryExp = new QueryExpression(msdyn_workorder.EntityLogicalName);
                    //queryExp.ColumnSet = new ColumnSet("hil_claimstatus", "msdyn_workorderid", "hil_tatachievementslab");
                    //queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                    //queryExp.Criteria.AddCondition("hil_claimstatus", ConditionOperator.NotIn, new object[] { 0, 4, 3, 8, 7 });
                    //queryExp.Criteria.AddCondition("msdyn_substatus", ConditionOperator.Equal, new Guid("1727FA6C-FA0F-E911-A94E-000D3AF060A1"));
                    //queryExp.Criteria.AddCondition("hil_owneraccount", ConditionOperator.Equal, entASP.GetAttributeValue<EntityReference>("hil_owneraccount").Id);
                    ////queryExp.Criteria.AddCondition("hil_owneraccount", ConditionOperator.Equal, new Guid("B5B58346-310B-E911-A94E-000D3AF06A98"));
                    //queryExp.Criteria.AddCondition("hil_isocr", ConditionOperator.NotEqual, true);
                    //queryExp.Criteria.AddCondition(new ConditionExpression("msdyn_timeclosed", ConditionOperator.OnOrAfter, _startDate));
                    //queryExp.Criteria.AddCondition(new ConditionExpression("msdyn_timeclosed", ConditionOperator.OnOrBefore, _endDate));
                    //entcollTATSlabJobs = _service.RetrieveMultiple(queryExp);
                    //if (entcollTATSlabJobs.Entities.Count == 0)
                    //{
                    //    GenerateTATClaimLines(_service, _performaInvoiceId, _erFranchise.Id, Convert.ToDateTime(_startDate), Convert.ToDateTime(_endDate));
                    //}
                    #endregion
                }
                #endregion
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        public static void RefreshJobClaimStatusBasedOnSAWActivity(IOrganizationService _service)
        {
            #region Variable declaration
            Guid _fiscalMonthId = Guid.Empty;
            DateTime? _startDate = null;
            DateTime? _endDate = null;
            string _strStartDate = null;
            string _strEndDate = null;
            Guid _performaInvoiceId = Guid.Empty;
            int _totalClosedJobs = 0;
            string _fetchXML = string.Empty;
            Decimal _fixedCharges = 0;
            QueryExpression queryExp;
            EntityCollection entcoll;
            EntityCollection entcollJobs;
            EntityReference _erFiscalMonth = null;
            EntityReference _erFranchise = null;
            OptionSetValue _claimStatus = null;
            string[] _jobColumns = { "hil_fiscalmonth", "msdyn_name", "createdon", "hil_owneraccount", "hil_channelpartnercategory", "hil_warrantystatus", "hil_countryclassification", "hil_tatachievementslab", "hil_tatcategory", "hil_schemecode", "hil_salesoffice", "hil_productcategory", "hil_kkgcode", "hil_typeofassignee", "hil_brand", "hil_productsubcategory", "hil_callsubtype", "hil_claimstatus", "hil_isocr", "hil_laborinwarranty", "hil_isgascharged", "hil_sparepartuse", "hil_jobclosemobile", "msdyn_timeclosed", "msdyn_customerasset", "hil_quantity", "hil_purchasedate" };
            string _performaRemarks = string.Empty;
            Entity entUpdate;
            #endregion

            #region Fetching Fiscal Month Date Range
            queryExp = new QueryExpression("hil_claimperiod");
            queryExp.ColumnSet = new ColumnSet("hil_fromdate", "hil_todate");
            queryExp.Criteria = new FilterExpression(LogicalOperator.And);
            queryExp.Criteria.AddCondition("hil_isapplicable", ConditionOperator.Equal, true);
            queryExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);

            entcoll = _service.RetrieveMultiple(queryExp);
            if (entcoll.Entities.Count > 0)
            {
                _startDate = entcoll.Entities[0].GetAttributeValue<DateTime>("hil_fromdate").AddMinutes(330);
                _strStartDate = _startDate.Value.Year.ToString() + "-" + _startDate.Value.Month.ToString().PadLeft(2, '0') + "-" + _startDate.Value.Day.ToString().PadLeft(2, '0');
                _endDate = entcoll.Entities[0].GetAttributeValue<DateTime>("hil_todate").AddMinutes(330);
                _strEndDate = _endDate.Value.Year.ToString() + "-" + _endDate.Value.Month.ToString().PadLeft(2, '0') + "-" + _endDate.Value.Day.ToString().PadLeft(2, '0');
                _erFiscalMonth = entcoll.Entities[0].ToEntityReference();
            }
            if (_erFiscalMonth == null)
            {
                Console.WriteLine("Fiscal Month is not defined." + DateTime.Now.ToString());
                return;
            }
            #endregion

            try
            {
                _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                      <entity name='msdyn_workorder'>
                        <attribute name='msdyn_name' />
                        <attribute name='msdyn_workorderid' />
                        <attribute name='msdyn_timeclosed' />
                        <attribute name='hil_claimstatus' />
                        <attribute name='hil_generateclaim' />
                        <attribute name='hil_owneraccount' />
                        <order attribute='msdyn_timeclosed' descending='true' />
                        <filter type='and'>
                          <condition attribute='msdyn_substatus' operator='eq' value='{1727FA6C-FA0F-E911-A94E-000D3AF060A1}' />
                          <condition attribute='hil_isocr' operator='ne' value='1' />
                          <condition attribute='hil_fiscalmonth' operator='eq' value='" + _erFiscalMonth.Id + @"' />
                          <condition attribute='hil_callsubtype' operator='not-null' />
                          <condition attribute='hil_salesoffice' operator='not-null' />
                          <condition attribute='hil_owneraccount' operator='not-null' />
                          <condition attribute='msdyn_customerasset' operator='not-null' />
                          <condition attribute='hil_claimstatus' operator='eq' value='1' />
                          <condition attribute='hil_typeofassignee' operator='in'>
                            <value uiname='Franchise' uitype='position'>{4A1AA189-1208-E911-A94D-000D3AF0694E}</value>
                            <value uiname='Franchise Technician' uitype='position'>{0197EA9B-1208-E911-A94D-000D3AF0694E}</value>
                          </condition>
                          <condition attribute='hil_productcategory' operator='not-null' />
                        </filter>
                        <link-entity name='account' from='accountid' to='hil_owneraccount' visible='false' link-type='outer' alias='a_061b877eea04e911a94d000d3af06c56'>
                          <attribute name='hil_salesoffice' />
                        </link-entity>
                        <link-entity name='hil_sawactivity' from='hil_jobid' to='msdyn_workorderid' link-type='inner' alias='ad'>
                          <filter type='and'>
                            <condition attribute='hil_approvalstatus' operator='eq' value='3' />
                          </filter>
                        </link-entity>
                      </entity>
                    </fetch>";
                try
                {
                    entcollJobs = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                    _totalClosedJobs = entcollJobs.Entities.Count;
                    int i = 0;
                    foreach (Entity entJob in entcollJobs.Entities)
                    {
                        string _fetchXML1 = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                          <entity name='hil_sawactivity'>
                            <attribute name='hil_sawactivityid' />
                            <attribute name='hil_name' />
                            <attribute name='createdon' />
                            <attribute name='statecode' />
                            <attribute name='hil_approvalstatus' />
                            <order attribute='hil_name' descending='false' />
                            <filter type='and'>
                              <condition attribute='hil_jobid' operator='eq' value='{entJob.Id}' />
                              <condition attribute='statecode' operator='eq' value='0' />
                              <condition attribute='hil_approvalstatus' operator='ne' value='3' />
                            </filter>
                          </entity>
                        </fetch>";
                        EntityCollection entcollSAW = _service.RetrieveMultiple(new FetchExpression(_fetchXML1));
                        if (entcollSAW.Entities.Count == 0)
                        {
                            entUpdate = new Entity("msdyn_workorder", entJob.Id);
                            entUpdate["hil_claimstatus"] = new OptionSetValue(4);
                            _service.Update(entUpdate);
                        }
                        Console.WriteLine(i++.ToString() + "/" + _totalClosedJobs.ToString());
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        public static void GenerateClaimLinesCountryClassification(IOrganizationService _service)
        {
            #region Variable declaration
            Guid _fiscalMonthId = Guid.Empty;
            DateTime? _startDate = null;
            DateTime? _endDate = null;
            string _strStartDate = null;
            string _strEndDate = null;
            Guid _performaInvoiceId = Guid.Empty;
            int _totalClosedJobs = 0;
            string _fetchXML = string.Empty;
            Decimal _fixedCharges = 0;
            QueryExpression queryExp;
            EntityCollection entcoll;
            EntityCollection entcollJobs;
            EntityReference _erFiscalMonth = null;
            EntityReference _erFranchise = null;
            OptionSetValue _claimStatus = null;
            string[] _jobColumns = { "hil_fiscalmonth", "msdyn_name", "createdon", "hil_owneraccount", "hil_channelpartnercategory", "hil_warrantystatus", "hil_countryclassification", "hil_tatachievementslab", "hil_tatcategory", "hil_schemecode", "hil_salesoffice", "hil_productcategory", "hil_kkgcode", "hil_typeofassignee", "hil_brand", "hil_productsubcategory", "hil_callsubtype", "hil_claimstatus", "hil_isocr", "hil_laborinwarranty", "hil_isgascharged", "hil_sparepartuse", "hil_jobclosemobile", "msdyn_timeclosed", "msdyn_customerasset", "hil_quantity", "hil_purchasedate" };
            string _performaRemarks = string.Empty;
            Entity entUpdate;
            #endregion

            #region Fetching Fiscal Month Date Range
            //DateTime _processDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);
            //DateTime _processDate = new DateTime(2022, 10, 20);
            queryExp = new QueryExpression("hil_claimperiod");
            queryExp.ColumnSet = new ColumnSet("hil_fromdate", "hil_todate");
            queryExp.Criteria = new FilterExpression(LogicalOperator.And);
            queryExp.Criteria.AddCondition("hil_isapplicable", ConditionOperator.Equal, true);
            queryExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
            //queryExp.Criteria.AddCondition("hil_fromdate", ConditionOperator.OnOrBefore, _processDate);
            //queryExp.Criteria.AddCondition("hil_todate", ConditionOperator.OnOrAfter, _processDate);

            entcoll = _service.RetrieveMultiple(queryExp);
            if (entcoll.Entities.Count > 0)
            {
                _startDate = entcoll.Entities[0].GetAttributeValue<DateTime>("hil_fromdate").AddMinutes(330);
                _strStartDate = _startDate.Value.Year.ToString() + "-" + _startDate.Value.Month.ToString().PadLeft(2, '0') + "-" + _startDate.Value.Day.ToString().PadLeft(2, '0');
                _endDate = entcoll.Entities[0].GetAttributeValue<DateTime>("hil_todate").AddMinutes(330);
                _strEndDate = _endDate.Value.Year.ToString() + "-" + _endDate.Value.Month.ToString().PadLeft(2, '0') + "-" + _endDate.Value.Day.ToString().PadLeft(2, '0');
                _erFiscalMonth = entcoll.Entities[0].ToEntityReference();
            }
            if (_erFiscalMonth == null)
            {
                Console.WriteLine("Fiscal Month is not defined." + DateTime.Now.ToString());
                return;
            }
            #endregion

            #region Refreshing Claim Parameters & Generating applicable Claim Lines
            try
            {
                /* <condition attribute='hil_claimstatus' operator='in'>
                    <value>4</value>
                    <value>7</value>
                    </condition>
                */
                _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                        <entity name='msdyn_workorder'>
                        <attribute name='hil_fiscalmonth' />
                        <attribute name='msdyn_name' />
                        <attribute name='createdon' />
                        <attribute name='hil_owneraccount' />
                        <attribute name='hil_channelpartnercategory' />
                        <attribute name='hil_warrantystatus' />
                        <attribute name='hil_countryclassification' />
                        <attribute name='hil_tatachievementslab' />
                        <attribute name='hil_tatcategory' />
                        <attribute name='hil_schemecode' />
                        <attribute name='hil_salesoffice' />
                        <attribute name='hil_productcategory' />
                        <attribute name='hil_kkgcode' />
                        <attribute name='hil_typeofassignee' />
                        <attribute name='hil_brand' />
                        <attribute name='hil_productsubcategory' />
                        <attribute name='hil_callsubtype' />
                        <attribute name='hil_claimstatus' />
                        <attribute name='hil_isocr' />
                        <attribute name='hil_laborinwarranty' />
                        <attribute name='hil_isgascharged' />
                        <attribute name='hil_sparepartuse' />
                        <attribute name='hil_jobclosemobile' />
                        <attribute name='msdyn_timeclosed' />
                        <attribute name='msdyn_customerasset' />
                        <attribute name='hil_quantity' />
                        <attribute name='hil_purchasedate' />
                        <attribute name='hil_sourceofjob' />
                        <attribute name='hil_pincode' />
                        <attribute name='hil_reviewforcountryclassification' />
                        <attribute name='hil_countryclassification' />
                        <attribute name='msdyn_timeclosed' />
                        <order attribute='msdyn_timeclosed' descending='true' />
                        <filter type='and'>
                            <condition attribute='hil_fiscalmonth' operator='eq' value='" + _erFiscalMonth.Id + @"' />
                            <condition attribute='hil_generateclaim' operator='eq' value='0' />
                            <condition attribute='hil_claimstatus' operator='ne' value='3' />
                            <condition attribute='hil_reviewforcountryclassification' operator='eq' value='1' />
                            <condition attribute='hil_countryclassification' operator='eq' value='1' />
                        </filter>
                        <link-entity name='account' from='accountid' to='hil_owneraccount' visible='false' link-type='outer' alias='custCatg'>
                            <attribute name='hil_category' />
                        </link-entity>
                    </entity>
                    </fetch>";
                try
                {
                    entcollJobs = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                    _totalClosedJobs = entcollJobs.Entities.Count;
                    int i = 0;
                    DateTime _closedOn, _asOnDate;
                    foreach (Entity entJob in entcollJobs.Entities)
                    {
                        _performaInvoiceId = Guid.Empty;
                        _fixedCharges = 0;
                        _erFranchise = entJob.GetAttributeValue<EntityReference>("hil_owneraccount");
                        _performaInvoiceId = GeneratePerformaInvoice(_service, _erFranchise.Id, _erFiscalMonth.Id);

                        _closedOn = entJob.GetAttributeValue<DateTime>("msdyn_timeclosed").AddMinutes(330);

                        //Start Block --- Uncomment This Code Block on 20th Day of Clain Month
                        _asOnDate = DateTime.Now.AddHours(-72);
                        DateTime _processDate = DateTime.Now;
                        if (_processDate.Day!=20 && _closedOn >= _asOnDate && entJob.GetAttributeValue<bool>("hil_reviewforcountryclassification") && entJob.GetAttributeValue<OptionSetValue>("hil_countryclassification").Value == 1)
                        {
                            Console.WriteLine(i++.ToString() + "/" + _totalClosedJobs.ToString());
                            continue;
                        }
                        // END Block

                        Console.WriteLine("Job # " + entJob.GetAttributeValue<string>("msdyn_name"));
                        #region Fixed Charge Calculation
                        _fixedCharges = CreateFixedChargeClaimLine(_service, _performaInvoiceId, ClaimCategoryConst._FixedCompensation, _erFiscalMonth, _erFranchise, Convert.ToDateTime(_endDate));
                        _claimStatus = entJob.GetAttributeValue<OptionSetValue>("hil_claimstatus");

                        if (_performaInvoiceId != null && (_claimStatus.Value == 4 || _claimStatus.Value == 7))
                        {
                            GenerateClaimLines(entJob, _performaInvoiceId, _service, _erFiscalMonth);
                            Console.WriteLine("Claim Cline generated");
                        }
                        entUpdate = new Entity("msdyn_workorder", entJob.Id);
                        entUpdate["hil_generateclaim"] = true;
                        entUpdate["hil_claimamount"] = new Money(0);
                        _service.Update(entUpdate);
                        #endregion

                        Console.WriteLine(i++.ToString() + "/" + _totalClosedJobs.ToString());
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
                #endregion
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
       }
        public static void GenerateClaimLinesBackLogProdSubcatgWise(IOrganizationService _service)
        {

            #region Variable declaration
            Guid _fiscalMonthId = new Guid("DC417E88-E62B-EB11-A813-0022486E5C9D");
            //Guid _fiscalMonthId = Guid.Empty;
            DateTime? _startDate = null;
            DateTime? _endDate = null;
            Guid _performaInvoiceId = Guid.Empty;
            int _totalClosedJobs = 0;
            string _fetchXML = string.Empty;
            Decimal _fixedCharges = 0;
            QueryExpression queryExp;
            EntityCollection entcoll;
            EntityCollection entcollJobs;
            EntityCollection entcollTATSlabJobs;
            EntityReference _erFiscalMonth = null;
            EntityReference _erFranchise = null;
            OptionSetValue _claimStatus = null;
            string[] _jobColumns = { "hil_fiscalmonth", "msdyn_name", "createdon", "hil_owneraccount", "hil_channelpartnercategory", "hil_warrantystatus", "hil_countryclassification", "hil_tatachievementslab", "hil_tatcategory", "hil_schemecode", "hil_salesoffice", "hil_productcategory", "hil_kkgcode", "hil_typeofassignee", "hil_brand", "hil_productsubcategory", "hil_callsubtype", "hil_claimstatus", "hil_isocr", "hil_laborinwarranty", "hil_isgascharged", "hil_sparepartuse", "hil_jobclosemobile", "msdyn_timeclosed", "msdyn_customerasset", "hil_quantity", "hil_purchasedate" };
            string _performaRemarks = string.Empty;
            Entity entUpdate;
            #endregion

            #region Fetching Fiscal Month Date Range
            if (_fiscalMonthId != Guid.Empty)
            {
                Entity ent = _service.Retrieve("hil_claimperiod", _fiscalMonthId, new ColumnSet("hil_fromdate", "hil_todate"));
                if (ent != null)
                {
                    _startDate = ent.GetAttributeValue<DateTime>("hil_fromdate").AddMinutes(330);
                    _endDate = ent.GetAttributeValue<DateTime>("hil_todate").AddMinutes(330);
                    _erFiscalMonth = ent.ToEntityReference();
                }
            }
            else
            {

                DateTime _processDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);

                queryExp = new QueryExpression("hil_claimperiod");
                queryExp.ColumnSet = new ColumnSet("hil_fromdate", "hil_todate");
                queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                queryExp.Criteria.AddCondition("hil_fromdate", ConditionOperator.OnOrBefore, _processDate);
                queryExp.Criteria.AddCondition("hil_todate", ConditionOperator.OnOrAfter, _processDate);
                entcoll = _service.RetrieveMultiple(queryExp);
                if (entcoll.Entities.Count > 0)
                {
                    _startDate = entcoll.Entities[0].GetAttributeValue<DateTime>("hil_fromdate").AddMinutes(330);
                    _endDate = entcoll.Entities[0].GetAttributeValue<DateTime>("hil_todate").AddMinutes(330);
                    _erFiscalMonth = entcoll.Entities[0].ToEntityReference();
                }
            }
            #endregion

            #region Refreshing Claim Parameters & Generating applicable Claim Lines
            try
            {
                //<condition attribute='hil_owneraccount' operator='eq' value='{B5B58346-310B-E911-A94E-000D3AF06A98}' />
                //<condition attribute='hil_generateclaim' operator='ne' value='1' />

                _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                    <entity name='msdyn_workorder'>
                    <attribute name='hil_owneraccount' />
                    <filter type='and'>
                        <condition attribute='msdyn_substatus' operator='eq' value='{1727FA6C-FA0F-E911-A94E-000D3AF060A1}' />
                        <condition attribute='hil_isocr' operator='ne' value='1' />
                        <condition attribute='hil_claimstatus' operator='eq' value='4' />
                        <condition attribute='hil_productsubcategory' operator='eq' value='{871B7022-410B-E911-A94F-000D3AF00F43}' />
                        <condition attribute='hil_fiscalmonth' operator='eq' value='{DC417E88-E62B-EB11-A813-0022486E5C9D}' />
                        <condition attribute='msdyn_timeclosed' operator='on-or-after' value='2020-10-21' />
                        <condition attribute='msdyn_timeclosed' operator='on-or-before' value='2020-11-20' />
                        <condition attribute='hil_typeofassignee' operator='in'>
                            <value uiname='Franchise' uitype='position'>{4A1AA189-1208-E911-A94D-000D3AF0694E}</value>
                            <value uiname='Franchise Technician' uitype='position'>{0197EA9B-1208-E911-A94D-000D3AF0694E}</value>
                        </condition>
                    </filter>
                    </entity>
                    </fetch>";

                /*
                <link-entity name='account' from='accountid' to='hil_owneraccount' link-type='inner' alias='ad'>
                <filter type='and'>
                <condition attribute='hil_salesoffice' operator='eq' value='{df2f0cc1-bcf7-e811-a94c-000d3af0677f}' />
                </filter>
                </link-entity>
                 * */
                entcoll = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                foreach (Entity entASP in entcoll.Entities)
                {
                    _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                        <entity name='msdyn_workorder'>
                        <attribute name='hil_fiscalmonth' />
                        <attribute name='msdyn_name' />
                        <attribute name='createdon' />
                        <attribute name='hil_owneraccount' />
                        <attribute name='hil_channelpartnercategory' />
                        <attribute name='hil_warrantystatus' />
                        <attribute name='hil_countryclassification' />
                        <attribute name='hil_tatachievementslab' />
                        <attribute name='hil_tatcategory' />
                        <attribute name='hil_schemecode' />
                        <attribute name='hil_salesoffice' />
                        <attribute name='hil_productcategory' />
                        <attribute name='hil_kkgcode' />
                        <attribute name='hil_typeofassignee' />
                        <attribute name='hil_brand' />
                        <attribute name='hil_productsubcategory' />
                        <attribute name='hil_callsubtype' />
                        <attribute name='hil_claimstatus' />
                        <attribute name='hil_isocr' />
                        <attribute name='hil_laborinwarranty' />
                        <attribute name='hil_isgascharged' />
                        <attribute name='hil_sparepartuse' />
                        <attribute name='hil_jobclosemobile' />
                        <attribute name='msdyn_timeclosed' />
                        <attribute name='msdyn_customerasset' />
                        <attribute name='hil_quantity' />
                        <attribute name='hil_purchasedate' />
                        <attribute name='hil_sourceofjob' />
                        <attribute name='hil_pincode' />
                        <filter type='and'>
                            <condition attribute='hil_owneraccount' operator='eq'  value='" + entASP.GetAttributeValue<EntityReference>("hil_owneraccount").Id + @"' />
                            <condition attribute='msdyn_substatus' operator='eq' value='{1727FA6C-FA0F-E911-A94E-000D3AF060A1}' />
                            <condition attribute='hil_isocr' operator='ne' value='1' />
                            <condition attribute='hil_claimstatus' operator='eq' value='4' />
                            <condition attribute='hil_productsubcategory' operator='eq' value='{871B7022-410B-E911-A94F-000D3AF00F43}' />
                            <condition attribute='hil_fiscalmonth' operator='eq' value='{DC417E88-E62B-EB11-A813-0022486E5C9D}' />
                            <condition attribute='msdyn_timeclosed' operator='on-or-after' value='2020-10-21' />
                            <condition attribute='msdyn_timeclosed' operator='on-or-before' value='2020-11-20' />
                            <condition attribute='hil_typeofassignee' operator='in'>
                                <value uiname='Franchise' uitype='position'>{4A1AA189-1208-E911-A94D-000D3AF0694E}</value>
                                <value uiname='Franchise Technician' uitype='position'>{0197EA9B-1208-E911-A94D-000D3AF0694E}</value>
                            </condition>
                        </filter>
                        <link-entity name='account' from='accountid' to='hil_owneraccount' visible='false' link-type='outer' alias='custCatg'>
                            <attribute name='hil_category' />
                        </link-entity>
                    </entity>
                    </fetch>";
                    try
                    {
                        entcollJobs = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                        _totalClosedJobs = entcollJobs.Entities.Count;
                        if (_totalClosedJobs == 0) { break; }
                        int i = 0;
                        foreach (Entity entJob in entcollJobs.Entities)
                        {
                            _performaInvoiceId = Guid.Empty;
                            _fixedCharges = 0;
                            _erFranchise = entJob.GetAttributeValue<EntityReference>("hil_owneraccount");
                            _performaInvoiceId = GeneratePerformaInvoice(_service, _erFranchise.Id, _erFiscalMonth.Id);

                            Console.WriteLine("Job # " + entJob.GetAttributeValue<string>("msdyn_name"));
                            #region Fixed Charge Calculation
                            _fixedCharges = CreateFixedChargeClaimLine(_service, _performaInvoiceId, ClaimCategoryConst._FixedCompensation, _erFiscalMonth, _erFranchise, Convert.ToDateTime(_endDate));
                            _claimStatus = entJob.GetAttributeValue<OptionSetValue>("hil_claimstatus");

                            if (_performaInvoiceId != null && _claimStatus.Value == 4)
                            {
                                GenerateClaimLines(entJob, _performaInvoiceId, _service, _erFiscalMonth);
                                Console.WriteLine("Claim Cline generated");
                            }
                            entUpdate = new Entity("msdyn_workorder", entJob.Id);
                            entUpdate["hil_generateclaim"] = true;
                            entUpdate["hil_claimamount"] = new Money(0);
                            _service.Update(entUpdate);
                            #endregion
                            Console.WriteLine(i++.ToString() + "/" + _totalClosedJobs.ToString());
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    //#region Calculating TAT Incentive Slab
                    //Console.WriteLine("Calculating TAT Incentive Slab");
                    //// Generate TAT Slab if all Claim status of all Work Order of that Franchisee not in (CLaim approved, Claim Rejected, Penalty Imposed, Penalty Waived)
                    //queryExp = new QueryExpression(msdyn_workorder.EntityLogicalName);
                    //queryExp.ColumnSet = new ColumnSet("hil_claimstatus", "msdyn_workorderid", "hil_tatachievementslab");
                    //queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                    //queryExp.Criteria.AddCondition("hil_claimstatus", ConditionOperator.NotIn, new object[] { 0, 4, 3, 8, 7 });
                    //queryExp.Criteria.AddCondition("msdyn_substatus", ConditionOperator.Equal, new Guid("1727FA6C-FA0F-E911-A94E-000D3AF060A1"));
                    //queryExp.Criteria.AddCondition("hil_owneraccount", ConditionOperator.Equal, entASP.GetAttributeValue<EntityReference>("hil_owneraccount").Id);
                    ////queryExp.Criteria.AddCondition("hil_owneraccount", ConditionOperator.Equal, new Guid("B5B58346-310B-E911-A94E-000D3AF06A98"));
                    //queryExp.Criteria.AddCondition("hil_isocr", ConditionOperator.NotEqual, true);
                    //queryExp.Criteria.AddCondition(new ConditionExpression("msdyn_timeclosed", ConditionOperator.OnOrAfter, _startDate));
                    //queryExp.Criteria.AddCondition(new ConditionExpression("msdyn_timeclosed", ConditionOperator.OnOrBefore, _endDate));
                    //entcollTATSlabJobs = _service.RetrieveMultiple(queryExp);
                    //if (entcollTATSlabJobs.Entities.Count == 0)
                    //{
                    //    GenerateTATClaimLines(_service, _performaInvoiceId, _erFranchise.Id, Convert.ToDateTime(_startDate), Convert.ToDateTime(_endDate));
                    //}
                    //#endregion
                }
                #endregion
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        public static void GenerateTATIncentivesResetChannelPartnerFlag(IOrganizationService _service)
        {

            #region Variable declaration
            Guid _fiscalMonthId = Guid.Empty;
            string _fetchXML = string.Empty;
            QueryExpression queryExp;
            EntityCollection entcoll;
            EntityReference _erFiscalMonth = null;
            #endregion

            #region Fetching Fiscal Month Date Range
            queryExp = new QueryExpression("hil_claimperiod");
            queryExp.ColumnSet = new ColumnSet("hil_fromdate", "hil_todate");
            queryExp.Criteria = new FilterExpression(LogicalOperator.And);
            queryExp.Criteria.AddCondition("hil_isapplicable", ConditionOperator.Equal, true);
            queryExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
            entcoll = _service.RetrieveMultiple(queryExp);
            if (entcoll.Entities.Count > 0)
            {
                _erFiscalMonth = entcoll.Entities[0].ToEntityReference();
            }
            if (_erFiscalMonth == null)
            {
                Console.WriteLine("Fiscal Month is not defined." + DateTime.Now.ToString());
                return;
            }
            #endregion

            try
            {
                _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                    <entity name='msdyn_workorder'>
                    <attribute name='hil_owneraccount' />
                    <filter type='and'>
                        <condition attribute='hil_fiscalmonth' operator='eq' value='" + _erFiscalMonth.Id + @"' />
                    </filter>
                    <link-entity name='account' from='accountid' to='hil_owneraccount' link-type='inner' alias='ae'>
                        <filter type='and'>
                            <condition attribute='hil_claimprocessed' operator='eq' value='1' />
                        </filter>
                    </link-entity>
                    </entity>
                    </fetch>";

                entcoll = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                int i = 1;
                Entity entCustomer = null;
                foreach (Entity entASP in entcoll.Entities)
                {
                    entCustomer = new Entity("account");
                    entCustomer.Id = entASP.GetAttributeValue<EntityReference>("hil_owneraccount").Id;
                    entCustomer["hil_claimprocessed"] = false;
                    _service.Update(entCustomer);
                    Console.WriteLine("Flag reset: " + (i++).ToString());
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        public static void GenerateTATIncentivesClaimASPWise(IOrganizationService _service)
        {
            #region Variable declaration
            Guid _fiscalMonthId = Guid.Empty;
            DateTime? _startDate = null;
            DateTime? _endDate = null;
            Guid _performaInvoiceId = Guid.Empty;
            string _fetchXML = string.Empty;
            QueryExpression queryExp;
            EntityCollection entcoll;
            EntityCollection entcollTATSlabJobs;
            EntityReference _erFiscalMonth = null;
            EntityReference _erFranchise = null;
            string[] _jobColumns = { "hil_fiscalmonth", "msdyn_name", "createdon", "hil_owneraccount", "hil_channelpartnercategory", "hil_warrantystatus", "hil_countryclassification", "hil_tatachievementslab", "hil_tatcategory", "hil_schemecode", "hil_salesoffice", "hil_productcategory", "hil_kkgcode", "hil_typeofassignee", "hil_brand", "hil_productsubcategory", "hil_callsubtype", "hil_claimstatus", "hil_isocr", "hil_laborinwarranty", "hil_isgascharged", "hil_sparepartuse", "hil_jobclosemobile", "msdyn_timeclosed", "msdyn_customerasset", "hil_quantity", "hil_purchasedate" };
            string _performaRemarks = string.Empty;
            #endregion

            #region Fetching Fiscal Month Date Range
            var _salesOfficeId = ConfigurationManager.AppSettings["SalesOfficeId"].ToString();
            //DateTime _processDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);
            queryExp = new QueryExpression("hil_claimperiod");
            queryExp.ColumnSet = new ColumnSet("hil_fromdate", "hil_todate");
            queryExp.Criteria = new FilterExpression(LogicalOperator.And);
            queryExp.Criteria.AddCondition("hil_isapplicable", ConditionOperator.Equal, true);
            queryExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
            entcoll = _service.RetrieveMultiple(queryExp);
            if (entcoll.Entities.Count > 0)
            {
                _startDate = entcoll.Entities[0].GetAttributeValue<DateTime>("hil_fromdate").AddMinutes(330);
                _endDate = entcoll.Entities[0].GetAttributeValue<DateTime>("hil_todate").AddMinutes(330);
                _erFiscalMonth = entcoll.Entities[0].ToEntityReference();
            }
            if (_erFiscalMonth == null)
            {
                Console.WriteLine("Fiscal Month is not defined." + DateTime.Now.ToString());
                return;
            }
            #endregion

            #region Refreshing Claim Parameters & Generating applicable Claim Lines
            try
            {
                //<condition attribute='hil_salesoffice' operator='eq' value='{ee8cffde-bcf7-e811-a94c-000d3af06cd4}' />
                //<condition attribute='hil_brand' operator='eq' value='2' />
                //<condition attribute='hil_salesoffice' operator='eq' value='{"+ _salesOfficeId + @"}' />
                //<condition attribute='hil_claimprocessed' operator='ne' value='1' />

                _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                    <entity name='msdyn_workorder'>
                    <attribute name='hil_owneraccount' />
                    <filter type='and'>
                        <condition attribute='hil_fiscalmonth' operator='eq' value='" + _erFiscalMonth.Id + @"' />
                    </filter>
                    <link-entity name='account' from='accountid' to='hil_owneraccount' link-type='inner' alias='ae'>
                        <filter type='and'>
                            <condition attribute='hil_claimprocessed' operator='ne' value='1' />
                        </filter>
                    </link-entity>
                    </entity>
                    </fetch>";

                entcoll = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                int i = 1;
                Entity entCustomer = null;
                foreach (Entity entASP in entcoll.Entities)
                {
                    try
                    {
                        _performaInvoiceId = Guid.Empty;
                        _erFranchise = entASP.GetAttributeValue<EntityReference>("hil_owneraccount");
                        _performaInvoiceId = GeneratePerformaInvoice(_service, _erFranchise.Id, _erFiscalMonth.Id);

                        #region Calculating TAT Incentive Slab
                        Console.WriteLine("Calculating TAT Incentive Slab... " + (i++).ToString() + "/" + entcoll.Entities.Count.ToString());
                        // Generate TAT Slab if all Claim status of all Work Order of that Franchisee not in (CLaim approved, Claim Rejected, Penalty Imposed, Penalty Waived)
                        queryExp = new QueryExpression(msdyn_workorder.EntityLogicalName);
                        queryExp.ColumnSet = new ColumnSet("msdyn_name", "hil_claimstatus", "msdyn_workorderid", "hil_tatachievementslab");
                        queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                        queryExp.Criteria.AddCondition("hil_claimstatus", ConditionOperator.NotIn, new object[] { 0, 4, 3, 8, 7 });
                        queryExp.Criteria.AddCondition("msdyn_substatus", ConditionOperator.Equal, new Guid("1727FA6C-FA0F-E911-A94E-000D3AF060A1"));
                        queryExp.Criteria.AddCondition("hil_owneraccount", ConditionOperator.Equal, entASP.GetAttributeValue<EntityReference>("hil_owneraccount").Id);
                        queryExp.Criteria.AddCondition("hil_isocr", ConditionOperator.NotEqual, true);
                        queryExp.Criteria.AddCondition(new ConditionExpression("msdyn_timeclosed", ConditionOperator.OnOrAfter, _startDate));
                        queryExp.Criteria.AddCondition(new ConditionExpression("msdyn_timeclosed", ConditionOperator.OnOrBefore, _endDate));
                        entcollTATSlabJobs = _service.RetrieveMultiple(queryExp);
                        if (entcollTATSlabJobs.Entities.Count == 0)
                        {
                            GenerateTATClaimLines(_service, _performaInvoiceId, _erFranchise.Id, Convert.ToDateTime(_startDate), Convert.ToDateTime(_endDate));
                        }
                        entCustomer = new Entity("account");
                        entCustomer.Id = entASP.GetAttributeValue<EntityReference>("hil_owneraccount").Id;
                        entCustomer["hil_claimprocessed"] = true;
                        _service.Update(entCustomer);
                        #endregion
                    }
                    catch (Exception)
                    {

                        throw;
                    }
                }
                #endregion
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        public static void GenerateTATIncentivesClaimASPWise_Temp(IOrganizationService _service)
        {
            #region Variable declaration
            Guid _fiscalMonthId = Guid.Empty;
            DateTime? _startDate = null;
            DateTime? _endDate = null;
            Guid _performaInvoiceId = Guid.Empty;
            string _fetchXML = string.Empty;
            QueryExpression queryExp;
            EntityCollection entcoll;
            EntityCollection entcollTATSlabJobs;
            EntityReference _erFiscalMonth = null;
            EntityReference _erFranchise = null;
            string[] _jobColumns = { "hil_fiscalmonth", "msdyn_name", "createdon", "hil_owneraccount", "hil_channelpartnercategory", "hil_warrantystatus", "hil_countryclassification", "hil_tatachievementslab", "hil_tatcategory", "hil_schemecode", "hil_salesoffice", "hil_productcategory", "hil_kkgcode", "hil_typeofassignee", "hil_brand", "hil_productsubcategory", "hil_callsubtype", "hil_claimstatus", "hil_isocr", "hil_laborinwarranty", "hil_isgascharged", "hil_sparepartuse", "hil_jobclosemobile", "msdyn_timeclosed", "msdyn_customerasset", "hil_quantity", "hil_purchasedate" };
            string _performaRemarks = string.Empty;
            #endregion

            #region Fetching Fiscal Month Date Range
            //DateTime _processDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);
            queryExp = new QueryExpression("hil_claimperiod");
            queryExp.ColumnSet = new ColumnSet("hil_fromdate", "hil_todate");
            queryExp.Criteria = new FilterExpression(LogicalOperator.And);
            queryExp.Criteria.AddCondition("hil_isapplicable", ConditionOperator.Equal, true);
            queryExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
            entcoll = _service.RetrieveMultiple(queryExp);
            if (entcoll.Entities.Count > 0)
            {
                _startDate = entcoll.Entities[0].GetAttributeValue<DateTime>("hil_fromdate").AddMinutes(330);
                _endDate = entcoll.Entities[0].GetAttributeValue<DateTime>("hil_todate").AddMinutes(330);
                _erFiscalMonth = entcoll.Entities[0].ToEntityReference();
            }
            if (_erFiscalMonth == null)
            {
                Console.WriteLine("Fiscal Month is not defined." + DateTime.Now.ToString());
                return;
            }
            #endregion

            #region Refreshing Claim Parameters & Generating applicable Claim Lines
            try
            {
                //<condition attribute='hil_salesoffice' operator='eq' value='{ee8cffde-bcf7-e811-a94c-000d3af06cd4}' />
                _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                <entity name='msdyn_workorder'>
                <attribute name='hil_owneraccount' />
                <filter type='and'>
                    <condition attribute='hil_fiscalmonth' operator='eq' value='" + _erFiscalMonth.Id + @"' />
                    <condition attribute='hil_owneraccount' operator='in'>
                    <value uiname='PRO SERVICES (ASP)' uitype='account'>{CE6AC05D-310B-E911-A94E-000D3AF06091}</value>
                    <value uiname='Prakash Refrigeration And Elec' uitype='account'>{0CD4B31D-8E39-EA11-A813-000D3AF057DD}</value>
                    <value uiname='Quick Service Network' uitype='account'>{7A3BCE64-310B-E911-A94E-000D3AF06091}</value>
                    <value uiname='Quick Solution Point' uitype='account'>{4BD04A73-8D22-EC11-B6E6-6045BD731341}</value>
                    <value uiname='Red Star' uitype='account'>{42A41452-0B11-EA11-A811-000D3AF057DD}</value>
                    <value uiname='Ruby Electric Works' uitype='account'>{4D51D33F-F8EB-E911-A812-000D3AF058B5}</value>
                    <value uiname='Rukmini Enterprises' uitype='account'>{7B07E6A6-FE99-EC11-B401-6045BDAAFC34}</value>
                    <value uiname='Sagar Enterprises' uitype='account'>{316975A7-FE99-EC11-B401-6045BDAAF7C7}</value>
                    <value uiname='Sai Purvi Enterprises' uitype='account'>{848960A0-310B-E911-A94D-000D3AF0677F}</value>
                    <value uiname='Saini Ref. And Air Cond. Servi' uitype='account'>{538F05B5-B347-EA11-A812-000D3AF057DD}</value>
                    <value uiname='Sarotia Enterprises' uitype='account'>{0A57EC3D-FC4D-EA11-A812-000D3AF05D7B}</value>
                    <value uiname='Sai Ultimate Cooling Solutions' uitype='account'>{CA262794-310B-E911-A94E-000D3AF06CD4}</value>
                    <value uiname='Serv Simplified Appliances Pvt' uitype='account'>{B358886C-7B32-EA11-A813-000D3AF05D7B}</value>
                    <value uiname='Skyline Electronics' uitype='account'>{B7042FA1-D7B4-E911-A95B-000D3AF0694E}</value>
                    <value uiname='SKY Refrigeration' uitype='account'>{3F86ECB2-B347-EA11-A812-000D3AF055B6}</value>
                    <value uiname='Smart Care Solutions' uitype='account'>{3660EFCB-310B-E911-A94D-000D3AF06C56}</value>
                    <value uiname='Smart Electronics' uitype='account'>{EDDD998A-FC58-EA11-A811-000D3AF05F57}</value>
                    <value uiname='S.P. Enterprise' uitype='account'>{C8A4F6CC-310B-E911-A94E-000D3AF06A98}</value>
                    <value uiname='Sri Sreenivaasa Services Cente' uitype='account'>{2B0767B4-BD91-EB11-B1AC-0022486E7094}</value>
                    <value uiname='STAR REFRIGERATION (ASP)' uitype='account'>{63053BEE-310B-E911-A94D-000D3AF0694E}</value>
                    <value uiname='Technotronic Enterprises' uitype='account'>{8C5A3701-320B-E911-A94E-000D3AF06A98}</value>
                    <value uiname='Ugam Refrigeration' uitype='account'>{73C6D114-8DF9-EA11-A815-000D3AF055B6}</value>
                    <value uiname='Unique Service Centre' uitype='account'>{61E25E6A-EA00-EC11-94EF-6045BD72E02A}</value>
                    <value uiname='VAHANVATI AIR (ASP)' uitype='account'>{B376ADBE-300B-E911-A94D-000D3AF06C56}</value>
                    <value uiname='Vansh Enterprises' uitype='account'>{456BA423-B558-EC11-8F8F-002248D4E579}</value>
                    <value uiname='Yadav Cool Point' uitype='account'>{735028FE-300B-E911-A94D-000D3AF06C56}</value>
                    <value uiname='Yes Electric' uitype='account'>{BD239DF6-443E-EA11-A812-000D3AF0563C}</value>
                    <value uiname='Aarti Refrigration &amp; Air Condi' uitype='account'>{EAF10715-300B-E911-A94E-000D3AF06CD4}</value>
                    <value uiname='Aayat Refrigeration' uitype='account'>{DA7636EB-10A2-EA11-A812-000D3AF05D7B}</value>
                    <value uiname='ANAND REFRIGERATION(ASP)' uitype='account'>{57ED9930-300B-E911-A94D-000D3AF0694E}</value>
                    <value uiname='Anurag Multi Service' uitype='account'>{2DE0A42B-300B-E911-A94D-000D3AF0677F}</value>
                    <value uiname='Bhavya Enterprises' uitype='account'>{3F94BE4E-300B-E911-A94E-000D3AF06A98}</value>
                    <value uiname='Bharti And Sons' uitype='account'>{ECB09D5E-8AC0-E911-A957-000D3AF0677F}</value>
                    <value uiname='COOL SERVICE CENTER - ASP' uitype='account'>{17585065-300B-E911-A94D-000D3AF06C56}</value>
                    <value uiname='Digital Service Center' uitype='account'>{719A340D-41BA-E911-A961-000D3AF06091}</value>
                    <value uiname='Future Services' uitype='account'>{7C42681B-BB2E-EB11-A813-000D3AF055B6}</value>
                    <value uiname='LUCKY REFRIGRATION (ASP)' uitype='account'>{CF1C4334-310B-E911-A94E-000D3AF06A98}</value>
                    <value uiname='Mahi Enterprises' uitype='account'>{D4A747CF-E8DD-E911-A812-000D3AF057DD}</value>
                    <value uiname='Modern Electronics' uitype='account'>{71EC6720-310B-E911-A94E-000D3AF06091}</value>
                    <value uiname='M.s. Electronics' uitype='account'>{1FC6FDB1-B7B0-EC11-9840-6045BDAA92C3}</value>
                    <value uiname='Naina Electrical Works' uitype='account'>{67CD0332-310B-E911-A94D-000D3AF0677F}</value>
                    <value uiname='Purvi Electronics' uitype='account'>{2501C57A-91B2-EA11-A812-000D3AF05A4B}</value>
                    <value uiname='RAAJ COMUNICATION (ASP)' uitype='account'>{29C755C6-3A0D-E911-A94E-000D3AF06CD4}</value>
                    <value uiname='Unique Solutions' uitype='account'>{BD5E18C1-300B-E911-A94E-000D3AF06A98}</value>
 	                <value uiname='JAHANVI ENTERPRISES (ASP)' uitype='account'>{A58F33E6-300B-E911-A94E-000D3AF06A98}</value>
                    <value uiname='Kamboj Enterprises' uitype='account'>{6DADAFF9-2301-EA11-A811-000D3AF057DD}</value>
                    <value uiname='Kavya Refrigeration' uitype='account'>{F8D8461D-1652-EB11-A812-0022486E3BAA}</value>
                    <value uiname='K K Enterprises' uitype='account'>{A8359CF3-300B-E911-A94D-000D3AF0677F}</value>
                    <value uiname='K.l.service Point' uitype='account'>{C309C00B-89C9-EB11-BACC-6045BD7298AD}</value>
                    <value uiname='Krishna Sales' uitype='account'>{AF2F1CF3-300B-E911-A94E-000D3AF06CD4}</value>
                    <value uiname='MA PITHAD SERVICES-ASP' uitype='account'>{9D593B12-310B-E911-A94E-000D3AF06091}</value>
                    <value uiname='Mahi Enterprises' uitype='account'>{D4A747CF-E8DD-E911-A812-000D3AF057DD}</value>
                    <value uiname='Modern Electronics' uitype='account'>{71EC6720-310B-E911-A94E-000D3AF06091}</value>
                    <value uiname='M. S. Enterprises' uitype='account'>{18A64148-5BC9-EA11-A812-000D3AF055B6}</value>
                    <value uiname='M.s. Electronics' uitype='account'>{1FC6FDB1-B7B0-EC11-9840-6045BDAA92C3}</value>
                    <value uiname='New Bhagwati Electricals' uitype='account'>{531FB438-310B-E911-A94E-000D3AF06CD4}</value>
                    <value uiname='Pi-Tech' uitype='account'>{DE015054-310B-E911-A94E-000D3AF06CD4}</value>
                    <value uiname='POONAM ELECTRONICS-ASP' uitype='account'>{BD6AC05D-310B-E911-A94E-000D3AF06091}</value>
                    <value uiname='Prathvi Enterprises' uitype='account'>{D5A7ACA5-AF38-EC11-8C64-000D3AF07EE3}</value>
                    <value uiname='Preet Electric Works' uitype='account'>{C21F825A-310B-E911-A94E-000D3AF06CD4}</value>
                    <value uiname='PRO SERVICES (ASP)' uitype='account'>{CE6AC05D-310B-E911-A94E-000D3AF06091}</value>
                    <value uiname='Prakash Refrigeration And Elec' uitype='account'>{0CD4B31D-8E39-EA11-A813-000D3AF057DD}</value>
                    <value uiname='Quick Service Network' uitype='account'>{7A3BCE64-310B-E911-A94E-000D3AF06091}</value>
                    <value uiname='Quick Solution Point' uitype='account'>{4BD04A73-8D22-EC11-B6E6-6045BD731341}</value>
                    <value uiname='Red Star' uitype='account'>{42A41452-0B11-EA11-A811-000D3AF057DD}</value>
                    <value uiname='Ruby Electric Works' uitype='account'>{4D51D33F-F8EB-E911-A812-000D3AF058B5}</value>
                    <value uiname='Rukmini Enterprises' uitype='account'>{7B07E6A6-FE99-EC11-B401-6045BDAAFC34}</value>
                    <value uiname='Sagar Enterprises' uitype='account'>{316975A7-FE99-EC11-B401-6045BDAAF7C7}</value>
                    <value uiname='Sai Purvi Enterprises' uitype='account'>{848960A0-310B-E911-A94D-000D3AF0677F}</value>
                    <value uiname='Saini Ref. And Air Cond. Servi' uitype='account'>{538F05B5-B347-EA11-A812-000D3AF057DD}</value>
                    <value uiname='Sarotia Enterprises' uitype='account'>{0A57EC3D-FC4D-EA11-A812-000D3AF05D7B}</value>
                    <value uiname='Sai Ultimate Cooling Solutions' uitype='account'>{CA262794-310B-E911-A94E-000D3AF06CD4}</value>
                    <value uiname='Serv Simplified Appliances Pvt' uitype='account'>{B358886C-7B32-EA11-A813-000D3AF05D7B}</value>
                    <value uiname='Skyline Electronics' uitype='account'>{B7042FA1-D7B4-E911-A95B-000D3AF0694E}</value>
                    <value uiname='SKY Refrigeration' uitype='account'>{3F86ECB2-B347-EA11-A812-000D3AF055B6}</value>
                    <value uiname='Smart Care Solutions' uitype='account'>{3660EFCB-310B-E911-A94D-000D3AF06C56}</value>
                    <value uiname='Smart Electronics' uitype='account'>{EDDD998A-FC58-EA11-A811-000D3AF05F57}</value>
                    <value uiname='S.P. Enterprise' uitype='account'>{C8A4F6CC-310B-E911-A94E-000D3AF06A98}</value>
                    <value uiname='Sri Sreenivaasa Services Cente' uitype='account'>{2B0767B4-BD91-EB11-B1AC-0022486E7094}</value>
                    <value uiname='STAR REFRIGERATION (ASP)' uitype='account'>{63053BEE-310B-E911-A94D-000D3AF0694E}</value>
                    <value uiname='Technotronic Enterprises' uitype='account'>{8C5A3701-320B-E911-A94E-000D3AF06A98}</value>
                    <value uiname='Ugam Refrigeration' uitype='account'>{73C6D114-8DF9-EA11-A815-000D3AF055B6}</value>
                    <value uiname='Unique Service Centre' uitype='account'>{61E25E6A-EA00-EC11-94EF-6045BD72E02A}</value>
                    <value uiname='VAHANVATI AIR (ASP)' uitype='account'>{B376ADBE-300B-E911-A94D-000D3AF06C56}</value>
                    <value uiname='Vansh Enterprises' uitype='account'>{456BA423-B558-EC11-8F8F-002248D4E579}</value>
                  </condition>
                </filter>
                <link-entity name='account' from='accountid' to='hil_owneraccount' link-type='inner' alias='ae'>
                    <filter type='and'>
                        <condition attribute='hil_claimprocessed' operator='ne' value='1' />
                    </filter>
                </link-entity>
                </entity>
                </fetch>";

                entcoll = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                int i = 1;
                Entity entCustomer = null;
                foreach (Entity entASP in entcoll.Entities)
                {
                    _performaInvoiceId = Guid.Empty;
                    _erFranchise = entASP.GetAttributeValue<EntityReference>("hil_owneraccount");
                    _performaInvoiceId = GeneratePerformaInvoice(_service, _erFranchise.Id, _erFiscalMonth.Id);

                    #region Calculating TAT Incentive Slab
                    Console.WriteLine("Calculating TAT Incentive Slab... " + (i++).ToString() + "/" + entcoll.Entities.Count.ToString());
                    // Generate TAT Slab if all Claim status of all Work Order of that Franchisee not in (CLaim approved, Claim Rejected, Penalty Imposed, Penalty Waived)
                    queryExp = new QueryExpression(msdyn_workorder.EntityLogicalName);
                    queryExp.ColumnSet = new ColumnSet("msdyn_name", "hil_claimstatus", "msdyn_workorderid", "hil_tatachievementslab");
                    queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                    queryExp.Criteria.AddCondition("hil_claimstatus", ConditionOperator.NotIn, new object[] { 0, 4, 3, 8, 7 });
                    queryExp.Criteria.AddCondition("msdyn_substatus", ConditionOperator.Equal, new Guid("1727FA6C-FA0F-E911-A94E-000D3AF060A1"));
                    queryExp.Criteria.AddCondition("hil_owneraccount", ConditionOperator.Equal, entASP.GetAttributeValue<EntityReference>("hil_owneraccount").Id);
                    queryExp.Criteria.AddCondition("hil_isocr", ConditionOperator.NotEqual, true);
                    queryExp.Criteria.AddCondition(new ConditionExpression("msdyn_timeclosed", ConditionOperator.OnOrAfter, _startDate));
                    queryExp.Criteria.AddCondition(new ConditionExpression("msdyn_timeclosed", ConditionOperator.OnOrBefore, _endDate));
                    entcollTATSlabJobs = _service.RetrieveMultiple(queryExp);
                    if (entcollTATSlabJobs.Entities.Count == 0)
                    {
                        GenerateTATClaimLines(_service, _performaInvoiceId, _erFranchise.Id, Convert.ToDateTime(_startDate), Convert.ToDateTime(_endDate));
                    }
                    entCustomer = new Entity("account");
                    entCustomer.Id = entASP.GetAttributeValue<EntityReference>("hil_owneraccount").Id;
                    entCustomer["hil_claimprocessed"] = true;
                    _service.Update(entCustomer);
                    #endregion
                }
                #endregion
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        public static void GenerateTATIncentivesClaimASPWiseUAT(IOrganizationService _service)
        {
            #region Variable declaration
            Guid _fiscalMonthId = Guid.Empty;
            DateTime? _startDate = null;
            DateTime? _endDate = null;
            Guid _performaInvoiceId = Guid.Empty;
            string _fetchXML = string.Empty;
            QueryExpression queryExp;
            EntityCollection entcoll;
            EntityCollection entcollTATSlabJobs;
            EntityReference _erFiscalMonth = null;
            EntityReference _erFranchise = null;
            string[] _jobColumns = { "hil_fiscalmonth", "msdyn_name", "createdon", "hil_owneraccount", "hil_channelpartnercategory", "hil_warrantystatus", "hil_countryclassification", "hil_tatachievementslab", "hil_tatcategory", "hil_schemecode", "hil_salesoffice", "hil_productcategory", "hil_kkgcode", "hil_typeofassignee", "hil_brand", "hil_productsubcategory", "hil_callsubtype", "hil_claimstatus", "hil_isocr", "hil_laborinwarranty", "hil_isgascharged", "hil_sparepartuse", "hil_jobclosemobile", "msdyn_timeclosed", "msdyn_customerasset", "hil_quantity", "hil_purchasedate" };
            string _performaRemarks = string.Empty;
            #endregion

            #region Fetching Fiscal Month Date Range
            DateTime _processDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 20);
            queryExp = new QueryExpression("hil_claimperiod");
            queryExp.ColumnSet = new ColumnSet("hil_fromdate", "hil_todate");
            queryExp.Criteria = new FilterExpression(LogicalOperator.And);
            queryExp.Criteria.AddCondition("hil_isapplicable", ConditionOperator.Equal, true);
            queryExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
            entcoll = _service.RetrieveMultiple(queryExp);
            if (entcoll.Entities.Count > 0)
            {
                _startDate = entcoll.Entities[0].GetAttributeValue<DateTime>("hil_fromdate").AddMinutes(330);
                _endDate = entcoll.Entities[0].GetAttributeValue<DateTime>("hil_todate").AddMinutes(330);
                _erFiscalMonth = entcoll.Entities[0].ToEntityReference();
            }
            if (_erFiscalMonth == null)
            {
                Console.WriteLine("Fiscal Month is not defined." + DateTime.Now.ToString());
                return;
            }
            #endregion

            #region Refreshing Claim Parameters & Generating applicable Claim Lines
            try
            {
                //Channel Partner Code : 90cb15a6-7275-eb11-a812-000d3af05600
                // Sales Office Code:  df2f0cc1-bcf7-e811-a94c-000d3af0677f
                _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                <entity name='msdyn_workorder'>
                <attribute name='hil_owneraccount' />
                <filter type='and'>
                    <condition attribute='hil_fiscalmonth' operator='eq' value='" + _erFiscalMonth.Id + @"' />
                     <condition attribute='msdyn_workorderid' operator='in'>
                        <value>{1fcf2b67-0cb7-ec11-983f-6045bd7283d4}</value>
                        <value>{1f370a35-27ba-ec11-983f-6045bdad57ad}</value>
                        <value>{9c47945e-4aba-ec11-983f-6045bdad520f}</value>
                        <value>{14ea366f-4eba-ec11-983f-6045bdad5518}</value>
                        <value>{193ad4dc-57ba-ec11-983f-6045bdad5f85}</value>
                        <value>{bdfdb1b3-d6ba-ec11-983f-6045bdad5f2b}</value>
                        <value>{0b8ffa25-f1ba-ec11-983f-6045bdad5f2b}</value>
                        <value>{81d4dee5-b2bb-ec11-983f-6045bdad50d5}</value>
                        <value>{e900dea3-b8bb-ec11-983f-6045bdad5671}</value>
                        <value>{3725f7c6-c7bb-ec11-983f-6045bdad5079}</value>
                        <value>{2f1d66e7-7bbc-ec11-983f-6045bdad5446}</value>
                        <value>{0ef532ef-8cbc-ec11-983f-6045bdad5446}</value>
                        <value>{25e4801b-a4bc-ec11-983f-6045bdad5446}</value>
                        <value>{001d58ab-b6bc-ec11-983f-6045bdad50a5}</value>
                        <value>{3d050abb-42bd-ec11-983f-6045bdad5518}</value>
                        <value>{c7271d4c-68bd-ec11-983f-6045bdad5671}</value>
                        <value>{a809e3cc-74bd-ec11-9840-6045bdad5c32}</value>
                        <value>{6bb213ad-75bd-ec11-983f-6045bdad50d5}</value>
                        <value>{14da81b3-d9be-ec11-983e-0022486e5a4c}</value>
                        <value>{350167d6-e9bf-ec11-983e-6045bda58441}</value>
                        <value>{528c50d2-67c0-ec11-983e-6045bdaa8ed8}</value>
                            </condition>
                </filter>
                </entity>
                </fetch>";
                
                entcoll = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                int i = 1;
                Entity entCustomer = null;
                foreach (Entity entASP in entcoll.Entities)
                {
                    _performaInvoiceId = Guid.Empty;
                    _erFranchise = entASP.GetAttributeValue<EntityReference>("hil_owneraccount");
                    _performaInvoiceId = GeneratePerformaInvoice(_service, _erFranchise.Id, _erFiscalMonth.Id);

                    #region Calculating TAT Incentive Slab
                    Console.WriteLine("Calculating TAT Incentive Slab... " + (i++).ToString() + "/" + entcoll.Entities.Count.ToString());
                    // Generate TAT Slab if all Claim status of all Work Order of that Franchisee not in (CLaim approved, Claim Rejected, Penalty Imposed, Penalty Waived)
                    queryExp = new QueryExpression(msdyn_workorder.EntityLogicalName);
                    queryExp.ColumnSet = new ColumnSet("msdyn_name", "hil_claimstatus", "msdyn_workorderid", "hil_tatachievementslab");
                    queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                    queryExp.Criteria.AddCondition("hil_claimstatus", ConditionOperator.NotIn, new object[] { 0, 4, 3, 8, 7 });
                    queryExp.Criteria.AddCondition("msdyn_substatus", ConditionOperator.Equal, new Guid("1727FA6C-FA0F-E911-A94E-000D3AF060A1"));
                    queryExp.Criteria.AddCondition("hil_owneraccount", ConditionOperator.Equal, entASP.GetAttributeValue<EntityReference>("hil_owneraccount").Id);
                    queryExp.Criteria.AddCondition("hil_isocr", ConditionOperator.NotEqual, true);
                    queryExp.Criteria.AddCondition(new ConditionExpression("msdyn_timeclosed", ConditionOperator.OnOrAfter, _startDate));
                    queryExp.Criteria.AddCondition(new ConditionExpression("msdyn_timeclosed", ConditionOperator.OnOrBefore, _endDate));
                    entcollTATSlabJobs = _service.RetrieveMultiple(queryExp);
                    if (entcollTATSlabJobs.Entities.Count == 0)
                    {
                        GenerateTATClaimLines(_service, _performaInvoiceId, _erFranchise.Id, Convert.ToDateTime(_startDate), Convert.ToDateTime(_endDate));
                    }
                    entCustomer = new Entity("account");
                    entCustomer.Id = entASP.GetAttributeValue<EntityReference>("hil_owneraccount").Id;
                    entCustomer["hil_claimprocessed"] = true;
                    _service.Update(entCustomer);
                    #endregion
                }
                #endregion
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        public static void UpdatePerformaInvoiceSummary(IOrganizationService _service)
        {

            #region Variable declaration
            Guid _fiscalMonthId = new Guid("DC417E88-E62B-EB11-A813-0022486E5C9D");
            //Guid _fiscalMonthId = Guid.Empty;
            DateTime? _startDate = null;
            DateTime? _endDate = null;
            string _fetchXML = string.Empty;
            QueryExpression queryExp;
            EntityCollection entcoll;
            EntityReference _erFiscalMonth = null;
            
            #endregion

            #region Fetching Fiscal Month Date Range
            //DateTime _processDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);
            queryExp = new QueryExpression("hil_claimperiod");
            queryExp.ColumnSet = new ColumnSet("hil_fromdate", "hil_todate");
            queryExp.Criteria = new FilterExpression(LogicalOperator.And);
            queryExp.Criteria.AddCondition("hil_isapplicable", ConditionOperator.Equal, true);
            queryExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
            entcoll = _service.RetrieveMultiple(queryExp);
            if (entcoll.Entities.Count > 0)
            {
                _startDate = entcoll.Entities[0].GetAttributeValue<DateTime>("hil_fromdate").AddMinutes(330);
                _endDate = entcoll.Entities[0].GetAttributeValue<DateTime>("hil_todate").AddMinutes(330);
                _erFiscalMonth = entcoll.Entities[0].ToEntityReference();
            }
            if (_erFiscalMonth == null)
            {
                Console.WriteLine("Fiscal Month is not defined." + DateTime.Now.ToString());
                return;
            }
            #endregion
            _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
              <entity name='hil_claimheader'>
                <attribute name='hil_claimheaderid' />
                <attribute name='hil_name' />
                <order attribute='createdon' descending='false' />
                <filter type='and'>
                  <condition attribute='hil_fiscalmonth' operator='eq' value='{_erFiscalMonth.Id}' />
                  <filter type='or'>
                    <condition attribute='hil_branch' operator='null' />
                    <condition attribute='hil_totaljobsclosed' operator='null' />
                    <condition attribute='hil_totalclaimamount' operator='null' />
                    <condition attribute='hil_totalclaimvalue' operator='null' />
                    <condition attribute='hil_fixedcharges' operator='null' />
                  </filter>
                </filter>
                <link-entity name='account' from='accountid' to='hil_franchisee' visible='false' link-type='outer' alias='cp'>
                    <attribute name='hil_salesoffice' />
                </link-entity>
              </entity>
            </fetch>";
            entcoll = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
            if (entcoll.Entities.Count > 0)
            {
                EntityReference _SalesOffice = null;
                int _totalClaimableJobs = 0;
                decimal _JobClaimAmount = 0;
                decimal _totalClaimAmount = 0;
                decimal _fixedCharge = 0;
                decimal _expenseOverheadAmount = 0;
                foreach (Entity entJob in entcoll.Entities)
                {


                }
            }
            else { 
            
            }
        }
        public static void ProcessClaimAmountSum(IOrganizationService _service)
        {

            #region Variable declaration
            Guid _fiscalMonthId = new Guid("DC417E88-E62B-EB11-A813-0022486E5C9D");
            //Guid _fiscalMonthId = Guid.Empty;
            DateTime? _startDate = null;
            DateTime? _endDate = null;
            Guid _performaInvoiceId = Guid.Empty;
            int _totalClosedJobs = 0;
            string _fetchXML = string.Empty;
            QueryExpression queryExp;
            EntityCollection entcoll;
            EntityCollection entcollJobs;
            EntityReference _erFiscalMonth = null;
            string[] _jobColumns = { "hil_fiscalmonth", "msdyn_name", "createdon", "hil_owneraccount", "hil_channelpartnercategory", "hil_warrantystatus", "hil_countryclassification", "hil_tatachievementslab", "hil_tatcategory", "hil_schemecode", "hil_salesoffice", "hil_productcategory", "hil_kkgcode", "hil_typeofassignee", "hil_brand", "hil_productsubcategory", "hil_callsubtype", "hil_claimstatus", "hil_isocr", "hil_laborinwarranty", "hil_isgascharged", "hil_sparepartuse", "hil_jobclosemobile", "msdyn_timeclosed", "msdyn_customerasset", "hil_quantity", "hil_purchasedate" };
            Entity entUpdate;
            string _performaRemarks = string.Empty;
            #endregion

            #region Fetching Fiscal Month Date Range
            if (_fiscalMonthId != Guid.Empty)
            {
                Entity ent = _service.Retrieve("hil_claimperiod", _fiscalMonthId, new ColumnSet("hil_fromdate", "hil_todate"));
                if (ent != null)
                {
                    _startDate = ent.GetAttributeValue<DateTime>("hil_fromdate").AddMinutes(330);
                    _endDate = ent.GetAttributeValue<DateTime>("hil_todate").AddMinutes(330);
                    _erFiscalMonth = ent.ToEntityReference();
                }
            }
            else
            {

                DateTime _processDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);

                queryExp = new QueryExpression("hil_claimperiod");
                queryExp.ColumnSet = new ColumnSet("hil_fromdate", "hil_todate");
                queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                queryExp.Criteria.AddCondition("hil_fromdate", ConditionOperator.OnOrBefore, _processDate);
                queryExp.Criteria.AddCondition("hil_todate", ConditionOperator.OnOrAfter, _processDate);
                entcoll = _service.RetrieveMultiple(queryExp);
                if (entcoll.Entities.Count > 0)
                {
                    _startDate = entcoll.Entities[0].GetAttributeValue<DateTime>("hil_fromdate").AddMinutes(330);
                    _endDate = entcoll.Entities[0].GetAttributeValue<DateTime>("hil_todate").AddMinutes(330);
                    _erFiscalMonth = entcoll.Entities[0].ToEntityReference();
                }
            }
            #endregion

            #region Refreshing Claim Parameters & Generating applicable Claim Lines
            try
            {
                _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                <entity name='msdyn_workorder'>
                <attribute name='msdyn_name' />
                <attribute name='createdon' />
                <attribute name='hil_productsubcategory' />
                <attribute name='hil_customerref' />
                <attribute name='hil_callsubtype' />
                <attribute name='msdyn_workorderid' />
                <order attribute='msdyn_timeclosed' descending='false' />
                <filter type='and'>
                    <condition attribute='msdyn_timeclosed' operator='on-or-after' value='2020-09-21' />
                    <condition attribute='msdyn_timeclosed' operator='on-or-before' value='2020-10-20' />
                    <condition attribute='hil_generateclaim' operator='eq' value='1' />
                    <filter type='or'>
                    <condition attribute='hil_claimamount' operator='eq' value='0' />
                    <condition attribute='hil_claimamount' operator='null' />
                    </filter>
                    <condition attribute='msdyn_substatus' operator='eq' value='{1727FA6C-FA0F-E911-A94E-000D3AF060A1}' />
                    <condition attribute='hil_typeofassignee' operator='in'>
                    <value uiname='Franchise' uitype='position'>{4A1AA189-1208-E911-A94D-000D3AF0694E}</value>
                    <value uiname='Franchise Technician' uitype='position'>{0197EA9B-1208-E911-A94D-000D3AF0694E}</value>
                    </condition>
                    <condition attribute='hil_isocr' operator='ne' value='1' />
                </filter>
                    <link-entity name='hil_claimline' from='hil_jobid' to='msdyn_workorderid' link-type='inner' alias='ab' />
                </entity>
                </fetch>";
                try
                {
                    while (true)
                    {
                        entcollJobs = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                        if (entcollJobs.Entities.Count == 0) { break; }
                        _totalClosedJobs = entcollJobs.Entities.Count;
                        int i = 0;
                        foreach (Entity entJob in entcollJobs.Entities)
                        {
                            Console.WriteLine("Calculating Claim amount");
                            _fetchXML = @"<fetch distinct='false' mapping='logical' aggregate='true'>
                            <entity name='hil_claimline'>
                            <attribute name='hil_claimamount' alias='claimamount' aggregate='sum'/> 
                            <attribute name='hil_jobid' alias='jobid' groupby='true' /> 
                            <filter type='and'>
                            <filter type='or'>
                                <condition attribute='hil_jobid' operator='eq' value='{" + entJob.Id + @"}' />
                            </filter>
                            </filter>
                            </entity>
                            </fetch>";
                            Decimal _JobClaimAmount = 0;
                            EntityCollection entcollCA = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                            if (entcollCA.Entities.Count > 0)
                            {
                                Guid _claimJobId = ((EntityReference)((AliasedValue)entcollCA.Entities[0].Attributes["jobid"]).Value).Id;
                                _JobClaimAmount = ((Money)((AliasedValue)entcollCA.Entities[0]["claimamount"]).Value).Value;
                                if (_claimJobId != null)
                                {
                                    entUpdate = new Entity("msdyn_workorder", _claimJobId);
                                    entUpdate["hil_generateclaim"] = true;
                                    entUpdate["hil_claimamount"] = new Money(_JobClaimAmount);
                                    _service.Update(entUpdate);
                                }
                            }
                            else
                            {
                                entUpdate = new Entity("msdyn_workorder", entJob.Id);
                                entUpdate["hil_generateclaim"] = true;
                                entUpdate["hil_claimamount"] = new Money(0);
                                _service.Update(entUpdate);
                            }
                            Console.WriteLine(i++.ToString() + "/" + _totalClosedJobs.ToString());
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
                #endregion

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        private static decimal CreateFixedChargeClaimLine(IOrganizationService _service, Guid _performInvoiceId, string _claimCatwgory, EntityReference _erClaimPeriod, EntityReference _erChannelPartner, DateTime _endDate)
        {
            Decimal _fixedCharge = 0;
            QueryExpression queryExp;
            EntityCollection entcoll;
            try
            {
                if (!ClaimLineExists(_service, null, _erChannelPartner.Id, _erClaimPeriod.Id, _claimCatwgory))
                {
                    queryExp = new QueryExpression("hil_fixedcompensation");
                    queryExp.ColumnSet = new ColumnSet("hil_amount");
                    queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                    queryExp.Criteria.AddCondition("hil_channelpartnercode", ConditionOperator.Equal, _erChannelPartner.Id);
                    queryExp.Criteria.AddCondition("hil_startdate", ConditionOperator.OnOrBefore, _endDate);
                    queryExp.Criteria.AddCondition("hil_enddate", ConditionOperator.OnOrAfter, _endDate);
                    entcoll = _service.RetrieveMultiple(queryExp);
                    if (entcoll.Entities.Count > 0)
                    {
                        _fixedCharge = entcoll.Entities[0].GetAttributeValue<Money>("hil_amount").Value;

                        Entity entObj = new Entity("hil_claimline");
                        Guid _claimLineId = Guid.Empty;
                        entObj["hil_franchisee"] = _erChannelPartner;
                        entObj["hil_claimperiod"] = _erClaimPeriod;
                        entObj["hil_claimcategory"] = new EntityReference("hil_claimcategory", new Guid(_claimCatwgory));
                        entObj["hil_claimheader"] = new EntityReference("hil_claimheader", _performInvoiceId);
                        entObj["hil_claimamount"] = new Money(_fixedCharge);
                        _claimLineId = _service.Create(entObj);
                        Entity ent = _service.Retrieve(Account.EntityLogicalName, _erChannelPartner.Id, new ColumnSet("ownerid"));
                        if (ent != null)
                        {
                            AssignRecord("hil_claimline", ent.GetAttributeValue<EntityReference>("ownerid").Id, _claimLineId, _service);
                        }

                    }
                }
                else
                {
                    queryExp = new QueryExpression("hil_claimline");
                    queryExp.ColumnSet = new ColumnSet("hil_claimamount");
                    queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                    queryExp.Criteria.AddCondition("hil_claimheader", ConditionOperator.Equal, _performInvoiceId);
                    queryExp.Criteria.AddCondition("hil_claimcategory", ConditionOperator.Equal, new Guid(_claimCatwgory));
                    queryExp.Criteria.AddCondition("hil_claimperiod", ConditionOperator.Equal, _erClaimPeriod.Id);
                    queryExp.Criteria.AddCondition("hil_franchisee", ConditionOperator.Equal, _erChannelPartner.Id);
                    entcoll = _service.RetrieveMultiple(queryExp);
                    if (entcoll.Entities.Count > 0)
                    {
                        _fixedCharge = entcoll.Entities[0].GetAttributeValue<Money>("hil_claimamount").Value;
                    }
                }
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
            return _fixedCharge;
        }
        public static void RejectPendingSAWApprovals(IOrganizationService _service)
        {
            QueryExpression queryExp;
            EntityCollection entCol;
            //queryExp = new QueryExpression("hil_sawactivityapproval");
            //queryExp.ColumnSet = new ColumnSet("hil_approvalstatus", "hil_approverremarks");
            //queryExp.Criteria = new FilterExpression(LogicalOperator.And);
            //queryExp.Criteria.AddCondition("createdon", ConditionOperator.OnOrBefore, "2022-06-20");
            //queryExp.Criteria.AddCondition("hil_approvalstatus", ConditionOperator.In, new object[] { 1, 2, 5, 6 });
            //<condition attribute='hil_approver' operator='eq' value='{17B37EC0-93A7-E911-A97A-000D3AF063A5}' />

            string _fetxhXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
            <entity name='hil_sawactivityapproval'>
            <attribute name='hil_approvalstatus' />
            <attribute name='hil_approverremarks' />
            <order attribute='createdon' descending='false' />
            <filter type='and'>
                <condition attribute='hil_approvalstatus' operator='in'>
                <value>1</value>
                <value>2</value>
                <value>5</value>
                <value>6</value>
                </condition>
                <condition attribute='createdon' operator='on-or-before' value='2024-05-20' />
            </filter> 
            <link-entity name='msdyn_workorder' from='msdyn_workorderid' to='hil_jobid' link-type='inner' alias='job'>
                <attribute name='msdyn_substatus' />
                <filter type='and'>
                <condition attribute='msdyn_substatus' operator='in'>
                    <value uiname='Closed' uitype='msdyn_workordersubstatus'>{1727FA6C-FA0F-E911-A94E-000D3AF060A1}</value>
                    <value uiname='KKG Audit Failed' uitype='msdyn_workordersubstatus'>{6C8F2123-5106-EA11-A811-000D3AF057DD}</value>
                </condition>
                </filter>
            </link-entity>
            </entity>
            </fetch>";
            entCol = _service.RetrieveMultiple(new FetchExpression(_fetxhXML));
            int i = 1;
            //Disable SAW Activity Approval PReUpdate Step
            foreach (Entity ent in entCol.Entities)
            {
                try
                {
                    ent["hil_approverremarks"] = "Due to non-review by Branch, SAW & Claim has been rejected for this Job. Please check with your Branch Service Head for the same.";
                    ent["hil_approvalstatus"] = new OptionSetValue(4);
                    _service.Update(ent);
                    Console.WriteLine("Processing..." + i++.ToString() + "/" + entCol.Entities.Count);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Processing..." + i++.ToString() + "/" + entCol.Entities.Count + ex.Message);
                }
                //ent["hil_approver"] = new EntityReference("systemuser", new Guid("5190416c-0782-e911-a959-000d3af06a98"));
                //_service.Update(ent);
                
            }
        }
        public static void DeleteDuplicateClaimLines(IOrganizationService service)
        {
            string filePath = @"C:\Kuldeep khare\DuplicateClaimLine.xlsx";
            string conn = string.Empty;
            DataTable dtexcel = new DataTable();

            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            if (excelApp != null)
            {
                Microsoft.Office.Interop.Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(filePath, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Microsoft.Office.Interop.Excel.Worksheet excelWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)excelWorkbook.Sheets[1];

                Microsoft.Office.Interop.Excel.Range excelRange = excelWorksheet.UsedRange;
                Microsoft.Office.Interop.Excel.Range range;
                string claimheader, claimcategory, jobid;
                QueryExpression Query1;
                EntityCollection entcoll1;

                for (int i = 2; i <= excelRange.Rows.Count; i++)
                {
                    try
                    {
                        Console.WriteLine(i.ToString() + "/" + excelRange.Rows.Count.ToString());
                        range = (excelWorksheet.Cells[i, 1] as Microsoft.Office.Interop.Excel.Range);
                        claimheader = range.Value.ToString();
                        range = (excelWorksheet.Cells[i, 2] as Microsoft.Office.Interop.Excel.Range);
                        claimcategory = range.Value.ToString();
                        range = (excelWorksheet.Cells[i, 3] as Microsoft.Office.Interop.Excel.Range);
                        if (range.Value != null)
                        {
                            jobid = range.Value.ToString();
                            Query1 = new QueryExpression("hil_claimline");
                            Query1.ColumnSet = new ColumnSet("hil_name", "hil_claimheader", "hil_claimcategory", "hil_jobid", "statecode", "statuscode");
                            Query1.Criteria = new FilterExpression(LogicalOperator.And);
                            Query1.Criteria.AddCondition("hil_claimheader", ConditionOperator.Equal, new Guid(claimheader));
                            Query1.Criteria.AddCondition("hil_claimcategory", ConditionOperator.Equal, new Guid(claimcategory));
                            Query1.Criteria.AddCondition("hil_jobid", ConditionOperator.Equal, new Guid(jobid));
                            Query1.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                            Query1.AddOrder("createdon", OrderType.Descending);
                            entcoll1 = service.RetrieveMultiple(Query1);
                            if (entcoll1.Entities.Count > 1)
                            {
                                Entity ent = entcoll1.Entities[0];
                                Console.WriteLine("Found.. Claim Header# " + ent.GetAttributeValue<EntityReference>("hil_claimheader").Name + " Claimcategory# " + ent.GetAttributeValue<EntityReference>("hil_claimcategory").Name + " Job# " + ent.GetAttributeValue<EntityReference>("hil_jobid").Name);
                                ent["hil_name"] = "Found Duplicate";
                                ent["statecode"] = new OptionSetValue(1);
                                ent["statuscode"] = new OptionSetValue(2);
                                service.Update(ent);
                            }
                        }
                        Console.WriteLine(i.ToString() + "/" + excelRange.Rows.Count.ToString());
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(i.ToString() + "/" + excelRange.Rows.Count.ToString() + ex.Message);
                    }
                }
            }
        }

        public static void DeleteDuplicateInvoices(IOrganizationService service)
        {
            string filePath = @"C:\Kuldeep khare\AMCInvoice.xlsx";
            string conn = string.Empty;
            DataTable dtexcel = new DataTable();

            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            if (excelApp != null)
            {
                Microsoft.Office.Interop.Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(filePath, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Microsoft.Office.Interop.Excel.Worksheet excelWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)excelWorkbook.Sheets[1];

                Microsoft.Office.Interop.Excel.Range excelRange = excelWorksheet.UsedRange;
                Microsoft.Office.Interop.Excel.Range range;
                string _invoiceid;
                QueryExpression Query1;
                EntityCollection entcoll1;

                for (int i = 1; i <= excelRange.Rows.Count; i++)
                {
                    try
                    {
                        Console.WriteLine(i.ToString() + "/" + excelRange.Rows.Count.ToString());
                        range = (excelWorksheet.Cells[i, 1] as Microsoft.Office.Interop.Excel.Range);
                        _invoiceid = range.Value.ToString();
                        if (_invoiceid != null)
                        {
                            string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                                <entity name='invoice'>
                                <attribute name='name' />
                                <attribute name='invoiceid' />
                                <attribute name='customerid' />
                                <filter type='and'>
                                    <condition attribute='name' operator='eq' value='{_invoiceid}' />
                                </filter>
                                <link-entity name='msdyn_paymentdetail' from='msdyn_invoice' to='invoiceid' link-type='inner' alias='aj'>
                                    <filter type='and'>
                                    <condition attribute='statuscode' operator='ne' value='910590000' />
                                    </filter>
                                </link-entity>
                                </entity>
                                </fetch>";
                            entcoll1 = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                            Console.WriteLine("Found: "+ entcoll1.Entities.Count);
                            int j = 65;
                            foreach (Entity ent in entcoll1.Entities)
                            {
                                Entity _entUpdate = new Entity(ent.LogicalName, ent.Id);
                                _entUpdate["name"] = ent.GetAttributeValue<string>("name") + (Char)(j++);
                                service.Update(_entUpdate);
                                Console.WriteLine("Invoice# " + ent.GetAttributeValue<string>("name"));
                            }
                        }
                        Console.WriteLine(i.ToString() + "/" + excelRange.Rows.Count.ToString());
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(i.ToString() + "/" + excelRange.Rows.Count.ToString() + ex.Message);
                    }
                }
            }
        }
    }
    public class ClaimCategoryConst
    {
        public const string _AbnormalClosure = "8dbfb0c4-33f3-ea11-a815-000d3af05a4b";
        public const string _BaseCallCharge = "864a1bcc-6ee5-ea11-a817-000d3af0501c";
        public const string _FixedCompensation = "824a1bcc-6ee5-ea11-a817-000d3af0501c";
        public const string _KKGAuditFailed = "944a1bcc-6ee5-ea11-a817-000d3af0501c";
        public const string _KKGClosure = "acfd1d05-33f3-ea11-a815-000d3af057dd";
        public const string _LocalPurchase = "39ab9283-c7e5-ea11-a817-000d3af05f57";
        public const string _LocaRepair = "ae26c489-c7e5-ea11-a817-000d3af05f57";
        public const string _MinimumGuarantee = "844a1bcc-6ee5-ea11-a817-000d3af0501c";
        public const string _MobileClosureIncentive = "8a4a1bcc-6ee5-ea11-a817-000d3af0501c";
        public const string _ProductTransport = "a0ae2d9c-c7e5-ea11-a817-000d3af05f57";
        public const string _SchemeIncentive = "904a1bcc-6ee5-ea11-a817-000d3af0501c";
        public const string _SpecialIncentive = "8e4a1bcc-6ee5-ea11-a817-000d3af0501c";
        public const string _TATBreachPenalty = "924a1bcc-6ee5-ea11-a817-000d3af0501c";
        public const string _TATIncentive = "8c4a1bcc-6ee5-ea11-a817-000d3af0501c";
        public const string _UpcountryTravelCharge = "884a1bcc-6ee5-ea11-a817-000d3af0501c";
    }
    public class ClaimLineDTO
    {
        public Guid OwnerId { get; set; }
        public EntityReference PerformInvoiceId { get; set; }
        public decimal ClaimAmount { get; set; }
        public string ClaimCatwgory { get; set; }
        public EntityReference ClaimPeriod { get; set; }
        public EntityReference ChannelPartner { get; set; }
        public EntityReference ProdCatg { get; set; }
        public EntityReference CallSubtype { get; set; }
        public string ActivityCode { get; set; }
    }
    public class ClaimLineActivityData
    {
        public string ActivityCode { get; set; }
        public EntityReference ProdCategory { get; set; }

        public EntityReference CallSubtype { get; set; }
    }
}
