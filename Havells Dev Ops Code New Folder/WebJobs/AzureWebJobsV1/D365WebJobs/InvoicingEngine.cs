using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace D365WebJobs
{
    public class InvoicingEngine
    {
        public ReturnObject CreateDepositInvoice(IOrganizationService service, Entity entity, string[] bookingIds)
        {
            ReturnObject _retObject = new ReturnObject();

            if (entity.Contains("ogre_depositapplicable") && entity.GetAttributeValue<bool>("ogre_depositapplicable"))
            {
                QueryExpression query = new QueryExpression("ogre_rentalbooking");
                query.ColumnSet = new ColumnSet(true);
                query.Criteria = new FilterExpression(LogicalOperator.Or);
                foreach (string id in bookingIds)
                    query.Criteria.AddCondition("ogre_rentalbookingid", ConditionOperator.Equal, new Guid(id));

                EntityCollection bookingsColl = service.RetrieveMultiple(query);

                Entity InvoiceEnt = new Entity("ogre_rentalinvoice");
                InvoiceEnt["ogre_invoicetype"] = new EntityReference("ogre_rentalinvoicetype", new Guid("6caf5955-f9a0-ee11-be37-0022481b6df4"));//Deposit Invoice
                InvoiceEnt["ogre_contract"] = entity.ToEntityReference();
                InvoiceEnt["ogre_customer"] = entity.GetAttributeValue<EntityReference>("ogre_account");
                InvoiceEnt["ogre_invoicedate"] = DateTime.Now;
                InvoiceEnt["ogre_invoicestatus"] = new EntityReference("ogre_invoicestatus", new Guid("a3fc6e8b-f9a0-ee11-be37-0022481b6df4"));//Invoice Status - Draft

                EntityReference InvoiceHeaderID = new EntityReference("ogre_rentalinvoice", service.Create(InvoiceEnt));

                Entity _contractHeader = service.Retrieve("ogre_contract", entity.Id, new ColumnSet("ogre_depositchargeunitquantity", "ogre_depositchargeunit", "ogre_depositamount"));
                string _fetchXMLConttract = $@"";
                string _invoiceLineId = service.Retrieve("ogre_rentalinvoice", InvoiceHeaderID.Id, new ColumnSet("ogre_name")).Attributes["ogre_name"] + InvoiceHeaderID.Name + "-";
                int _rowIndex = 1;
                Entity _entBookingUpdate = null;
                foreach (Entity bookingEnt in bookingsColl.Entities)
                {
                    Entity InvoiceLine = new Entity("ogre_rentalinvoiceline");
                    InvoiceLine["ogre_chargetype"] = new EntityReference("ogre_product", new Guid("f5ccf9e1-e6a0-ee11-be37-0022481b6df4"));//Deposit Charge
                    InvoiceLine["ogre_quantity"] = _contractHeader.Contains("ogre_depositchargeunitquantity") ? _contractHeader.GetAttributeValue<Int32>("ogre_depositchargeunitquantity") : 0;
                    InvoiceLine["ogre_unit"] = _contractHeader.Contains("ogre_depositchargeunit") ? _contractHeader.GetAttributeValue<EntityReference>("ogre_depositchargeunit") : null;
                    InvoiceLine["ogre_rate"] = _contractHeader.Contains("ogre_depositamount") ? _contractHeader.GetAttributeValue<Money>("ogre_depositamount") : new Money(new decimal(0));
                    InvoiceLine["ogre_chargedescription"] = "Deposit against Booking# " + bookingEnt.GetAttributeValue<string>("ogre_bookingnumber");
                    InvoiceLine["ogre_taxtype"] = null;
                    InvoiceLine["ogre_taxper"] = new decimal(0.00);
                    InvoiceLine["ogre_invoicefrom"] = null;
                    InvoiceLine["ogre_invoiceuntil"] = null;
                    InvoiceLine["ogre_name"] = _invoiceLineId + _rowIndex++.ToString().PadLeft(3, '0');
                    InvoiceLine["ogre_rentalinvoice"] = InvoiceHeaderID;
                    InvoiceLine["ogre_rentalbooking"] = bookingEnt.ToEntityReference();
                    service.Create(InvoiceLine);

                    _entBookingUpdate = new Entity(bookingEnt.LogicalName, bookingEnt.Id);
                    _entBookingUpdate["ogre_depositinvoice"] = InvoiceHeaderID;
                    _entBookingUpdate["ogre_depositinvoiceraised"] = true;
                    _entBookingUpdate["ogre_depositinvoiceraisedon"] = DateTime.Now;
                    service.Update(_entBookingUpdate);
                }

                _retObject.Status = true;
                _retObject.StatusRemarks = "Done";
            }
            else
            {
                _retObject.Status = false;
                _retObject.StatusRemarks = "Deposit is not applicable.";
            }
            return _retObject;
        }
        public ReturnObject CreateInvoice(IOrganizationService _service, Entity entBooking)
        {
            ReturnObject _retObject = new ReturnObject();

            try
            {
                double _billableDays = 0;
                EntityReference _invoiceType = null;
                DateTime? _InvoiceFrom = null;
                DateTime? _InvoiceUntil = null;
                DateTime? _BillFrom = null;
                Money _chargeRate = new Money(new decimal(0));
                Decimal _billableAmount = new Decimal(0);
                Money _firstInvoiceAmount = new Money(new decimal(0));
                bool _firstInvoiceRaised = false;
                string _chargeDescfription = string.Empty;
                DateTime? _nextInvoiceDate = entBooking.Contains("ogre_nextinvoicedate") ? entBooking.GetAttributeValue<DateTime>("ogre_nextinvoicedate") : new DateTime();
                bool _firstInvoicePaymentMandatory = entBooking.Contains("ogre_firstrentalinvoicepaymentmandatory") ? entBooking.GetAttributeValue<bool>("ogre_firstrentalinvoicepaymentmandatory") : false;

                string _fetchXMLLine = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                <entity name='ogre_rentalbookingline'>
                <attribute name='ogre_rentalbookinglineid' />
                <attribute name='ogre_rate' />
                <attribute name='ogre_billingfrom' />
                <attribute name='ogre_firstinvoiceamount' />
                <order attribute='ogre_linenumber' descending='false' />
                <filter type='and'>
                    <condition attribute='ogre_rentalbooking' operator='eq' value='{entBooking.Id}' />
                    <condition attribute='ogre_chargetype' operator='eq' value='{{81A1B60F-2698-EE11-BE37-002248C6F70D}}' />
                    <condition attribute='statecode' operator='eq' value='0' />
                </filter>
                </entity>
                </fetch>";
                EntityCollection _entColLine = _service.RetrieveMultiple(new FetchExpression(_fetchXMLLine));
                if (_entColLine.Entities.Count > 0)
                {
                    _chargeRate = _entColLine.Entities[0].GetAttributeValue<Money>("ogre_rate");
                    _billableAmount = _chargeRate.Value;

                    if (_entColLine.Entities[0].Contains("ogre_firstinvoiceamount"))
                    {
                        _firstInvoiceAmount = _entColLine.Entities[0].GetAttributeValue<Money>("ogre_firstinvoiceamount");
                        if(_firstInvoiceAmount.Value>0 || _nextInvoiceDate.Value.Year!=1) _firstInvoiceRaised = true;
                    }
                    if (_entColLine.Entities[0].Contains("ogre_billingfrom"))
                    {
                        _BillFrom = _entColLine.Entities[0].GetAttributeValue<DateTime>("ogre_billingfrom");
                    }
                }

                if (_firstInvoicePaymentMandatory && !_firstInvoiceRaised && _nextInvoiceDate.Value.Year == 1)
                {
                    _chargeDescfription = "One Month Advance Rental";
                    if (_firstInvoiceRaised && _BillFrom == null)
                    {
                        _retObject.Status = false;
                        _retObject.StatusRemarks = "Please input Bill From.";
                        return _retObject;
                    }
                    _invoiceType = new EntityReference("ogre_rentalinvoicetype", new Guid("877c5062-f9a0-ee11-be37-0022481b6df4"));
                }
                else
                {
                    _invoiceType = new EntityReference("ogre_rentalinvoicetype", new Guid("63af5955-f9a0-ee11-be37-0022481b6df4"));

                    if (_BillFrom == null)
                    {
                        _retObject.Status = false;
                        _retObject.StatusRemarks = "No Invoice to produce";
                        return _retObject;
                    }
                    _chargeDescfription = "Rental against Booking# " + entBooking.GetAttributeValue<string>("ogre_bookingnumber");
                    if (_nextInvoiceDate == null || _nextInvoiceDate.Value.Year == 1)
                    {
                        _InvoiceFrom = _BillFrom;
                    }
                    else
                    {
                        _InvoiceFrom = _nextInvoiceDate;
                    }
                    _InvoiceUntil = new DateTime(_InvoiceFrom.Value.Year, _InvoiceFrom.Value.Month, DateTime.DaysInMonth(_InvoiceFrom.Value.Year, _InvoiceFrom.Value.Month));
                    if (_InvoiceFrom.Value.Day == 1)
                        _billableDays = 30;
                    else
                        _billableDays = Math.Round((_InvoiceUntil - _InvoiceFrom).Value.TotalDays + 1,0);
                    _billableAmount = Math.Round((Convert.ToDecimal(_billableDays) / 30) * _chargeRate.Value, 0);
                    if (_firstInvoiceAmount.Value != 0)//&& _nextInvoiceDate.Value.Year==1
                    {
                        _billableAmount = _firstInvoiceAmount.Value - _billableAmount;
                        if (_billableAmount > 0)
                        {
                            Entity _entLine = new Entity(_entColLine.Entities[0].LogicalName, _entColLine.Entities[0].Id);
                            _entLine["ogre_firstinvoiceamount"] = new Money(_billableAmount);
                            _service.Update(_entLine);
                            _entLine = new Entity(entBooking.LogicalName, entBooking.Id);
                            _entLine["ogre_nextinvoicedate"] = _InvoiceUntil.Value.AddDays(1);
                            _service.Update(_entLine);

                            _retObject.Status = false;
                            _retObject.StatusRemarks = "No Invoice to produce";
                            return _retObject;
                        }
                        else
                        {
                            Entity _entLine = new Entity(_entColLine.Entities[0].LogicalName, _entColLine.Entities[0].Id);
                            _entLine["ogre_firstinvoiceamount"] = new Money(new decimal(0));
                            _service.Update(_entLine);
                            _billableAmount = Math.Abs(_billableAmount);
                        }
                    }

                    _nextInvoiceDate = _InvoiceUntil.Value.AddDays(1);
                }

                Entity InvoiceEnt = new Entity("ogre_rentalinvoice");
                InvoiceEnt["ogre_invoicetype"] = _invoiceType;// new EntityReference("ogre_rentalinvoicetype", new Guid("63af5955-f9a0-ee11-be37-0022481b6df4"));//Rental Invoice
                if (entBooking.Contains("ogre_contract"))
                    InvoiceEnt["ogre_contract"] = entBooking.GetAttributeValue<EntityReference>("ogre_contract");
                if (entBooking.Contains("ogre_customer"))
                    InvoiceEnt["ogre_customer"] = entBooking.GetAttributeValue<EntityReference>("ogre_customer");
                InvoiceEnt["ogre_invoicedate"] = DateTime.Now;
                InvoiceEnt["ogre_booking"] = entBooking.ToEntityReference();

                InvoiceEnt["ogre_invoicestatus"] = new EntityReference("ogre_invoicestatus", new Guid("a3fc6e8b-f9a0-ee11-be37-0022481b6df4"));//Invoice Status - Draft
                EntityReference InvoiceHeaderID = new EntityReference("ogre_rentalinvoice", _service.Create(InvoiceEnt));
                string _fetchXMLBookingLines = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                <entity name='ogre_rentalbookingline'>
                <attribute name='ogre_billingfrom' />
                <attribute name='ogre_objectid' />
                <attribute name='ogre_unit' />
                <attribute name='ogre_rate' />
                <attribute name='ogre_invoicepattern' />
                <attribute name='ogre_contract' />
                <attribute name='ogre_chargetype' />
                <attribute name='ogre_status' />
                <attribute name='ogre_billinguntil' />
                <attribute name='ogre_billingtype' />
                <attribute name='ogre_billingperiod' />
                <attribute name='ogre_billingcycle' />
                <attribute name='ogre_quantity' />
                <attribute name='ogre_rentalperiod' />
                <attribute name='ogre_amount' />
                <attribute name='ogre_rentalbookinglineid' />
                <filter type='and'>
                <condition attribute='statecode' operator='eq' value='0' />
                <condition attribute='ogre_rentalbooking' operator='eq' value='{entBooking.Id}' />
                </filter>
                </entity>
                </fetch>";
                EntityCollection entColLines = _service.RetrieveMultiple(new FetchExpression(_fetchXMLBookingLines));
                foreach (Entity entLine in entColLines.Entities)
                {
                    if (entLine.Contains("ogre_chargetype"))
                    {
                        Entity _Product = _service.Retrieve(entLine.GetAttributeValue<EntityReference>("ogre_chargetype").LogicalName, entLine.GetAttributeValue<EntityReference>("ogre_chargetype").Id, new ColumnSet("ogre_taxgroupid"));
                        if (_Product.Contains("ogre_taxgroupid"))
                        {
                            Entity _Tax = _service.Retrieve(_Product.GetAttributeValue<EntityReference>("ogre_taxgroupid").LogicalName, _Product.GetAttributeValue<EntityReference>("ogre_taxgroupid").Id, new ColumnSet("ogre_percentage"));
                            Entity InvoiceLine = new Entity("ogre_rentalinvoiceline");
                            InvoiceLine["ogre_chargetype"] = _Product.ToEntityReference();
                            if (entLine.Contains("ogre_quantity"))
                                InvoiceLine["ogre_quantity"] = entLine.GetAttributeValue<int>("ogre_quantity");
                            if (entLine.Contains("ogre_unit"))
                                InvoiceLine["ogre_unit"] = entLine.GetAttributeValue<EntityReference>("ogre_unit");

                            InvoiceLine["ogre_rate"] = new Money(_billableAmount);

                            InvoiceLine["ogre_chargedescription"] = _chargeDescfription;// "Rental against Booking# " + entBooking.GetAttributeValue<string>("ogre_linenumber");
                            InvoiceLine["ogre_taxtype"] = _Tax.ToEntityReference();
                            InvoiceLine["ogre_taxper"] = _Tax.Contains("ogre_percentage") ? _Tax.GetAttributeValue<decimal>("ogre_percentage") : new decimal(0.00);
                            if (_firstInvoiceRaised)
                            {
                                InvoiceLine["ogre_invoicefrom"] = _InvoiceFrom.Value;
                                InvoiceLine["ogre_invoiceuntil"] = _InvoiceUntil.Value;
                            }
                            InvoiceLine["ogre_name"] = "INV-001";
                            InvoiceLine["ogre_rentalinvoice"] = InvoiceHeaderID;
                            InvoiceLine["ogre_rentalbooking"] = entBooking.ToEntityReference();
                            _service.Create(InvoiceLine);
                        }
                        else
                        {
                            _retObject.Status = false;
                            _retObject.StatusRemarks = "Tax Group ID is Not Define.";
                            return _retObject;
                        }
                    }
                    else
                    {
                        _retObject.Status = false;
                        _retObject.StatusRemarks = "Charge Type is Not Define.";
                        return _retObject;
                    }
                }

                if (_firstInvoicePaymentMandatory && !_firstInvoiceRaised)
                {
                    Entity _entLine = new Entity(_entColLine.Entities[0].LogicalName, _entColLine.Entities[0].Id);
                    _entLine["ogre_firstinvoiceamount"] = new Money(_billableAmount);
                    _service.Update(_entLine);
                }

                if (_nextInvoiceDate != null && _nextInvoiceDate.Value.Year!=1)
                {
                    Entity _entLine = new Entity(entBooking.LogicalName, entBooking.Id);
                    _entLine["ogre_nextinvoicedate"] = _nextInvoiceDate.Value;
                    _entLine["ogre_lastinvoicedate"] = _InvoiceUntil.Value;
                    _service.Update(_entLine);
                }
                _retObject.Status = true;
                _retObject.StatusRemarks = "Invoice has been created.";
            }
            catch (Exception ex)
            {
                _retObject.Status = false;
                _retObject.StatusRemarks = "ERROR! " + ex.Message;
                return _retObject;
            }
            return _retObject;
        }
    }
    public class ReturnObject
    {
        public bool Status { get; set; }
        public string StatusRemarks { get; set; }
    }
}
