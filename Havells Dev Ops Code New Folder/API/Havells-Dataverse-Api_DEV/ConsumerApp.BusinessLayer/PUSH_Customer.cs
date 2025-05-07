using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.ServiceModel.Web;
using System.Text;
using System.Net;
using Microsoft.Xrm.Sdk;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;
using System.ServiceModel.Description;
using System.Configuration;
using System.Threading.Tasks;

namespace ConsumerApp.BusinessLayer
{
    [DataContract]
    public class PUSH_Customer
    {
        [DataMember]
        bool IfExisting { get; set; }
        [DataMember(IsRequired = false)]
        Guid ContactGuId { get; set; }
        [DataMember(IsRequired = false)]
        AVAYA_CreateCustomer Customer = new AVAYA_CreateCustomer();
        [DataMember(IsRequired = false)]
        public List<AVAYA_AddWishList> WishList { get; set; }
        [DataMember(IsRequired = false)]
        public List<AVAYA_CreateProducts> ProductsPurchased { get; set; }
        [DataMember(IsRequired = false)]

        public List<AVAYA_Cart> _tempCart { get; set; }
        public ReturnInfo InitiateOperation(PUSH_Customer _pshCM)
        {
            ReturnInfo Info = new ReturnInfo();
            try
            {
                if(_pshCM.IfExisting == true)
                {
                    Info = Customer.CreateCustomer(Customer);
                    if (Info.CustomerGuid != Guid.Empty)
                    {
                        foreach (AVAYA_AddWishList Wsh in _pshCM.WishList)
                        {
                            Info = Wsh.AddWishList(Wsh, Info.CustomerGuid);
                        }
                        foreach (AVAYA_Cart Cart in _pshCM._tempCart)
                        {
                            Info = Cart.CreateCart(Cart, Info.CustomerGuid);
                        }
                        foreach (AVAYA_CreateProducts Asset in _pshCM.ProductsPurchased)
                        {
                            Info = Asset.CreateTransaction(Asset, Info.CustomerGuid);
                        }
                    }
                }
                else
                {
                    //ReturnInfo _ret = new ReturnInfo();
                    Info = Customer.CreateCustomer(Customer);
                    if(Info.CustomerGuid != Guid.Empty)
                    {
                        foreach (AVAYA_AddWishList Wsh in _pshCM.WishList)
                        {
                            Info = Wsh.AddWishList(Wsh, Info.CustomerGuid);
                        }
                        foreach (AVAYA_Cart Cart in _pshCM._tempCart)
                        {
                            Info = Cart.CreateCart(Cart, Info.CustomerGuid);
                        }
                        foreach (AVAYA_CreateProducts Asset in _pshCM.ProductsPurchased)
                        {
                            Info = Asset.CreateTransaction(Asset, Info.CustomerGuid);
                        }
                    }
                }
                
            }
            catch(Exception ex)
            {
                Info.CustomerGuid = Guid.Empty;
                Info.ErrorCode = "FAILURE";
                Info.ErrorDescription = ex.Message.ToUpper();
            }
            return Info;
        }
    }
}
