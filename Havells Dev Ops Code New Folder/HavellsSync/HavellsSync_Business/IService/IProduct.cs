using HavellsSync_ModelData.Product;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HavellsSync_Busines.IService
{
    public interface IProduct
    {
        Task<ValidateSerialNumResponse> ValidateSerialNumber(string serSerialNumber, string SessionId);
        Task<RegisterProductResponse> RegisterProduct(RegisterProductModel objProductData, string SessionId);
        Task<ProductDeatilsList> GetRegisteredProducts(RegisterProductModel registeredProduct, string SessionId);
    }
}
