using System;
using HavellsSync_ModelData.Product;

namespace HavellsSync_Data.IManager
{
    public interface IProductManager
    {
        Task<ValidateSerialNumResponse> ValidateSerialNumber(string serialNumber, string SessionId);
        Task<RegisterProductResponse> RegisterProduct(RegisterProductModel objProductData, string SessionId);
        Task<ProductDeatilsList> GetRegisteredProducts(RegisterProductModel registeredProduct, string SessionId);
    }
}
