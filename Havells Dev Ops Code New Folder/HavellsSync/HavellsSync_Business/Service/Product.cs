using HavellsSync_ModelData.Common;
using HavellsSync_ModelData.Product;
using HavellsSync_Data;
using HavellsSync_Busines.IService;
using HavellsSync_Data.IManager;

namespace HavellsSync_Busines.Service
{
    public class Product : IProduct
    {
        private readonly IProductManager _productManager;
        public Product(IProductManager productManager)
        {
            Check.Argument.IsNotNull(nameof(productManager), productManager);

            this._productManager = productManager;
        }
       
        public async Task<ValidateSerialNumResponse> ValidateSerialNumber(string SerialNumber, string SessionId)
        {
            return await _productManager.ValidateSerialNumber(SerialNumber, SessionId);
        }
        public async Task<RegisterProductResponse> RegisterProduct(RegisterProductModel objProductData, string SessionId)
        {
            return await _productManager.RegisterProduct(objProductData, SessionId);
        }

        public async Task<ProductDeatilsList> GetRegisteredProducts(RegisterProductModel registeredProduct,string SessionId)
        {
            return await _productManager.GetRegisteredProducts(registeredProduct,  SessionId);
        }
    }
}
