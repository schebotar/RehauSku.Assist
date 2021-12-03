using AngleSharp.Dom;
using ExcelDna.Integration;
using Newtonsoft.Json;
using System;
using System.Linq;
using System.Runtime.Caching;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Rehau.Sku.Assist
{
    public enum ResponseOrder
    {
        NoSettings,
        Relevance,
        Name,
        Price,
        Series
    }

    public enum ProductField
    {
        Name,
        Id,
        Price
    }

    static class SkuAssist
    {
        public static async Task<IProduct> GetProduct(string request)
        {
            Uri uri = request.ConvertToUri(ResponseOrder.NoSettings);

            Task<string> contentTask = Task.Run(() => HttpClientUtil.GetContentByUriAsync(uri));
            Task<IDocument> documentTask = await contentTask.ContinueWith(content => HttpClientUtil.ContentToDocAsync(content));

            return GetProduct(documentTask.Result);
        }

        public static IProduct GetProduct(IDocument d)
        {
            string script = d.Scripts
                   .Where(s => s.InnerHtml.Contains("dataLayer"))
                   .First()
                   .InnerHtml;

            string json = script
                .Substring(script.IndexOf("push(") + 5)
                .TrimEnd(new[] { ')', ';', '\n', ' ' });

            StoreResponce storeResponse = JsonConvert.DeserializeObject<StoreResponce>(json);
            IProduct product = storeResponse
                .Ecommerce
                .Impressions
                .Where(p => Regex.IsMatch(p.Id, @"\d{11}", RegexOptions.None))
                .FirstOrDefault();

            return product;
        }

        public static object GetProduct(string request, ProductField field)
        {
            IProduct product;

            if (MemoryCache.Default.Contains(request))
            {
                product = MemoryCache.Default[request] as IProduct;
            }

            else
            {
                object result = ExcelAsyncUtil.Run("RauName", new[] { request },
                    delegate
                    {
                        Task<IProduct> p = Task.Run(() => GetProduct(request));
                        return p.Result;
                    });

                if (result == null)
                    return "Не найдено";

                if (result.Equals(ExcelError.ExcelErrorNA))
                    return "Загрузка...";

                product = result as IProduct;
                MemoryCache.Default.Add(request, product, DateTime.Now.AddMinutes(10));
            }

            switch (field)
            {
                case ProductField.Name:
                    return product.Name;
                case ProductField.Id:
                    return product.Id;
                case ProductField.Price:
                    return product.Price;
                default:
                    return ExcelError.ExcelErrorValue;
            }
        }
    }
}