using AngleSharp.Dom;
using ExcelDna.Integration;
using Newtonsoft.Json;
using System;
using System.Globalization;
using System.Linq;
using System.Runtime.Caching;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RehauSku.Assist
{
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
            Uri uri = request.ConvertToUri();

            Task<string> contentTask = Task.Run(() => HttpClientUtil.GetContentByUriAsync(uri));
            Task<IDocument> documentTask = await contentTask.ContinueWith(content => HttpClientUtil.ContentToDocAsync(content));

            return GetProduct(documentTask.Result);
        }
        public static IProduct GetProduct(IDocument document)
        {
            try
            {
                string script = document
                    .Scripts
                    .Where(s => s.InnerHtml.Contains("dataLayer"))
                    .FirstOrDefault()
                    .InnerHtml;

                string json = script
                    .Substring(script.IndexOf("push(") + 5)
                    .TrimEnd(new[] { ')', ';', '\n', ' ' });

                if (!json.Contains("impressions"))
                    return null;

                StoreResponce storeResponse = JsonConvert.DeserializeObject<StoreResponce>(json);
                IProduct product = storeResponse
                    .Ecommerce
                    .Impressions
                    .Where(p => p.Id.IsRehauSku())
                    .FirstOrDefault();

                return product;
            }

            catch (NullReferenceException e)
            {
                MessageBox.Show(e.Message, "Ошибка получения данных", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
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
                    return "Не найдено :(";

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
                    return double.Parse((string)product.Price, CultureInfo.InvariantCulture);
                default:
                    return ExcelError.ExcelErrorValue;
            }
        }
        public static bool IsRehauSku(this string line)
        {
            return Regex.IsMatch(line, @"\d{11}") &&
                line[0].Equals('1') &&
                line[7].Equals('1');
        }
    }
}