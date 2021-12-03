using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using System;
using System.Linq;
using System.Threading.Tasks;
using System.Text.RegularExpressions;

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

    static class SkuAssist
    {
        public static async Task<IProduct> GetProduct(string request)
        {
            Uri uri = request.ConvertToUri(ResponseOrder.NoSettings);

            Task<string> contentTask = Task.Run(() => HttpClientUtil.GetContentByUriAsync(uri));
            Task<IDocument> documentTask = await contentTask.ContinueWith(content => HttpClientUtil.ContentToDocAsync(content));

            IProduct product = await documentTask.ContinueWith(doc => SkuAssist.GetFirstProduct(doc.Result));
            return product;
        }

        public static IProduct GetFirstProduct(IDocument doc)
        {
            return doc
                .All
                .Where(e => e.ClassName == "product-item__desc-top")
                .Where(e => Regex.IsMatch(e.Children[0].TextContent, @"\d{11}", RegexOptions.None))
                .Select(e => 
                    new Product(e.Children[0].TextContent,
                    e.Children[1].TextContent.Trim(new[] { '\n', ' ' })))
                .FirstOrDefault();
        }

        public static Uri GetFirstResultLink(IDocument doc)
        {
            var link = new Uri(doc
                .Links
                .Where(e => e.ClassName == "product-item__title-link js-name")
                .Select(l => ((IHtmlAnchorElement)l).Href)
                .FirstOrDefault());
            return link;
        }

        public static string GetFistResultImageLink(IDocument doc)
        {
            var imageSource = doc.Images
                .Where(x => x.ClassName == "product-item__image")
                .FirstOrDefault();
            return imageSource != null ? imageSource.Source : "Нет ссылки";
        }
    }
}