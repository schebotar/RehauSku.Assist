using System;
using ExcelDna.Integration;
using System.Net.Http;
using System.Threading.Tasks;

namespace Rehau.Sku.Assist
{
    public class Functions
    {
        private static HttpClient httpClient = new HttpClient();

        [ExcelFunction]
        public static async Task<string> RAUNAME(string request)
        {
            throw new NotImplementedException();
        }
    }
}