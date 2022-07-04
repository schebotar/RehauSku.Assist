using ExcelDna.Integration;

namespace RehauSku
{
    public class Functions
    {
        [ExcelFunction(Description = "Получение корректного артикула из строки")]
        public static object GETRAUSKU([ExcelArgument(Name = "\"Строка\"", Description = "строка, содержащая актикул")] string line)
        {
            if (RauSku.TryParse(line, out RauSku rausku))
            {
                return rausku.ToString();
            }

            else return ExcelError.ExcelErrorNA;
        }
    }
}