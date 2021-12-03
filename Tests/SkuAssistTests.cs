using NUnit.Framework;

namespace Rehau.Sku.Assist.Tests
{
    [TestFixture]
    public class SkuAssistTests
    {
        [Test]
        public static void BaseTest()
        {
            var result = Functions.RAUNAME("160001");
            Assert.AreEqual("Надвижная гильза REHAU RAUTITAN РХ (11600011001)", result);
        }
    }
}
