using NUnit.Framework;

namespace Rehau.Sku.Assist.Tests
{
    [TestFixture]
    public class SkuAssistTests
    {
        [Test]
        public async void BaseTest()
        {
            var result = await Functions.RAUNAME("160001");
            Assert.AreEqual("Надвижная гильза REHAU RAUTITAN РХ (11600011001)", result);
        }
    }
}
