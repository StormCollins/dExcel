namespace dExcelTests.OmicronUtilsTests;

using NUnit.Framework;
using dExcel.OmicronUtils;

[TestFixture]
public class OmicronUtilsTests
{
    [Test]
    public void GetDataTest()
    {
        // OmicronCurveUtils.PullData();
        // OmicronCurveUtils.SerializeOmicronObject();

        string interestRateSwapJson =
            "{" +
                "\"type\":" +
                "{" +
                    "\"$type\":\"InterestRateSwap\"," +
                    "\"ReferenceIndex\": " +
                    "{" +
                        "\"$type\":\"RateIndex\", " +
                        "\"Name\":\"JIBAR\", " +
                        "\"Tenor\":{\"amount\":3,\"unit\":\"Month\"}" +
                    "}," +
                    "\"PaymentFrequency\":{\"amount\":3,\"unit\":\"Month\"}," +
                    "\"Tenor\":{\"amount\":10,\"unit\":\"Year\"}" +
                "}," +
                "\"date\":\"2023-02-14T00:00:00\"," +
                "\"value\":0.08779999999999999" +
            "}";
    }
}
