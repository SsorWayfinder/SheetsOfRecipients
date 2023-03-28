using Microsoft.VisualStudio.TestTools.UnitTesting;
using SheetReader;

namespace SheetReaderTest;

[TestClass]
public class SheetTest
{
    [TestMethod]
    public void TestSheet()
    {
        // arrange
        Sheet testSheet;
        // act
        testSheet = new Sheet("../../../../sheets/SubawardBudgetExample1.xlsx");
        // assert
        Assert.IsNotNull(testSheet);
        
        Assert.IsNotNull(testSheet.subawards["Indiana"]);
        Assert.AreEqual(25000, testSheet.subawards["Indiana"]);

        Assert.IsNotNull(testSheet.subawards["Mayo"]);
        Assert.AreEqual(20637, testSheet.subawards["Mayo"]);

        Assert.IsNotNull(testSheet.subawards["Purdue"]);
        Assert.AreEqual(25000, testSheet.subawards["Purdue"]);

        Assert.IsNotNull(testSheet.subawards["Florida"]);
        Assert.AreEqual(25000, testSheet.subawards["Florida"]);

        Assert.AreEqual(4, testSheet.subawards.Count);

    }


}