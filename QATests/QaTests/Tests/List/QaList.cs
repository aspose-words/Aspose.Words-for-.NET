using Aspose.Words;
using Aspose.Words.Lists;
using NUnit.Framework;

namespace QaTests.Tests
{
    using List = Aspose.Words.Lists.List;

    [TestFixture]
    internal class QaList : QaTestsBase
    {
        private readonly string _image = MyDir + @"Images\Test_636_852.gif";

        [Test]
        public void ListLevel_PictureBullet()
        {
            Document doc = new Document();

            // Create a list with template
            List list = doc.Lists.Add(ListTemplate.BulletCircle);

            // Create picture bullet for the current list level
            list.ListLevels[0].CreatePictureBullet();

            // Set your own picture bullet image through the ImageData
            list.ListLevels[0].ImageData.SetImage(this._image);

            Assert.IsTrue(list.ListLevels[0].ImageData.HasImage);
            
            // Delete picture bullet
            list.ListLevels[0].DeletePictureBullet();

            Assert.IsNull(list.ListLevels[0].ImageData);
        }
    }
}
