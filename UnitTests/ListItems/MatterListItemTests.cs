using Moq;
using NUnit.Framework;
using OpenQA.Selenium;
using System;
using UITests.PageModel;
using UITests.PageModel.Shared;

namespace UnitTests.ListItems
{
    public class MatterListItemTests
    {
        private Mock<IAppInstance> _appInstanceMock;
        private Mock<IWebElement> _webElementMock;

        [SetUp]
        public void SetUp()
        {
            _appInstanceMock = new Mock<IAppInstance>();
            _webElementMock = new Mock<IWebElement>();
        }

        public class TertiaryTextTests : MatterListItemTests
        {
            [SetUp]
            public void TertiaryTextTestsSetUp()
            {
                _webElementMock.Setup(x => x.FindElement(It.IsAny<By>())).Returns(_webElementMock.Object);
            }

            [Test]
            public void MatterListItem_WhenStatusDateAvailable_TertiaryTextIsParsedCorrectly()
            {
                const string tertiaryText = "9394 ● Eric Stone ● Open - 12/23/2009";
                _webElementMock.Setup(x => x.Text).Returns(tertiaryText);

                var matter = new MatterListItem(_appInstanceMock.Object, _webElementMock.Object);

                Assert.That(matter.Number, Is.EqualTo("9394"));
                Assert.That(matter.PrimaryInternalContact, Is.EqualTo("Eric Stone"));
                Assert.That(matter.Status, Is.EqualTo("Open"));
                Assert.That(matter.StatusDate, Is.EqualTo(Convert.ToDateTime("12/23/2009")));
            }

            [Test]
            public void MatterListItem_WhenStatusDateUnavailable_TertiaryTextIsParsedCorrectly()
            {
                const string tertiaryText = "MAT-22 ● Alice Lee ● Pending Assignment";
                _webElementMock.Setup(x => x.Text).Returns(tertiaryText);

                var matter = new MatterListItem(_appInstanceMock.Object, _webElementMock.Object);

                Assert.That(matter.Number, Is.EqualTo("MAT-22"));
                Assert.That(matter.PrimaryInternalContact, Is.EqualTo("Alice Lee"));
                Assert.That(matter.Status, Is.EqualTo("Pending Assignment"));
                Assert.That(matter.StatusDate, Is.Null);
            }

            [Test]
            public void MatterListItem_WhenSpendToDateAvailable_ReturnsSpendToDate()
            {
                const string tertiaryText = "2912 ● Eric Stone ● Open - 04/27/2011 ● 20.00 USD";
                _webElementMock.Setup(x => x.Text).Returns(tertiaryText);

                var matter = new MatterListItem(_appInstanceMock.Object, _webElementMock.Object);

                Assert.That(matter.SpendToDate, Is.EqualTo("20.00 USD"));
            }

            [Test]
            public void MatterListItem_WhenSpendToDateUnavailable_ReturnsNull()
            {
                const string tertiaryText = "2912 ● Eric Stone ● Open - 04/27/2011";
                _webElementMock.Setup(x => x.Text).Returns(tertiaryText);

                var matter = new MatterListItem(_appInstanceMock.Object, _webElementMock.Object);

                Assert.That(matter.SpendToDate, Is.Null);
            }
        }
    }
}
