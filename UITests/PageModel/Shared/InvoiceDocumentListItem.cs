using System;
using OpenQA.Selenium;

namespace UITests.PageModel.Shared
{
    public class InvoiceDocumentListItem : MatterDocumentListItem
    {
        public InvoiceDocumentListItem(IAppInstance app, IWebElement element) : base(app, element)
        {
        }

        public new void QuickFile()
        {
            if (IsFolder)
            {
                base.QuickFile();
            }
            else
            {
                throw new NotSupportedException();
            }
        }
    }
}
