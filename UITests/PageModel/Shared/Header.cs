namespace UITests.PageModel.Shared
{
    public class Header
    {
        private readonly IAppInstance _app;

        public Header(IAppInstance appInstance)
        {
            _app = appInstance;
        }

        public void NavigateBack() => _app.WaitAndClickThenWait(Selectors.Oc.BackButton);

        public void OpenHelp() => _app.WaitAndClickThenWait(Selectors.Oc.HelpIcon);

        public void OpenMattersAppTab() => _app.WaitAndClickThenWait(Selectors.Oc.MattersTab);

        public void OpenSpendAppTab() => _app.WaitAndClickThenWait(Selectors.Oc.SpendTab);

        public void OpenUploadHistory() => _app.WaitAndClickThenWait(Selectors.Oc.OpenUploadHistory);

        public void OpenUploadQueue() => _app.WaitAndClickThenWait(Selectors.Oc.UploadQueueIcon);

        public void CancelAllQueued() => _app.WaitAndClickThenWait(Selectors.Oc.CancelAll);

        public bool IsUploadEmailWaitingQueueDisplayed() => _app.IsElementDisplayed(Selectors.Oc.UploadIndicatorStopWatch);
    }
}
