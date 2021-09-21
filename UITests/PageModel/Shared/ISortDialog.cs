namespace UITests.PageModel.Shared
{
    public interface ISortDialog
    {
        void OpenSortDialog();
        void CloseSortDialog();
        string[] GetSortOptions();
        void RestoreSortDefaults();
        void Sort(string option, SortOrder sortOrder);
        bool IsSortRestoreDefaultPresent();
    }
}
