namespace UITests.PageModel.Shared.InputControls
{
    public interface IIndexedInputControl
    {
        string SetByIndex(int index, bool clearInput = false);
    }
}