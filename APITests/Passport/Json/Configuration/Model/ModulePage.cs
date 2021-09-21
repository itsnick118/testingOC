using System.Diagnostics.CodeAnalysis;

namespace APITests.Passport.Json.Configuration.Model
{
    [SuppressMessage("ReSharper", "InconsistentNaming")]
    public enum ModulePage
    {
        documentsAdd,
        documentsAddCmis,
        documentsUpload,
        documentsUploadCmis,
        eventAdd,
        eventEdit,
        eventView,
        history,
        listSavedSearch,
        matterDocuments,
        matterDocumentsCmis,
        matterDocumentsFolder,
        matterDocumentsFolderCmis,
        matterEmails,
        matterEmailsCmis,
        matterEmailsFolder,
        matterEmailsFolderCmis,
        matterEvents,
        matterFavorites,
        matterList,
        matterNarratives,
        matterPeople,
        mattersListSavedSearch,
        narrativesAdd,
        narrativesEdit,
        narrativesView,
        personAdd,
        personEdit,
        personView,
        summary
    }
}
