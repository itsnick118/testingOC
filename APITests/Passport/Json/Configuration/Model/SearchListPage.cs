using System.Runtime.Serialization;

namespace APITests.Passport.Json.Configuration.Model
{
    public enum SearchListPage
    {
        EmbeddedPersonListForTasksOc,
        EmbeddedPersonListOc,

        [EnumMember(Value="matterPersonRIT-ExternalOC")]
        MatterPersonRitExternalOc,

        [EnumMember(Value="matterPersonRIT-ExternalOtherOC")]
        MatterPersonRitExternalOtherOc,

        [EnumMember(Value="matterPersonRIT-InternalOC")]
        MatterPersonRitInternalOc,

        MatterPersonTypeActiveListScreenOc,
        PersonListForExternalOtherMatterPersonOc,
        PersonListInternalPersonTypeOc,
        PersonOrganizationForMatterExternalPersonOc
    }
}
