using System.Runtime.Serialization;

namespace APITests.Passport
{
    internal enum OcModule
    {
        [EnumMember(Value="matter")]
        Matter,
        [EnumMember(Value="spend")]
        Spend,
        [EnumMember(Value="globalDocuments")]
        GlobalDocuments
    }
}
