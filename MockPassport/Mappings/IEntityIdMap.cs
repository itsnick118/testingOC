using System.Collections.Generic;

namespace MockPassport.Mappings
{
    public interface IEntityIdMap
    {
        int AdjustmentLineItem { get; }
        int EmailDocument { get; }
    }

    public interface IEntityNameMap
    {
        IDictionary<string, string> MetaDataFiles { get; }
    }
}