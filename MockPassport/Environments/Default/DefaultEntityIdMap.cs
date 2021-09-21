using MockPassport.Mappings;

namespace MockPassport.Environments.Default
{
    public class DefaultEntityIdMap: IEntityIdMap
    {
        public int AdjustmentLineItem => 43;
        public int EmailDocument => 153;
    }
}