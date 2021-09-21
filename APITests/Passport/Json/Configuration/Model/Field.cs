namespace APITests.Passport.Json.Configuration.Model
{
    public class Field
    {
        public string Attribute { get; set; }
        public int[] ColumnId { get; set; }
        public bool Conditional { get; set; }
        public string ConditionalDisableFieldName { get; set; }
        public string ConditionalDisableFieldValue { get; set; }
        public dynamic ConditionalFilter { get; set; }
        public string ControlType { get; set; }
        public string DisplayName { get; set; }
        public string FilterDescription { get; set; }
        public int Id { get; set; }
        public string Label { get; set; }
        public bool LiveUpdate { get; set; }
        public string Name { get; set; }
        public bool ReadOnly { get; set; }
        public string RelatedEntity { get; set; }
        public bool Required { get; set; }
        public string SectionGroup { get; set; }
        public bool SelectedByDefault { get; set; }
        public int Size { get; set; }
        public string SourceListScreen { get; set; }
        public string Type { get; set; }
        public string Value { get; set; }
    }
}
