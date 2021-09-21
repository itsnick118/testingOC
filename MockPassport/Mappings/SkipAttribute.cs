using System;

namespace MockPassport.Mappings
{
    [AttributeUsage(AttributeTargets.Method)]
    public class SkipAttribute : Attribute
    {
        // Will prevent running the Update method for a given IUpdatable mapping definition.
    }
}