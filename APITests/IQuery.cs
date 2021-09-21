using System.Collections.Generic;

namespace APITests
{ 
    internal interface IQuery
    {
        IDictionary<string, string> AsDictionary();
    }
}
