using System.Collections.Generic;

namespace SPLibraryMetadata
{
    public class SharePointFieldMetadata
    {
        public IEnumerable<string> Choices { get; set; }
        public string DefaultValue { get; set; }
    }
}
