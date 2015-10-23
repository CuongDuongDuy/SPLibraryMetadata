using System;
using System.Collections.Generic;

namespace SPLibraryMetadata
{
    public class ClaimLibraryModel
    {
        public ClaimLibraryModel()
        {
            Items = new List<ClaimLibraryItemModel>();
        }

        public Guid Id { get; set; }
        public string Title { get; set; }
        public string Description { get; set; }
        public DateTime Created { get; set; }
        public List<ClaimLibraryItemModel> Items { get; set; }

    }

    public class ClaimLibraryItemModel
    {
        public int ID { get; set; }
        public DateTime almDate { get; set; }
        public string almAuthor { get; set; }
        public string almAddressee { get; set; }
        public string almNarrative { get; set; }
        public string almClaims_DocumentType { get; set; }
        public string File_x0020_Size { get; set; }
    }
}
