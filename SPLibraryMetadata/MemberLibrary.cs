using System;
using System.Collections.Generic;

namespace SPLibraryMetadata
{
    public class MemberLibraryModel
    {
        public MemberLibraryModel()
        {
            Items = new List<MemberLibraryItemModel>();
        }

        public Guid Id { get; set; }
        public string Title { get; set; }
        public string Description { get; set; }
        public DateTime Created { get; set; }
        public List<MemberLibraryItemModel> Items { get; set; }
    }

    public class MemberLibraryItemModel
    {
        public int ID { get; set; }
        public DateTime almDate { get; set; }
        public string almAuthor { get; set; }
        public string almAddressee { get; set; }
        public string almNarrative { get; set; }
        public string almMembers_DocumentType { get; set; }
        public string FileLeafRef { get; set; }
        public string FileRef { get; set; }
        public string File_x0020_Size { get; set; }
    }
}
