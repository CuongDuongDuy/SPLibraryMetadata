﻿using System;
using System.Collections.Generic;
using System.Security.Policy;

namespace SPLibraryMetadata
{
    public class AdLibraryModel
    {
        public AdLibraryModel()
        {
            Items = new List<AdLibraryItemModel>();
        }

        public Guid Id { get; set; }
        public string Title { get; set; }
        public string Description { get; set; }
        public DateTime Created { get; set; }
        public List<AdLibraryItemModel> Items { get; set; }
    }

    public class AdLibraryItemModel
    {
        public int ID { get; set; }
        public string Title { get; set; }
        public DateTime Modified { get; set; }
        public string DocumentType { get; set; }
        public string _Comments { get; set; }
        public string FileLeafRef { get; set; }
        public string FileRef { get; set; }
        public string File_x0020_Size { get; set; }
    }
}
