using System;
using Microsoft.SharePoint.Client;

namespace SPLibraryMetadata
{
    public enum PagingStatus
    {
        Initializing,
        Loading,
        Idle
    }

    public class PagingIntegrationMetadata
    {
        public PagingStatus Status { get; set; }
        public PagingNavigationMove Move { get; set; }
        public int CurrentPage { get; set; }
        public int ItemsPerPage { get; set; }
        public int TotalRows { get; set; }

    }

    public class CamlQueryIntegrationMetadata
    {
        public CamlQuery Query { get; set; }
        public string PagingInformation { get; set; }
        public PagingIntegrationMetadata PagingIntegrationMetadata { get; set; }

    }
}
