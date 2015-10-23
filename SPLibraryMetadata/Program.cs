using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace SPLibraryMetadata
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            const string webFullUrl = "https://adittech.sharepoint.com/sites/appdevelopment/";
            const string libTitle = @"Documents";

            var fieldCriteria = new List<FieldCriterionInformation>
            {
                new FieldCriterionInformation
                {
                    Name = "_Comments",
                    Type = FieldCriterionDataType.Note,
                    ComparisonOperatorOperator = FieldCriterionComparisonOperator.Contains,
                    Value = "Wallpaper"
                },
                new FieldCriterionInformation
                {
                    Name = "DocumentType",
                    Type = FieldCriterionDataType.Choice,
                    ComparisonOperatorOperator = FieldCriterionComparisonOperator.Eq,
                    Value = "Documents"
                }
            };

            var orderedFields = new List<OrderedField>
            {
                new OrderedField("DocumentType", OrderedFieldDirection.Descending)
            };

            var defaultOrderedFields = new List<OrderedField>
            {
                new OrderedField("Title"),
            };


            // If defaultOrderedFields == null, will use default value in CamlQueryExtension: ("almDate", OrderedFieldDirection.Descending)
            var cqExtension = new CamlQueryExtension(typeof(AdLibraryItemModel), fieldCriteria, orderedFields, defaultOrderedFields, 10);
            var camlQuery = cqExtension.GetCamlQuery(PagingNavigationMove.Current);
            var adLibraryModel = SharePointMetadataHelper.GetLibraryMetadata(typeof(AdLibraryModel), webFullUrl, libTitle, camlQuery, cqExtension.UpdateCurrentPageQuery) as AdLibraryModel;

            camlQuery = cqExtension.GetCamlQuery(PagingNavigationMove.Next);
            adLibraryModel = SharePointMetadataHelper.GetLibraryMetadata(typeof(AdLibraryModel), webFullUrl, libTitle, camlQuery, cqExtension.UpdateCurrentPageQuery) as AdLibraryModel;

            camlQuery = cqExtension.GetCamlQuery(PagingNavigationMove.Next);
            adLibraryModel = SharePointMetadataHelper.GetLibraryMetadata(typeof(AdLibraryModel), webFullUrl, libTitle, camlQuery, cqExtension.UpdateCurrentPageQuery) as AdLibraryModel;

            camlQuery = cqExtension.GetCamlQuery(PagingNavigationMove.Next);
            adLibraryModel = SharePointMetadataHelper.GetLibraryMetadata(typeof(AdLibraryModel), webFullUrl, libTitle, camlQuery, cqExtension.UpdateCurrentPageQuery) as AdLibraryModel;

            camlQuery = cqExtension.GetCamlQuery(PagingNavigationMove.Previous);
            adLibraryModel = SharePointMetadataHelper.GetLibraryMetadata(typeof(AdLibraryModel), webFullUrl, libTitle, camlQuery, cqExtension.UpdateCurrentPageQuery) as AdLibraryModel;

            camlQuery = cqExtension.GetCamlQuery(PagingNavigationMove.Previous);
            adLibraryModel = SharePointMetadataHelper.GetLibraryMetadata(typeof(AdLibraryModel), webFullUrl, libTitle, camlQuery, cqExtension.UpdateCurrentPageQuery) as AdLibraryModel;

            camlQuery = cqExtension.GetCamlQuery(PagingNavigationMove.Current);
            adLibraryModel = SharePointMetadataHelper.GetLibraryMetadata(typeof(AdLibraryModel), webFullUrl, libTitle, camlQuery, cqExtension.UpdateCurrentPageQuery) as AdLibraryModel;

            Console.WriteLine("Press any key to continue...");
            Console.ReadKey();
        }
    }
}
