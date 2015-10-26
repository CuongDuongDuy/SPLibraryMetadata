using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Security.Cryptography.X509Certificates;

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
            var pgingExtension = new PagingViewModel(10);
            var cqExtension = new CamlQueryExtension(typeof (AdLibraryItemModel), fieldCriteria, orderedFields,
                defaultOrderedFields, FieldCriteriaOperator.And, pgingExtension);

            var camlQueryIntegrationMetadata = cqExtension.GetCamlQueryIntegratedMetadata();
            var adLibraryModel = SharePointMetadataHelper.GetLibraryMetadata(typeof(AdLibraryModel), webFullUrl, libTitle, camlQueryIntegrationMetadata) as AdLibraryModel;
            adLibraryModel.Show();

            pgingExtension.RowsPerPage = 30;
            camlQueryIntegrationMetadata = cqExtension.GetCamlQueryIntegratedMetadata();
            adLibraryModel = SharePointMetadataHelper.GetLibraryMetadata(typeof(AdLibraryModel), webFullUrl, libTitle, camlQueryIntegrationMetadata) as AdLibraryModel;
            adLibraryModel.Show();

            fieldCriteria = new List<FieldCriterionInformation>
            {
                new FieldCriterionInformation
                {
                    Name = "DocumentType",
                    Type = FieldCriterionDataType.Choice,
                    ComparisonOperatorOperator = FieldCriterionComparisonOperator.Eq,
                    Value = "Data"
                }
            };

            camlQueryIntegrationMetadata = cqExtension.GetCamlQueryIntegratedMetadata(fieldCriteria,
                FieldCriteriaOperator.And);
            adLibraryModel = SharePointMetadataHelper.GetLibraryMetadata(typeof(AdLibraryModel), webFullUrl, libTitle, camlQueryIntegrationMetadata) as AdLibraryModel;
            adLibraryModel.Show();

            pgingExtension.MoveNextCommand(PagingNavigationMove.Next);
            camlQueryIntegrationMetadata = cqExtension.GetCamlQueryIntegratedMetadata();
            adLibraryModel = SharePointMetadataHelper.GetLibraryMetadata(typeof(AdLibraryModel), webFullUrl, libTitle, camlQueryIntegrationMetadata) as AdLibraryModel;
            adLibraryModel.Show();

            Console.WriteLine("Press any key to continue...");
            Console.ReadKey();
        }
    }
}
