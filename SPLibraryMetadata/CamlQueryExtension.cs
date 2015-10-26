using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.SharePoint.Client;

namespace SPLibraryMetadata
{

    #region CamlQuery Extension with field selection, filter and paging

    public class CamlQueryExtension
    {
        private const string CamlQueryTemplate =
            @"<View><ViewFields>{0}</ViewFields><Query>{1}{2}</Query><RowLimit>{3}</RowLimit></View>";

        private const string FieldTemplate = @"<FieldRef Name='{0}'/>";
        private const string WhereTemplate = @"<Where>{0}</Where>";
        private const string FieldCriterionOperatorTemplate = @"<{0}>{1}</{0}>";
        private const string FieldCriterionTemplate = @"<{0}><FieldRef Name='{1}'/><Value Type='{2}'>{3}</Value></{0}>";
        private const string OrderByTemplate = @"<OrderBy>{0}</OrderBy>";
        private const string OrderFieldTemplate = @"<FieldRef Name='{0}' Ascending='{1}'/>";

        private PageInformation CurrentPageInformation { get; set; }
        private PagingIntegrationMetadata CurrentPagingIntegrationMetadata { get; set; }
        public Action<PagingIntegrationMetadata> IntegrateWithPagingAction { get; set; }
        private PagingNavigation PagingNavigationSetting { get; set; }
        private int ItemsPerPage { get; set; }
        private IEnumerable<string> SelectedFields { get; set; }
        private FieldCriteriaOperator CriteriaOperator { get; set; }
        private IEnumerable<FieldCriterionInformation> FieldCriteria { get; set; }
        private IEnumerable<OrderedField> OrderedFields { get; set; }

        private readonly IEnumerable<OrderedField> defaultOrderedFields = new List<OrderedField>
        {
            new OrderedField("almDate", OrderedFieldDirection.Descending)
        };

        #region Integrate with paging

        public void HandleEventFromPaging(PagingIntegrationMetadata metadata)
        {
            if (metadata.Move == PagingNavigationMove.Reset)
            {
                ItemsPerPage = metadata.ItemsPerPage;
            }
            CurrentPagingIntegrationMetadata = metadata;
            CurrentPageInformation = PagingNavigationSetting.GetPageInformation(metadata.Move);
        }

        public CamlQueryIntegrationMetadata GetCamlQueryIntegrateMetadata()
        {
            if (!CheckValidArguments())
            {
                return null;
            }
            var result = new CamlQueryIntegrationMetadata
            {
                Query = new CamlQuery {ViewXml = GetCamlQueryXml()}
            };
            var itemsPosition = new ListItemCollectionPosition
            {
                PagingInfo = CurrentPageInformation.Query
            };
            result.Query.ListItemCollectionPosition = itemsPosition;
            result.PagingIntegrationMetadata = CurrentPagingIntegrationMetadata;
            return result;
        }

        public CamlQueryIntegrationMetadata GetCamlQueryIntegrateMetadata(
            IEnumerable<FieldCriterionInformation> fieldCriteria)
        {
            FieldCriteria = fieldCriteria;
            ResetRowsPerPage();
            return GetCamlQueryIntegrateMetadata();
        }

        public CamlQueryIntegrationMetadata GetCamlQueryIntegrateMetadata(string filterValue)
        {
            if (!CheckValidArguments())
            {
                return null;
            }
            foreach (var fieldCriterionInformation in FieldCriteria)
            {
                fieldCriterionInformation.Value = filterValue;
            }

            ResetRowsPerPage();

            return GetCamlQueryIntegrateMetadata();
        }

        #endregion

        #region Integrate with itself

        public void IntegrateWithCamlQueryExtension(CamlQueryIntegrationMetadata metadata)
        {
            CurrentPageInformation.Query = metadata.PagingInformation;
            PagingNavigationSetting.UpdatePageInformation(CurrentPageInformation);
            metadata.PagingIntegrationMetadata.Status = PagingStatus.Idle;
            CurrentPagingIntegrationMetadata = metadata.PagingIntegrationMetadata;
            if (IntegrateWithPagingAction != null)
            {
                IntegrateWithPagingAction.Invoke(metadata.PagingIntegrationMetadata);
            }
        }

        #endregion

        #region Instructors

        private CamlQueryExtension(IEnumerable<string> selectedFields,
            IEnumerable<FieldCriterionInformation> fieldCriteria, IEnumerable<OrderedField> orderedFields,
            IEnumerable<OrderedField> defaultOrderedFields, int itemsPerPage = 30,
            FieldCriteriaOperator criteriaOperator = FieldCriteriaOperator.And)
        {
            SelectedFields = selectedFields;
            CriteriaOperator = criteriaOperator;
            FieldCriteria = fieldCriteria;
            OrderedFields = orderedFields;
            if (defaultOrderedFields != null)
            {
                this.defaultOrderedFields = defaultOrderedFields;
            }
            ItemsPerPage = itemsPerPage;
            PagingNavigationSetting = new PagingNavigation();
        }

        public CamlQueryExtension(Type itemsType, IEnumerable<FieldCriterionInformation> fieldCriteria,
            IEnumerable<OrderedField> orderedFields, IEnumerable<OrderedField> defaultOrderedFields,
            int itemsPerPage = 50, FieldCriteriaOperator criteriaOperator = FieldCriteriaOperator.And)
            : this(
                MapDataTypeToList(itemsType), fieldCriteria, orderedFields, defaultOrderedFields, itemsPerPage,
                criteriaOperator)
        {
        }

        #endregion

        #region Private Utilities

        private void ResetRowsPerPage()
        {
            CurrentPageInformation = PagingNavigationSetting.GetPageInformation(PagingNavigationMove.Reset);
            CurrentPagingIntegrationMetadata.Status = PagingStatus.Initializing;
            CurrentPagingIntegrationMetadata.Move = PagingNavigationMove.Reset;
            CurrentPagingIntegrationMetadata.ItemsPerPage = ItemsPerPage;
            CurrentPagingIntegrationMetadata.CurrentPage = 1;
        }
        private static IEnumerable<string> MapDataTypeToList(Type itemsType)
        {
            var result = new List<string>();
            foreach (var propertyInfo in itemsType.GetProperties())
            {
                result.Add(propertyInfo.Name);
            }
            return result;
        }

        private bool CheckValidArguments()
        {
            return SelectedFields != null && FieldCriteria != null && defaultOrderedFields != null &&
                   SelectedFields.Any() && FieldCriteria.Any() && defaultOrderedFields.Any();
        }

        private string GetViewFieldsString()
        {

            var result = string.Empty;
            foreach (var selectedField in SelectedFields)
            {
                result += string.Format(FieldTemplate, selectedField);
            }
            return result;
        }

        private string GetFieldCriteriaString()
        {
            if (!FieldCriteria.Any())
            {
                return string.Empty;
            }
            var element = string.Empty;
            foreach (var fieldCriterion in FieldCriteria)
            {
                if (string.IsNullOrEmpty(fieldCriterion.Value)) continue;
                element += string.Format(FieldCriterionTemplate, fieldCriterion.ComparisonOperatorOperator,
                    fieldCriterion.Name, fieldCriterion.Type, fieldCriterion.Value);
            }
            var result = string.Empty;
            if (!string.IsNullOrEmpty(element))
            {
                result = string.Format(WhereTemplate,
                    FieldCriteria.Count() >= 2
                        ? string.Format(FieldCriterionOperatorTemplate, CriteriaOperator, "{0}")
                        : element);
                result = string.Format(result, element);
            }
            return result;
        }

        private string GetOrderByString()
        {
            var element = string.Empty;
            if (OrderedFields != null)
            {
                foreach (var orderedField in OrderedFields)
                {
                    element += string.Format(OrderFieldTemplate, orderedField.Name,
                        orderedField.Direction == OrderedFieldDirection.Ascending ? "true" : "false");
                }
            }
            foreach (var orderedField in defaultOrderedFields)
            {
                element += string.Format(OrderFieldTemplate, orderedField.Name,
                    orderedField.Direction == OrderedFieldDirection.Ascending ? "true" : "false");
            }
            return string.Format(OrderByTemplate, element);
        }

        private string GetCamlQueryXml()
        {
            var result = string.Format(CamlQueryTemplate, GetViewFieldsString(), GetFieldCriteriaString(),
                GetOrderByString(), ItemsPerPage);
            return result;
        }

        #endregion
        
    }

    #endregion

    #region Field for filtering - Where clause

    public class FieldCriterionInformation
    {
        public string Name { get; set; }
        public FieldCriterionDataType Type { get; set; }
        public FieldCriterionComparisonOperator ComparisonOperatorOperator { get; set; }
        public string Value { get; set; }

        public FieldCriterionInformation(string name, string value)
        {
            Name = name;
            Type = FieldCriterionDataType.Text;
            ComparisonOperatorOperator = FieldCriterionComparisonOperator.Contains;
            Value = value;
        }

        public FieldCriterionInformation()
        {
        }
    }

    public enum FieldCriterionDataType
    {
        Text,
        Choice,
        Note
    }

    public enum FieldCriteriaOperator
    {
        Or,
        And
    }

    public enum FieldCriterionComparisonOperator
    {
        Contains,
        Eq
    }

    #endregion

    #region Field for order - Order clause

    public class OrderedField
    {
        public string Name { get; set; }
        public OrderedFieldDirection Direction { get; set; }

        public OrderedField(string name, OrderedFieldDirection direction = OrderedFieldDirection.Ascending)
        {
            Name = name;
            Direction = direction;
        }
    }

    public enum OrderedFieldDirection
    {
        Ascending,
        Descending
    }

    #endregion

    #region Paging Navigation

    public class PagingNavigation
    {
        private readonly PageInformation pageInformationDefault = new PageInformation(0, "Paged=TRUE&p_ID=0", PagingNavigationMove.Reset);
        private List<string> Queries { get; set; }
        private int CurrentIndex { get; set; }
        private PagingNavigationMove Move { get; set; }

        public PagingNavigation()
        {
            CurrentIndex = 0;
            Move = PagingNavigationMove.Current;
            Queries = new List<string>();
            ResetQueries();
        }

        private void ResetQueries()
        {
            Queries.Clear();
            Queries.Add(pageInformationDefault.Query);
        }

        public PageInformation GetPageInformation(PagingNavigationMove navigationMove)
        {
            PageInformation result = null;
            switch (navigationMove)
            {
                case PagingNavigationMove.Next:
                    result = new PageInformation(CurrentIndex + 1, Queries[CurrentIndex + 1], navigationMove);
                    break;
                case PagingNavigationMove.Previous:
                    result = new PageInformation(CurrentIndex - 1, Queries[CurrentIndex - 1], navigationMove);
                    break;
                case PagingNavigationMove.Current:
                    result = new PageInformation(CurrentIndex, Queries[CurrentIndex], navigationMove);
                    break;
                case PagingNavigationMove.Reset:
                    result = new PageInformation(0, "Paged=TRUE&p_ID=0", PagingNavigationMove.Reset);
                    break;
            }
            return result;
        }

        public void UpdatePageInformation(PageInformation pageInformation)
        {
            if (pageInformation == null) return;
            switch (pageInformation.Move)
            {
                case PagingNavigationMove.Next:
                    CurrentIndex++;
                    if (CurrentIndex == Queries.Count() - 1)
                    {
                        Queries.Add(pageInformation.Query);
                    }
                    break;
                case PagingNavigationMove.Previous:
                    CurrentIndex--;
                    break;
                case PagingNavigationMove.Current:
                    if (CurrentIndex == Queries.Count() - 1)
                    {
                        Queries.Add(pageInformation.Query);
                    }
                    break;
                case PagingNavigationMove.Reset:
                    CurrentIndex = 0;
                    ResetQueries();
                    Queries.Add(pageInformation.Query);
                    break;
                default:
                    break;
            }
        }
    }

    public class PageInformation
    {
        public int Index { get; set; }
        public string Query { get; set; }
        public PagingNavigationMove Move { get; set; }

        public PageInformation(int index, string query, PagingNavigationMove navigationMove)
        {
            Index = index;
            Query = query;
            Move = navigationMove;
        }
    }

    public enum PagingNavigationMove
    {
        Next ,
        Previous,
        Current,
        Reset
    }

    #endregion
}