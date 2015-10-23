using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.SharePoint.Client;

namespace SPLibraryMetadata
{
    public class CamlQueryExtension
    {
        private const string CamlQueryTemplate = @"<View><ViewFields>{0}</ViewFields><Query>{1}{2}</Query><RowLimit>{3}</RowLimit></View>";
        private const string FieldTemplate = @"<FieldRef Name='{0}'/>";
        private const string WhereTemplate = @"<Where>{0}</Where>";
        private const string FieldCriterionOperatorTemplate = @"<{0}>{1}</{0}>";
        private const string FieldCriterionTemplate = @"<{0}><FieldRef Name='{1}'/><Value Type='{2}'>{3}</Value></{0}>";
        private const string OrderByTemplate = @"<OrderBy>{0}</OrderBy>";
        private const string OrderFieldTemplate = @"<FieldRef Name='{0}' Ascending='{1}'/>";

        private PageInformation currentPageInformation;

        public PageInformation CurrentPageInformation
        {
            get { return currentPageInformation; }
        }

        private PagingNavigation PagingNavigationSetting { get; set; }
        public int NumberPerPage { get; set; }
        public IEnumerable<string> SelectedFields { get; set; }
        public FieldCriteriaOperator CriteriaOperator { get; set; }
        public IEnumerable<FieldCriterionInformation> FieldCriteria { get; set; }
        public IEnumerable<OrderedField> OrderedFields { get; set; }

        private readonly IEnumerable<OrderedField> defaultOrderedFields = new List<OrderedField>
        {
            new OrderedField("almDate", OrderedFieldDirection.Descending)
        };

        private bool CheckValidArguments()
        {
            return SelectedFields != null && FieldCriteria != null && defaultOrderedFields != null && SelectedFields.Any() && FieldCriteria.Any() && defaultOrderedFields.Any();
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
                    element += string.Format(OrderFieldTemplate, orderedField.Name, orderedField.Direction == OrderedFieldDirection.Ascending ? "true" : "false");
                }
            }
            foreach (var orderedField in defaultOrderedFields)
            {
                element += string.Format(OrderFieldTemplate, orderedField.Name, orderedField.Direction == OrderedFieldDirection.Ascending ? "true" : "false");
            }
            return string.Format(OrderByTemplate, element);
        }

        private string GetCamlQueryXml()
        {
            var result = string.Format(CamlQueryTemplate, GetViewFieldsString(), GetFieldCriteriaString(), GetOrderByString(), NumberPerPage);
            return result;
        }

        public CamlQuery GetCamlQuery(PagingNavigationMove navigationMove)
        {
            if (!CheckValidArguments())
            {
                return null;
            }
            var result = new CamlQuery
            {
                ViewXml = GetCamlQueryXml()
            };
            currentPageInformation = PagingNavigationSetting.GetPageInformation(navigationMove);
            var itemsPosition = new ListItemCollectionPosition
            {
                PagingInfo = CurrentPageInformation.Query
            };
            result.ListItemCollectionPosition = itemsPosition;
            return result;
        }

        public CamlQuery GetCamlQuery(string filterValue, PagingNavigationMove navigationMove)
        {
            if (!CheckValidArguments())
            {
                return null;
            }
            foreach (var fieldCriterionInformation in FieldCriteria)
            {
                fieldCriterionInformation.Value = filterValue;
            }
            return GetCamlQuery(navigationMove);
        }

        public CamlQuery GetCamlQuery(string filterValue, PagingNavigationMove navigationMove, FieldCriterionComparisonOperator comparisonOperator)
        {
            if (!CheckValidArguments())
            {
                return null;
            }
            foreach (var fieldCriterionInformation in FieldCriteria)
            {
                fieldCriterionInformation.ComparisonOperatorOperator = comparisonOperator;
                fieldCriterionInformation.Value = filterValue;
            }
            return GetCamlQuery(navigationMove);
        }

        public CamlQueryExtension(IEnumerable<string> selectedFields, IEnumerable<FieldCriterionInformation> fieldCriteria, IEnumerable<OrderedField> orderedFields, IEnumerable<OrderedField> defaultOrderedFields, int numberPerPage = 50, FieldCriteriaOperator criteriaOperator = FieldCriteriaOperator.And)
        {
            SelectedFields = selectedFields;
            CriteriaOperator = criteriaOperator;
            FieldCriteria = fieldCriteria;
            OrderedFields = orderedFields;
            if (defaultOrderedFields != null)
            {
                this.defaultOrderedFields = defaultOrderedFields;
            }
            NumberPerPage = numberPerPage;
            PagingNavigationSetting = new PagingNavigation();
        }

        public void UpdateCurrentPageQuery(string pagingInformation)
        {
            CurrentPageInformation.Query = pagingInformation;
            PagingNavigationSetting.UpdatePageInformation(CurrentPageInformation);
        }

        public CamlQueryExtension(Type itemsType, IEnumerable<FieldCriterionInformation> fieldCriteria, IEnumerable<OrderedField> orderedFields, IEnumerable<OrderedField> defaultOrderedFields, int numberPerPage = 50, FieldCriteriaOperator criteriaOperator = FieldCriteriaOperator.And)
            : this(MapDataTypeToList(itemsType), fieldCriteria, orderedFields, defaultOrderedFields, numberPerPage, criteriaOperator)
        {

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
    }

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

    public class PagingNavigation
    {
        private readonly PageInformation pageInformationDefault = new PageInformation(0, "Paged=TRUE&p_ID=0", PagingNavigationMove.RowsPerPage);
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

        public PageInformation GetPageInformation(PagingNavigationMove navigationMove = PagingNavigationMove.Next)
        {
            PageInformation result;
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
                case PagingNavigationMove.RowsPerPage:
                    result = pageInformationDefault;
                    break;
                default:
                    result = pageInformationDefault;
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
                case PagingNavigationMove.RowsPerPage:
                    CurrentIndex = 0;
                    ResetQueries();
                    Queries.Add(pageInformation.Query);
                    break;
                default:
                    break;
            }
        }
    }

    public enum PagingNavigationMove
    {
        Next,
        Previous,
        RowsPerPage,
        Current
    }

    #endregion

}