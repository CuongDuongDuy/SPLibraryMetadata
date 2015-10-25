using System;
using System.IO;
using System.Security;
using Microsoft.SharePoint.Client;
using File = Microsoft.SharePoint.Client.File;

namespace SPLibraryMetadata
{
    public class SharePointMetadataHelper
    {
        public static object GetLibraryMetadata(Type libraryType, string webFullUrl, string libraryTitle,
            CamlQueryIntergratedWithPaging camlQueryIntergratedWithPaging,
            Action<string, PagingIntegrationMetadata> updateCurrentPageQueryActionWithPaging)
        {
            if (string.IsNullOrEmpty(webFullUrl) || string.IsNullOrEmpty(libraryTitle)) return null;

            var context = new ClientContext(webFullUrl);

            var secure = new SecureString();
            foreach (var c in "wildbouy~123")
            {
                secure.AppendChar(c);
            }
            context.Credentials = new SharePointOnlineCredentials("adamduong@adittech.onmicrosoft.com", secure);

            var web = context.Web;
            var list = context.Web.Lists.GetByTitle(libraryTitle);

            // Obtain selection for web metadata
            context.Load(web);
            context.Load(list, includes => includes.ItemCount);

            // Get selected fields of items
            var listItemCollection = list.GetItems(camlQueryIntergratedWithPaging.Query);
            context.Load(listItemCollection);

            context.ExecuteQuery();

            camlQueryIntergratedWithPaging.Metadata.Status = PagingStatus.Idle;

            updateCurrentPageQueryActionWithPaging.Invoke(listItemCollection.ListItemCollectionPosition == null
                ? string.Empty
                : listItemCollection.ListItemCollectionPosition.PagingInfo, camlQueryIntergratedWithPaging.Metadata);

            var result = PropertiesMapper(web, list, listItemCollection, libraryType);
            return result;
        }

        public static object GetLibraryMetadata(Type libraryType, string webFullUrl, string libraryTitle, CamlQuery camlQuery, Action<string> updateCurrentPageQueryAction)
        {
            if (string.IsNullOrEmpty(webFullUrl) || string.IsNullOrEmpty(libraryTitle)) return null;

            var context = new ClientContext(webFullUrl);

            var secure = new SecureString();
                foreach (var c in "wildbouy~123")
                {
                    secure.AppendChar(c);
                }
            context.Credentials = new SharePointOnlineCredentials("adamduong@adittech.onmicrosoft.com", secure);

            var web = context.Web;
            var list = context.Web.Lists.GetByTitle(libraryTitle);

            // Obtain selection for web metadata
            context.Load(web);
            context.Load(list, includes => includes.ItemCount);

            // Get selected fields of items
            var listItemCollection = list.GetItems(camlQuery);
            context.Load(listItemCollection);

            context.ExecuteQuery();

            updateCurrentPageQueryAction.Invoke(listItemCollection.ListItemCollectionPosition == null
                ? string.Empty
                : listItemCollection.ListItemCollectionPosition.PagingInfo);

            var result = PropertiesMapper(web, list, listItemCollection, libraryType);
            return result;
        }

        public static object GetLibraryMetadata(Type libraryType, string webFullUrl, string libraryTitle)
        {
            if (string.IsNullOrEmpty(webFullUrl) || string.IsNullOrEmpty(libraryTitle)) return null;

            var context = new ClientContext(webFullUrl);

            var secure = new SecureString();
            foreach (var c in "wildbouy@123")
            {
                secure.AppendChar(c);
            }
            context.Credentials = new SharePointOnlineCredentials("adamduong@adittech.onmicrosoft.com", secure);

            var web = context.Web;
            var list = context.Web.Lists.GetByTitle(libraryTitle);

            // Obtain selection for web metadata
            context.Load(web);
            context.Load(list);

            // Get all fields of items
            var listItemCollection = list.GetItems(CamlQuery.CreateAllItemsQuery(20));
            context.Load(listItemCollection);

            context.ExecuteQuery();

            var result = PropertiesMapper(web, list, listItemCollection, libraryType);
            return result;
        }

        private static object PropertiesMapper(Web webMetadata, List libMetadata, ListItemCollection itemsMetadata, Type dataType)
        {
            if (webMetadata == null || libMetadata == null)
                return null;

            // Get type of item property
            var itemsProp = dataType.GetProperty("Items");
            var itemType = itemsProp.PropertyType.GetGenericArguments()[0];

            // Map data for library metadata
            var result = Activator.CreateInstance(dataType);
            foreach (var propertyInfo in dataType.GetProperties())
            {
                if (propertyInfo.Name != "Items")
                {
                    switch (propertyInfo.Name)
                    {
                        case "TotalItemCount":
                            propertyInfo.SetValue(result, libMetadata.ItemCount, null);
                            break;
                        default:
                            propertyInfo.SetValue(result, webMetadata.GetType().GetProperty(propertyInfo.Name).GetValue(webMetadata, null), null);
                            break;
                    }
                }
            }

            // Map data for library items metadata
            foreach (var itemMetadata in itemsMetadata)
            {
                var item = Activator.CreateInstance(itemType);
                foreach (var propertyInfo in itemType.GetProperties())
                {
                    if (itemMetadata.FieldValues.ContainsKey(propertyInfo.Name))
                    {
                        try
                        {
                            propertyInfo.SetValue(item, itemMetadata.FieldValues[propertyInfo.Name], null);
                        }
                        catch (ArgumentException)
                        {
                            var safeType = Nullable.GetUnderlyingType(propertyInfo.PropertyType) ?? propertyInfo.PropertyType;

                            var safeValue = (itemMetadata.FieldValues[propertyInfo.Name] == null) ? null : Convert.ChangeType(itemMetadata.FieldValues[propertyInfo.Name], safeType);
                            propertyInfo.SetValue(item, safeValue, null);
                        }

                    }
                }
                itemsProp.GetValue(result, null)
                    .GetType()
                    .GetMethod("Add")
                    .Invoke(itemsProp.GetValue(result, null), new[] { item });
            }
            return result;
        }

        public static Stream GetFile(string webFullUrl, string fileRelativeUrl)
        {
            if (string.IsNullOrEmpty(webFullUrl) || string.IsNullOrEmpty(fileRelativeUrl)) return null;

            var context = new ClientContext(webFullUrl);
            var web = context.Web;
            var file = web.GetFileByServerRelativeUrl(fileRelativeUrl);
            context.Load(file);
            context.ExecuteQuery();
            var fileInfo = File.OpenBinaryDirect(context, fileRelativeUrl);
            return fileInfo.Stream;
        }

        public static SharePointFieldMetadata GetFieldValues(string webFullUrl, string fieldTitle, string libraryTitle = "Documents")
        {
            SharePointFieldMetadata result;
            try
            {
                if (string.IsNullOrEmpty(webFullUrl) || string.IsNullOrEmpty(fieldTitle)) return null;

                var context = new ClientContext(webFullUrl);
                var web = context.Web;
                var list = web.Lists.GetByTitle(libraryTitle);
                var field = list.Fields.GetByInternalNameOrTitle(fieldTitle);
                var choiceField = context.CastTo<FieldChoice>(field);
                context.Load(field);
                context.ExecuteQuery();

                result = new SharePointFieldMetadata
                {
                    Choices = choiceField.Choices,
                    DefaultValue = choiceField.DefaultValue
                };
            }
            catch
            {
                result = new SharePointFieldMetadata();
            }
            return result;
        }
    }
}
