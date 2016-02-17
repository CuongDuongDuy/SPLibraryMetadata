using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Security;
using Microsoft.SharePoint.Client;
using File = Microsoft.SharePoint.Client.File;

namespace SPLibraryMetadata
{
    public class SharePointMetadataHelper
    {
        public static object GetLibraryMetadata(Type libraryType, string webFullUrl, string libraryTitle,
            CamlQryIntegrationMetadata camlQryIntegrationMetadata)
        {
            if (string.IsNullOrEmpty(webFullUrl) || string.IsNullOrEmpty(libraryTitle)) return null;

            var context = new ClientContext(webFullUrl);

            context.Credentials = CredentialCache.DefaultNetworkCredentials;

            var web = context.Web;
            var list = context.Web.Lists.GetByTitle(libraryTitle);

            // Obtain selection for web metadata
            context.Load(web);
            context.Load(list, includes => includes.ItemCount);

            // Get selected fields of items
            var listItemCollection = list.GetItems(camlQryIntegrationMetadata.Query);
            context.Load(listItemCollection);

            context.ExecuteQuery();

            camlQryIntegrationMetadata.PagingInformation = listItemCollection.ListItemCollectionPosition == null
                ? string.Empty
                : listItemCollection.ListItemCollectionPosition.PagingInfo;

            camlQryIntegrationMetadata.IntegrateToCamlQrExAction.Invoke(camlQryIntegrationMetadata);

            var result = PropertiesMapper(web, list, listItemCollection, libraryType);
            return result;
        }

        public static object GetLibraryMetadata(Type libraryType, string webFullUrl, string libraryTitle)
        {
            if (string.IsNullOrEmpty(webFullUrl) || string.IsNullOrEmpty(libraryTitle)) return null;

            var context = new ClientContext(webFullUrl);

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

        private static object PropertiesMapper(Web webMetadata, List libMetadata, ListItemCollection itemsMetadata,
            Type dataType)
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
                            propertyInfo.SetValue(result,
                                webMetadata.GetType().GetProperty(propertyInfo.Name).GetValue(webMetadata, null), null);
                            break;
                    }
                }
            }

            // Map data for library items metadata
            foreach (var itemMetadata in itemsMetadata)
            {
                var test = itemMetadata.Id;
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
                            var safeType = Nullable.GetUnderlyingType(propertyInfo.PropertyType) ??
                                           propertyInfo.PropertyType;

                            var safeValue = (itemMetadata.FieldValues[propertyInfo.Name] == null)
                                ? null
                                : Convert.ChangeType(itemMetadata.FieldValues[propertyInfo.Name], safeType);
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

        public static SharePointFieldMetadata GetFieldValues(string webFullUrl, string fieldTitle,
            string libraryTitle = "Documents")
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

        public static void CreateFolder(string webFullUrl, string libraryName, string folderName)
        {
            using (var clientContext = new ClientContext(webFullUrl))
            {
                var web = clientContext.Web;
                var list = web.Lists.GetByTitle(libraryName);
                list.RootFolder.Folders.Add(folderName);
                clientContext.ExecuteQuery();
            }
        }

        public static void ChangePermissionForLibrary(string webFullUrl, string libraryName,
            IEnumerable<SpPermission> permissions = null)
        {
            using (var ctx = new ClientContext(webFullUrl))
            {
                var list = ctx.Web.Lists.GetByTitle(libraryName);
                if (permissions == null)
                {
                    list.ResetRoleInheritance();
                }
                else
                {
                    list.BreakRoleInheritance(false, false);
                    var spPermissions = permissions as SpPermission[] ?? permissions.ToArray();
                    foreach (var permission in spPermissions)
                    {
                        foreach (var userOrGroup in permission.UsersOrGroups)
                        {
                            Principal user = ctx.Web.EnsureUser(userOrGroup);
                            var roleDefinition = ctx.Site.RootWeb.RoleDefinitions.GetByType(permission.RoleType);
                            var roleBindings = new RoleDefinitionBindingCollection(ctx) { roleDefinition };
                            list.RoleAssignments.Add(user, roleBindings);
                        }

                    }
                }
                ctx.ExecuteQuery();
            }
        }

        // Can not change permission for Folder on CSOM 2010 v14
        // Must apply to each item
        public static void ChangePermissionForFolder(string webFullUrl, string libraryName, string folderName,
            IEnumerable<SpPermission> permissions = null)
        {
            using (var ctx = new ClientContext(webFullUrl))
            {
                var relativeUrl = string.Format("{0}/{1}", libraryName, folderName);
                var folder = ctx.Web.GetFolderByServerRelativeUrl(relativeUrl);
                ctx.Load(folder.Files);
                ctx.ExecuteQuery();
                if (permissions == null)
                {
                    foreach (var file in folder.Files)
                    {
                        file.ListItemAllFields.ResetRoleInheritance();
                    }
                }
                else
                {
                    var spPermissions = permissions as SpPermission[] ?? permissions.ToArray();
                    foreach (var file in folder.Files)
                    {
                        file.ListItemAllFields.BreakRoleInheritance(false, false);
                        foreach (var permission in spPermissions)
                        {
                            foreach (var userOrGroup in permission.UsersOrGroups)
                            {
                                Principal user = ctx.Web.EnsureUser(userOrGroup);
                                var roleDefinition = ctx.Site.RootWeb.RoleDefinitions.GetByType(permission.RoleType);
                                var roleBindings = new RoleDefinitionBindingCollection(ctx) { roleDefinition };
                                file.ListItemAllFields.RoleAssignments.Add(user, roleBindings);
                            }

                        }
                    }
                }
                ctx.ExecuteQuery();
            }
        }

        public static void ChangePermissionForSite(string webFullUrl, IEnumerable<SpPermission> permissions = null)
        {
            using (var ctx = new ClientContext(webFullUrl))
            {
                var web = ctx.Web;
                ctx.Load(web);
                ctx.ExecuteQuery();
                if (permissions == null)
                {
                    web.ResetRoleInheritance();
                }
                else
                {
                    web.BreakRoleInheritance(false, false);
                    var spPermissions = permissions as SpPermission[] ?? permissions.ToArray();
                    foreach (var permission in spPermissions)
                    {
                        foreach (var userOrGroup in permission.UsersOrGroups)
                        {
                            var user = ctx.Web.EnsureUser(userOrGroup);
                            var roleDefinition = ctx.Site.RootWeb.RoleDefinitions.GetByType(permission.RoleType);
                            var roleBindings = new RoleDefinitionBindingCollection(ctx) { roleDefinition };
                            web.RoleAssignments.Add(user, roleBindings);
                        }

                    }
                }
                ctx.ExecuteQuery();
            }
        }
    }
}
