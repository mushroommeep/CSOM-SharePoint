using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Configuration.Json;
using Microsoft.SharePoint.Client;
using System;
using System.Threading.Tasks;
using System.Linq;
using Microsoft.SharePoint.Client.Taxonomy;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace CSOM_SharePoint
{
    public class CSOMHelper : IDisposable
    {
        private bool disposedValue;
        private const string termGroupName = "Site Collection - hatrannguyenngoc.sharepoint.com-sites-Naha";
        public CSOMHelper()
        {
        }

        public static async Task<Guid> CsomCreateTermSetAsync(ClientContext ctx, string termSetName)
        {
            try
            {
                // Get the TaxonomySession
                TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(ctx);
                // Get the term store by name
                TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
                // Get the term group by Name
                TermGroup termGroup = termStore.Groups.GetByName(termGroupName);
                ctx.Load(termGroup.TermSets);
                await ctx.ExecuteQueryAsync();

                bool has = termGroup.TermSets.Any(a => a.Name == termSetName);
                if (!has)
                {
                    //Set term set
                    Guid newTermSetGuid = Guid.NewGuid();
                    int lcid = 1033;
                    TermSet termSetColl = termGroup.CreateTermSet(termSetName, newTermSetGuid, lcid);

                    await ctx.ExecuteQueryAsync();
                   return newTermSetGuid;
                }
                else
                {
                    TermSet termSet = termGroup.TermSets.GetByName(termSetName);
                    ctx.Load(termSet);
                    await ctx.ExecuteQueryAsync();
                    return termSet.Id;
                }
                

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return Guid.Empty;
        }

        public static async Task<Guid> CsomCreateTermAsync(ClientContext ctx, string termName, string termSetName)
        {
            try
            {
                // Get the TaxonomySession
                TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(ctx);
                // Get the term store by name
                TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
                // Get the term group by Name
                TermGroup termGroup = termStore.Groups.GetByName(termGroupName);
                // Get the term set by Name
                TermSet termSet = termGroup.TermSets.GetByName(termSetName);

                var terms = termSet.GetAllTerms();

                ctx.Load(terms);
                await ctx.ExecuteQueryAsync();

                
                bool has = terms.Any(a => a.Name == termName);
                if (!has)
                {
                    // Int variable - new term lcid
                    int lcid = 1033;

                    // Guid - new term guid
                    Guid newTermId = Guid.NewGuid();

                    // Create a new term
                    Term newTerm = termSet.CreateTerm(termName, lcid, newTermId);
                    await ctx.ExecuteQueryAsync();
                    return newTermId;
                }
                else
                {
                    Term term = terms.GetByName(termName);
                    ctx.Load(term);
                    await ctx.ExecuteQueryAsync();
                    return term.Id;
                }


            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return Guid.Empty;
        }

        public static async Task CsomCreateList(ClientContext ctx, string listName, int templateType, string listDes = "New list description")
        {
            try
            {
                ListCollection listCollection = ctx.Web.Lists;

                ctx.Load(listCollection, ls => ls.Include(l => l.Title).Where(l => l.Title == listName));

                await ctx.ExecuteQueryAsync();

                if (listCollection.Count == 0)
                {
                    // The properties of the new custom list
                    ListCreationInformation creationInfo = new ListCreationInformation();
                    creationInfo.Title = listName;
                    creationInfo.Description = listDes;
                    creationInfo.TemplateType = templateType;

                    List newList = ctx.Web.Lists.Add(creationInfo);
                    ctx.Load(newList);

                    // Execute the query to the server.
                    await ctx.ExecuteQueryAsync();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

        }

        // create Site Field
        public static async Task CsomCreateSiteField(ClientContext ctx, string fieldName, string fieldType, Guid termSetId = default(Guid))
        {

            string fieldSchema = $"<Field Name='{fieldName}' DisplayName='{fieldName}' Type='{fieldType}' Hidden='False'/>";

            try
            {
                var fields = ctx.Web.Fields;
                ctx.Load(fields, ls => ls.Include(l => l.InternalName).Where(l => l.InternalName == fieldName));
                await ctx.ExecuteQueryAsync();

                if (fields.Count == 0)
                {
                    fields = ctx.Web.Fields;
                    //Adding site column to site  
                    var field = fields.AddFieldAsXml(fieldSchema, false, AddFieldOptions.AddFieldInternalNameHint);
                    //ctx.ExecuteQuery();

                    if (fieldType == "TaxonomyFieldTypeMulti" || fieldType == "TaxonomyFieldType")
                    {
                        TaxonomySession session = TaxonomySession.GetTaxonomySession(ctx);
                        TermStore termStore = session.GetDefaultSiteCollectionTermStore();
                        ctx.Load(termStore, ts => ts.Id);
                        ctx.ExecuteQuery();
                        TaxonomyField taxonomyField = ctx.CastTo<TaxonomyField>(field);
                        taxonomyField.SspId = termStore.Id;
                        taxonomyField.TermSetId = termSetId;
                        taxonomyField.TargetTemplate = String.Empty;
                        taxonomyField.AnchorId = Guid.Empty;
                        taxonomyField.Update();

                    }
                    field.Update();
                    ctx.Load(field);
                    await ctx.ExecuteQueryAsync();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
        }

        public static async Task CsomCreateSpContentType(ClientContext ctx, string contentTypesName)
        {
            try
            {
                var cntTypes = ctx.Web.ContentTypes;
                ctx.Load(cntTypes, ls => ls.Include(l => l.Name).Where(l => l.Name == contentTypesName));
                await ctx.ExecuteQueryAsync();

                if(cntTypes.Count == 0)
                {
                    ctx.Web.ContentTypes.Add(new ContentTypeCreationInformation
                    {
                        Name = contentTypesName
                }); 

                    await ctx.ExecuteQueryAsync();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        public static async Task CsomAddContentTypeToList(ClientContext ctx, string contentTypeName, string listName)
        {
            try
            {
                ListCollection lists = ctx.Web.Lists;
                ContentTypeCollection contentTypes = ctx.Web.ContentTypes;

                ctx.Load(lists, ls => ls.Include(l => l.Title).Where(l => l.Title == listName));
                ctx.Load(contentTypes, ls => ls.Include(l => l.Name).Where(l => l.Name == contentTypeName));

                await ctx.ExecuteQueryAsync();
                
                List list = lists.FirstOrDefault();
                ContentType contentType = contentTypes.FirstOrDefault();

                list.ContentTypesEnabled = true;
                ContentTypeCollection lcntTypes = list.ContentTypes;
                ctx.Load(lcntTypes);
                await ctx.ExecuteQueryAsync();

                bool has = false;
                foreach (ContentType ct in lcntTypes)
                {
                    if(ct.Name == contentTypeName)
                    {
                        has = true;
                        break;
                    };
                }

                if (!has) {
                    contentType.ReadOnly = false;
                    lcntTypes.AddExistingContentType(contentType); 
                }
                list.Update();
                ctx.Web.Update();
                await ctx.ExecuteQueryAsync();
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
        }

        public static async Task CsomAddFieldToContentType(ClientContext ctx, string contentTypeName, string fieldName)
        {
            FieldLinkCreationInformation fldLink = null;

            try
            {
                ContentTypeCollection contentTypes = ctx.Web.ContentTypes;
                ctx.Load(contentTypes, ls => ls.Include(l => l.Name).Where(l => l.Name == contentTypeName));
                await ctx.ExecuteQueryAsync();

                ContentType cntType = contentTypes.FirstOrDefault();

                FieldLinkCollection refFields = cntType.FieldLinks;
                ctx.Load(refFields);
                await ctx.ExecuteQueryAsync();

                foreach (var item in refFields)
                {
                    if (item.Name == fieldName)
                        return;
                }

                fldLink = new FieldLinkCreationInformation
                {
                    Field = ctx.Web.AvailableFields.GetByInternalNameOrTitle(fieldName)
                };
                cntType.FieldLinks.Add(fldLink);
                cntType.Update(true);
                ctx.Web.Update();
                await ctx.ExecuteQueryAsync();
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
        }

        public static async Task CsomChangeContentTypeOrder(ClientContext ctx, string listName, string cntTypeName)
        {
            List list = ctx.Web.Lists.GetByTitle(listName);
            ContentTypeCollection currentCtOrder = list.ContentTypes;
            ctx.Load(currentCtOrder);
            await ctx.ExecuteQueryAsync();

            List<ContentTypeId> reverseOrder = new List<ContentTypeId>();
            foreach (ContentType ct in currentCtOrder)
            {
                if (ct.Name.Equals(cntTypeName))
                {
                    reverseOrder.Add(ct.Id);
                }
            }
            list.RootFolder.UniqueContentTypeOrder = reverseOrder;
            list.RootFolder.Update();
            list.Update();
            await ctx.ExecuteQueryAsync();
        }

        public static async Task CsomCreateListItem(ClientContext ctx, string listName, string[][,] listItems, Dictionary<string, Guid> termIds)
        {
            try
            {
                List list = ctx.Web.Lists.GetByTitle(listName);

                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();

                foreach(var l in listItems)
                {
                    ListItem oListItem = list.AddItem(itemCreateInfo);
                    for (int i = 0; i < l.GetLength(0); i++)
                    {
                        Field oField = list.Fields.GetByInternalNameOrTitle(l[i, 0]);
                        ctx.Load(oField);
                        await ctx.ExecuteQueryAsync();

                        oListItem["Title"] = "Test Title " + i;
                        oListItem.Update();

                        ctx.Load(oListItem);
                        ctx.ExecuteQuery();

                        if (oField.TypeAsString == "TaxonomyFieldType" || oField.TypeAsString == "TaxonomyFieldTypeMulti") 
                        {
                            await CsomSetTaxonomyFieldValue(ctx, list, oListItem, l[i, 0], l[i, 1], termIds, oField.TypeAsString);
                        }
                        else
                        {
                            oListItem[l[i, 0]] = l[i, 1];
                            oListItem.Update();
                        }
                    }
                }

                await ctx.ExecuteQueryAsync();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

        }

        public static async Task CsomSetDefaultSiteField(ClientContext ctx, string listName, string fieldName, string defaultName = "", Guid defaultTermId = default(Guid))
        {
            try
            {
                List list = ctx.Web.Lists.GetByTitle(listName);

                // Get field from list using internal name or display name
                Field oField = list.Fields.GetByInternalNameOrTitle(fieldName);
                ctx.Load(oField);
                await ctx.ExecuteQueryAsync();

                if(oField.TypeAsString == "TaxonomyFieldType")
                {
                    var taxField = ctx.CastTo<TaxonomyField>(oField);
                    ctx.Load(taxField);
                    await ctx.ExecuteQueryAsync();

                    //initialize taxonomy field value
                    var defaultValue = new TaxonomyFieldValue();
                    defaultValue.WssId = -1;
                    defaultValue.Label = fieldName;
                    defaultValue.TermGuid = defaultTermId.ToString();

                    //retrieve validated taxonomy field value
                    var validatedValue = taxField.GetValidatedString(defaultValue);
                    await ctx.ExecuteQueryAsync();

                    //set default value for a taxonomy field
                    taxField.DefaultValue = validatedValue.Value;
                    taxField.Update();
                    await ctx.ExecuteQueryAsync();
                }
                else
                {
                    // Set field default value
                    oField.DefaultValue = defaultName;
                }

                oField.Update();
                await ctx.ExecuteQueryAsync();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

        }

        public static async Task CsomCreateListView(ClientContext ctx, string listName)
        {
            try
            {
                List list = ctx.Web.Lists.GetByTitle(listName);

                ViewCollection viewCollection = list.Views;
                ctx.Load(viewCollection);

                ViewCreationInformation viewCreationInformation = new ViewCreationInformation();
                viewCreationInformation.Title = "CSOM Test View";

                viewCreationInformation.ViewTypeKind = ViewType.Html;

                viewCreationInformation.Query = "<OrderBy><FieldRef Name='ID' Ascending='True'/></OrderBy><Where><Eq><FieldRef Name = 'taxCity' /><Value Type = 'TaxonomyFieldType'>Ho Chi Minh</Value></Eq></Where>";

                string CommaSeparateColumnNames = "ID,Title,taxCity,about";
                viewCreationInformation.ViewFields = CommaSeparateColumnNames.Split(',');

                View listView = viewCollection.Add(viewCreationInformation);
                await ctx.ExecuteQueryAsync();

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

        }

        public static async Task CsomBatchUpdateTextCol(ClientContext ctx, string listName, string fieldName, string updValue, string newValue)
        {
            try
            {
                int chunkSize = 2;
                int start = 0;
                int end;

                List list = ctx.Web.Lists.GetByTitle(listName);

                CamlQuery oQuery = new CamlQuery();
                oQuery.ViewXml = $@"<View><Query><Where>
                                                <Eq>
                                                <FieldRef Name='{fieldName}' />
                                                <Value Type='Text'>{updValue}</Value>
                                                </Eq>
                                                </Where></Query></View>";

                ListItemCollection oItems = list.GetItems(oQuery);

                ctx.Load(oItems);
                await ctx.ExecuteQueryAsync();

                int count = oItems.Count;
                end = start + chunkSize < count ? start + chunkSize : count;

                foreach (ListItem oListItem in oItems)
                {
                    start++;
                    oListItem[fieldName] = newValue;
                    oListItem.Update();
                    if(start == end)
                    {
                        start = end;
                        end = start + chunkSize < count ? start + chunkSize : count;
                        await ctx.ExecuteQueryAsync();
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

        }

        private static async Task CsomSetTaxonomyFieldValue( ClientContext ctx, List list, ListItem item, string fieldName, string label, Dictionary<string, Guid> termIds, string type)
        {
            try
            {
                var clientRuntimeContext = item.Context;
                var field = list.Fields.GetByTitle(fieldName);
                ctx.Load(field);
                await ctx.ExecuteQueryAsync();
                var taxField = clientRuntimeContext.CastTo<TaxonomyField>(field);

                if(type == "TaxonomyFieldType") //Single value
                {
                    taxField.SetFieldValueByValue(item, new TaxonomyFieldValue()
                    {
                        WssId = -1, // alway let it -1
                        Label = label,
                        TermGuid = termIds[label].ToString()
                    });
                }
                else if(type == "TaxonomyFieldTypeMulti") //multiple value
                {
                    string tagsString = String.Empty;
                    string[] labels = label.Split('|');
                    foreach (string lb in labels)
                    {
                        tagsString += $"-1;#{lb}|{termIds[lb].ToString()};#";
                    }
                    tagsString = tagsString.TrimEnd(new char[] { ';', '#' });
                    taxField.SetFieldValueByValueCollection(item, 
                        new TaxonomyFieldValueCollection(clientRuntimeContext, tagsString, taxField));
                }
                
                taxField.Update();
                item.Update();
                await ctx.ExecuteQueryAsync();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }

        public static async Task CsomCreateListColumn(ClientContext ctx, string listName, string fieldName, string fieldType, Guid termSetId = default(Guid))
         {

            string fieldSchema = $"<Field Name='{fieldName}' DisplayName='{fieldName}' Type='{fieldType}' Hidden='False'/>";

            try
            {
                List list = ctx.Web.Lists.GetByTitle(listName);

                ctx.Load(list.Fields, ls => ls.Include(l => l.InternalName).Where(l => l.InternalName == fieldName));
                await ctx.ExecuteQueryAsync();

                if (list.Fields != null && list.Fields.Count == 0)
                {
                    //Adding column to list
                    Field field = list.Fields.AddFieldAsXml(fieldSchema, false, AddFieldOptions.AddFieldInternalNameHint);
                    //ctx.ExecuteQuery();

                    if (fieldType == "TaxonomyFieldTypeMulti" || fieldType == "TaxonomyFieldType")
                    {
                        TaxonomySession session = TaxonomySession.GetTaxonomySession(ctx);
                        TermStore termStore = session.GetDefaultSiteCollectionTermStore();
                        ctx.Load(termStore, ts => ts.Id);
                        ctx.ExecuteQuery();
                        TaxonomyField taxonomyField = ctx.CastTo<TaxonomyField>(field);
                        taxonomyField.SspId = termStore.Id;
                        taxonomyField.TermSetId = termSetId;
                        taxonomyField.TargetTemplate = String.Empty;
                        taxonomyField.AnchorId = Guid.Empty;
                        taxonomyField.Update();
                    }

                    field.Update();
                    ctx.Load(field);
                    await ctx.ExecuteQueryAsync();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
        }

        public static async Task CsomUpdateUserColAdmin(ClientContext ctx, string listName, string fieldName)
        {
            try
            {
                int chunkSize = 2;
                int start = 0;
                int end;
                List list = ctx.Web.Lists.GetByTitle(listName);

                CamlQuery oQuery = CamlQuery.CreateAllItemsQuery();

                ListItemCollection oItems = list.GetItems(oQuery);
                Site oSite = ctx.Site;

                Field field = list.Fields.GetByTitle(fieldName);
                ctx.Load(field);
                ctx.Load(oSite, site => site.Owner);
                ctx.Load(oItems);

                await ctx.ExecuteQueryAsync();

                FieldUserValue userValue = new FieldUserValue();
                userValue.LookupId = oSite.Owner.Id;

                int count = oItems.Count;
                end = start + chunkSize < count ? start + chunkSize : count;

                foreach (ListItem oListItem in oItems)
                {
                    start++;
                    string itemInternalName = field.InternalName;
                    oListItem[itemInternalName] = userValue;
                    oListItem.Update();
                    if (start == end)
                    {
                        start = end;
                        end = start + chunkSize < count ? start + chunkSize : count;
                        await ctx.ExecuteQueryAsync();
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

        }

        private static string CsomCreateTaxonomyFieldXml(Guid termStoreId, Guid termSetId, string fieldName, string fieldType)
        {
            Guid txtFieldId = Guid.NewGuid();
            Guid taxFieldId = Guid.NewGuid();
            
            //Single valued, or multiple choice?
            bool isMulti = fieldType == "TaxonomyFieldTypeMulti"? true : false;
            //If it's single value, index it.
            string mult = isMulti ? "Mult='TRUE'" : "Indexed='TRUE'";

            string fieldSchema = $"<Field Type='{fieldType}' DisplayName='{fieldName}' ID='{taxFieldId.ToString("B")}' " +
                                       $"ShowField='Term1033' Required='FALSE' " +
                                       $"EnforceUniqueValues='FALSE' {mult} Sortable='FALSE' Name='{ fieldName.Replace(" ", "")}' >" +
                                       $"<Default/><Customization><ArrayOfProperty><Property>" +
                                       $"<Name>SspId</Name><Value xmlns:q1='http://www.w3.org/2001/XMLSchema' p4:type='q1:string' xmlns:p4='http://www.w3.org/2001/XMLSchema-instance'>{termStoreId.ToString("D")}</Value>" +
                                       $"</Property><Property><Name>GroupId</Name></Property><Property>" +
                                       $"<Name>TermSetId</Name><Value xmlns:q2='http://www.w3.org/2001/XMLSchema' p4:type='q2:string' xmlns:p4='http://www.w3.org/2001/XMLSchema-instance'>{termSetId.ToString("D")}</Value>" +
                                       $"</Property><Property><Name>AnchorId</Name><Value xmlns:q3='http://www.w3.org/2001/XMLSchema' p4:type='q3:string' xmlns:p4='http://www.w3.org/2001/XMLSchema-instance'>00000000-0000-0000-0000-000000000000</Value>" +
                                       $"</Property><Property><Name>UserCreated</Name><Value xmlns:q4='http://www.w3.org/2001/XMLSchema' p4:type='q4:boolean' xmlns:p4='http://www.w3.org/2001/XMLSchema-instance'>false</Value>" +
                                       $"</Property><Property><Name>Open</Name><Value xmlns:q5='http://www.w3.org/2001/XMLSchema' p4:type='q5:boolean' xmlns:p4='http://www.w3.org/2001/XMLSchema-instance'>false</Value>" +
                                       $"</Property><Property><Name>TextField</Name><Value xmlns:q6='http://www.w3.org/2001/XMLSchema' p4:type='q6:string' xmlns:p4='http://www.w3.org/2001/XMLSchema-instance'>{txtFieldId.ToString("B")}</Value>" +
                                       $"</Property><Property><Name>IsPathRendered</Name><Value xmlns:q7='http://www.w3.org/2001/XMLSchema' p4:type='q7:boolean' xmlns:p4='http://www.w3.org/2001/XMLSchema-instance'>true</Value>" +
                                       $"</Property><Property><Name>IsKeyword</Name><Value xmlns:q8='http://www.w3.org/2001/XMLSchema' p4:type='q8:boolean' xmlns:p4='http://www.w3.org/2001/XMLSchema-instance'>false</Value>" +
                                       $"</Property><Property><Name>TargetTemplate</Name></Property><Property><Name>CreateValuesInEditForm</Name><Value xmlns:q9='http://www.w3.org/2001/XMLSchema' p4:type='q9:boolean' xmlns:p4='http://www.w3.org/2001/XMLSchema-instance'>false</Value>" +
                                       $"</Property><Property><Name>FilterAssemblyStrongName</Name><Value xmlns:q10='http://www.w3.org/2001/XMLSchema' p4:type='q10:string' xmlns:p4='http://www.w3.org/2001/XMLSchema-instance'>Microsoft.SharePoint.Taxonomy, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c</Value>" +
                                       $"</Property><Property><Name>FilterClassName</Name><Value xmlns:q11='http://www.w3.org/2001/XMLSchema' p4:type='q11:string' xmlns:p4='http://www.w3.org/2001/XMLSchema-instance'>Microsoft.SharePoint.Taxonomy.TaxonomyField</Value>" +
                                       $"</Property><Property><Name>FilterMethodName</Name><Value xmlns:q12='http://www.w3.org/2001/XMLSchema' p4:type='q12:string' xmlns:p4='http://www.w3.org/2001/XMLSchema-instance'>GetFilteringHtml</Value>" +
                                       $"</Property><Property><Name>FilterJavascriptProperty</Name><Value xmlns:q13='http://www.w3.org/2001/XMLSchema' p4:type='q13:string' xmlns:p4='http://www.w3.org/2001/XMLSchema-instance'>FilteringJavascript</Value></Property></ArrayOfProperty></Customization></Field>";


                return fieldSchema;
        }

        private static string CsomCreateFieldXml(string fieldName, string fieldType)
        {
            string fieldSchema = $"<Field Name='{fieldName}' DisplayName='{fieldName}' Type='{fieldType}' Hidden='False'/>";

            return fieldSchema;
        }

        public static async Task CsomCreateSiteFieldWithXml(ClientContext ctx, string fieldName, string fieldType, Guid termSetId = default(Guid))
        {
            string fieldSchema = String.Empty;
            try
            {
                var fields = ctx.Web.Fields;
                ctx.Load(fields, ls => ls.Include(l => l.InternalName).Where(l => l.InternalName == fieldName));
                await ctx.ExecuteQueryAsync();

                if (fields.Count == 0)
                {
                    fields = ctx.Web.Fields;
                    if (fieldType == "TaxonomyFieldTypeMulti" || fieldType == "TaxonomyFieldType")
                    {
                        TaxonomySession session = TaxonomySession.GetTaxonomySession(ctx);
                        TermStore termStore = session.GetDefaultSiteCollectionTermStore();
                        ctx.Load(termStore);
                        await ctx.ExecuteQueryAsync();
                        fieldSchema = CsomCreateTaxonomyFieldXml(termStore.Id, termSetId, fieldName, fieldType);
                    }
                    else
                    {
                        fieldSchema = CsomCreateFieldXml(fieldName, fieldType);
                    }
                        //Adding site column to site  
                    var field = fields.AddFieldAsXml(fieldSchema, false, AddFieldOptions.AddFieldInternalNameHint);
                    ctx.Load(field);
                    await ctx.ExecuteQueryAsync();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
        }

        public static async Task CsomAddFieldToListContentType(ClientContext ctx, string contentTypeName, string fieldName, string listName)
        {
            FieldLinkCreationInformation fldLink = null;

            try
            {
                List list = ctx.Web.Lists.GetByTitle(listName);
                ContentTypeCollection contentTypes = list.ContentTypes;
                ctx.Load(contentTypes, ls => ls.Include(l => l.Name).Where(l => l.Name == contentTypeName));
                await ctx.ExecuteQueryAsync();

                ContentType cntType = contentTypes.FirstOrDefault();

                FieldLinkCollection refFields = cntType.FieldLinks;
                ctx.Load(refFields);
                await ctx.ExecuteQueryAsync();

                foreach (var item in refFields)
                {
                    if (item.Name == fieldName)
                        return;
                }

                fldLink = new FieldLinkCreationInformation
                {
                    Field = ctx.Web.AvailableFields.GetByInternalNameOrTitle(fieldName)
                };
                cntType.FieldLinks.Add(fldLink);
                cntType.Update(false);
                await ctx.ExecuteQueryAsync();
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
        }

        public static async Task CsomCreateListItemFolder(ClientContext ctx, string listName, string folderName, string parent)
        {
            try
            {
                List list = ctx.Web.Lists.GetByTitle(listName);
                //Enable Folder creation for the list
                list.EnableFolderCreation = true;
                list.Update();

                FolderCollection folders = list.RootFolder.Folders;
                ctx.Load(folders);

                await ctx.ExecuteQueryAsync();

                if(!folders.Any(x => x.Name == folderName))
                {
                    if(parent == "Root")
                    {
                        ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                        itemCreateInfo.UnderlyingObjectType = FileSystemObjectType.Folder;

                        // This will et the internal name/path of the file/folder
                        itemCreateInfo.LeafName = folderName;
                        ListItem listItem = list.AddItem(itemCreateInfo);

                        // Set folder Name
                        listItem["Title"] = "Folder Name";

                        listItem.Update();
                    }
                    else
                    {
                        foreach(Folder f in folders)
                        {
                            if(f.Name == parent)
                            {
                                f.Folders.Add(folderName);
                                f.Update();
                            }
                        }
                        
                    }

                    await ctx.ExecuteQueryAsync();
                }
                
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

        }

        public static async Task CsomCreateDocumentLibItem(ClientContext ctx, string listName, string folderName, string[][,] listItems, Dictionary<string, Guid> termIds)
        {
            try
            {
                List list = ctx.Web.Lists.GetByTitle(listName);

                foreach (var l in listItems)
                {
                    FileCreationInformation createFile = new FileCreationInformation();
                    int filePos = 0;
                    for (int i = 0; i < l.GetLength(0); i++)
                    {

                        if (l[i, 0] == "file" & !System.IO.File.Exists(l[i, 1]))
                        {
                            createFile.Url = l[i, 1];
                            //use byte array to set content of the file
                            string somestring = "hello there";
                            byte[] toBytes = Encoding.ASCII.GetBytes(somestring);

                            createFile.Content = toBytes;
                            filePos = i;
                            break;
                        }
                        else if (l[i, 0] == "file" & System.IO.File.Exists(l[i, 1]))
                        {
                            createFile.Url = System.IO.Path.GetFileName(l[i, 1]);
                            createFile.Content = System.IO.File.ReadAllBytes(l[i, 1]);
                            filePos = i;
                            break;
                        }
                    }

                    List<Folder> folders = GetAllFolders(list);

                    Microsoft.SharePoint.Client.File addedFile = null;

                    foreach (Folder f in folders)
                    {
                        if (f.Name == folderName)
                        {
                            // havent check if file exist in folder -> check or use exception handler
                            addedFile = f.Files.Add(createFile);
                            f.Update();
                            ctx.Load(addedFile);
                            await ctx.ExecuteQueryAsync();
                        }
                    }
                    
                    if(addedFile != null)
                    {
                        ListItem oListItem  = addedFile.ListItemAllFields;
                        for (int i = 0; i < l.GetLength(0); i++)
                        {
                            if(filePos != i)
                            {
                                Field oField = list.Fields.GetByInternalNameOrTitle(l[i, 0]);
                                ctx.Load(oField);
                                await ctx.ExecuteQueryAsync();

                                oListItem["Title"] = "Test Title " + i;
                                oListItem.Update();

                                ctx.Load(oListItem);
                                ctx.ExecuteQuery();

                                if (oField.TypeAsString == "TaxonomyFieldType" || oField.TypeAsString == "TaxonomyFieldTypeMulti")
                                {
                                    await CsomSetTaxonomyFieldValue(ctx, list, oListItem, l[i, 0], l[i, 1], termIds, oField.TypeAsString);
                                }
                                else
                                {
                                    oListItem[l[i, 0]] = l[i, 1];
                                    oListItem.Update();
                                }
                            }
                        }
                    }
                }

                await ctx.ExecuteQueryAsync();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

        }

        private static List<Folder> GetAllFolders(List list)
        {
            var ctx = list.Context;
            var folderItems = list.GetItems(CamlQuery.CreateAllFoldersQuery());
            ctx.Load(folderItems, icol => icol.Include(i => i.Folder));
            ctx.ExecuteQuery();
            var allFolders = folderItems.Select(i => i.Folder).ToList();
            return allFolders;
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    //Dispose code
                }

                disposedValue = true;
            }
        }

        public void Dispose()
        {
            // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }
}
