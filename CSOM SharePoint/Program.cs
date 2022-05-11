using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Configuration.Json;
using Microsoft.SharePoint.Client;
using System;
using System.Threading.Tasks;
using System.Linq;
using Microsoft.SharePoint.Client.Taxonomy;
using System.Collections.Generic;

namespace CSOM_SharePoint
{
    internal class Program
    {
        class SharepointInfo
        {
            public string SiteUrl { get; set; }
            public string Username { get; set; }
            public string Password { get; set; }
        }

        static async Task Main(string[] args)
        {
            try
            {
                using (var clientContexthelper = new ClientContextHelper())
                {
                    Guid termSetId = Guid.Empty;
                    Dictionary<string, Guid> termId = new Dictionary<string, Guid>();

                    ClientContext ctx = GetContext(clientContexthelper);
                    ctx.Load(ctx.Web);
                    await ctx.ExecuteQueryAsync();

                    Console.WriteLine($"Site {ctx.Web.Title}");

                    string termSetName = "city-hatr";
                    string[] terms = new string[] { "Ho Chi Minh", "Stockholm" };

                    //Using CSOM create a List name “CSOM Test”
                    await CSOMHelper.CsomCreateList(ctx, "CSOM Test", 100);

                    //Create term set “city-{yourname}” in dev tenant
                    termSetId = await CSOMHelper.CsomCreateTermSetAsync(ctx, termSetName);

                    //Create 2 terms “Ho Chi Minh” and “Stockholm” in termset “city-{yourname}”
                    foreach (string city in terms)
                    {
                        termId.Add(city, await CSOMHelper.CsomCreateTermAsync(ctx, city, termSetName));

                    }
                    ////Create site fields “about” type text 
                    //await CSOMHelper.CsomCreateSiteField(ctx, "about", "Text");

                    ////Create site fields “city” type taxonomy 
                    //await CSOMHelper.CsomCreateSiteField(ctx, "taxCity", "TaxonomyFieldType", termSetId);

                    ////Create site content type “CSOM Test content type” 
                    //await CSOMHelper.CsomCreateSpContentType(ctx, "CSOM Test content type");

                    ////add this content type to list “CSOM test”
                    //await CSOMHelper.CsomAddContentTypeToList(ctx, "CSOM Test content type", "CSOM Test");

                    ////add fields “about” and “city” to this content type
                    //string[] fields = new string[] { "about", "taxCity" };
                    //foreach (string f in fields)
                    //{
                    //    await CSOMHelper.CsomAddFieldToContentType(ctx, "CSOM Test content type", f);
                    //}

                    ////In list “CSOM test” set “CSOM Test content type” as default content type
                    //await CSOMHelper.CsomChangeContentTypeOrder(ctx, "CSOM Test", "CSOM Test content type");

                    ////Create 5 list items to list with some value in field “about” and “city”
                    //string[][,] listItems = new string[][,]
                    //{
                    //    new string[2,2]{   { "taxCity", "Ho Chi Minh" }, { "about", "about1"} },
                    //    new string[2,2]{   { "taxCity", "Stockholm" }, { "about", "about2"} },
                    //    new string[2,2]{   { "taxCity", "Stockholm" }, { "about", "about3"} },
                    //    new string[2,2]{   { "taxCity", "Ho Chi Minh" }, { "about", "about4"} },
                    //    new string[2,2]{   { "taxCity", "Stockholm" }, { "about", "about5" } },
                    //};
                    //await CSOMHelper.CsomCreateListItem(ctx, "CSOM Test", listItems, termId);

                    ////Update site field “about” set default value for it to “about default” then create 2 new list items.
                    //await CSOMHelper.CsomSetDefaultSiteField(ctx, "CSOM Test", "about", "about default");

                    //string[][,] listItems2 = new string[][,]
                    //{
                    //    new string[1,2]{   { "taxCity", "Ho Chi Minh" } },
                    //    new string[1,2]{   { "taxCity", "Stockholm" } },
                    //};
                    //await CSOMHelper.CsomCreateListItem(ctx, "CSOM Test", listItems2, termId);

                    ////Update site field “city” set default value for it to “Ho Chi Minh” then create 2 new list items.
                    //await CSOMHelper.CsomSetDefaultSiteField(ctx, "CSOM Test", "taxCity", defaultTermId: termId["Ho Chi Minh"]);

                    //string[][,] listItems3 = new string[][,]
                    //{
                    //    new string[1,2]{   { "about", "Banana" } },
                    //    new string[1,2]{   { "about", "Potato" } },
                    //};
                    //await CSOMHelper.CsomCreateListItem(ctx, "CSOM Test", listItems3, termId);


                    ///*EXX 2*/
                    ////Write CAML query to get list items where field “about” is not “about default”
                    //await SimpleCamlQueryAsync(ctx);

                    ////Create List View by CSOM 
                    //await CSOMHelper.CsomCreateListView(ctx, "CSOM Test");

                    ////update list items in batch
                    //await CSOMHelper.CsomBatchUpdateTextCol(ctx, "CSOM Test", "about", "about default", "Update script");

                    ////Create List column
                    //await CSOMHelper.CsomCreateListColumn(ctx, "CSOM Test", "author", "User");

                    ////
                    //await CSOMHelper.CsomUpdateUserColAdmin(ctx, "CSOM Test", "author");


                    ///*EXX 3*/
                    ////Create Taxonomy Field which allow multi values, with name “cities” map to your termset.
                    //await CSOMHelper.CsomCreateSiteFieldWithXml(ctx, "cities", "TaxonomyFieldTypeMulti", termSetId);

                    ////Add field “cities” to content type “CSOM Test content type” make sure don’t need update list but added field should be available in your list “CSOM test”
                    //await CSOMHelper.CsomAddFieldToContentType(ctx, "CSOM Test content type", "cities");

                    ////Add 3 list item to list “CSOM test” and set multi value to field “cities” 
                    //string[][,] listItems4 = new string[][,]
                    //{
                    //    new string[3,2]{   { "taxCity", "Ho Chi Minh" }, { "about", "aboutEx31"}, {"cities", "Ho Chi Minh" } },
                    //    new string[3,2]{   { "taxCity", "Stockholm" }, { "about", "aboutEx32" },  {"cities", "Stockholm" } },
                    //    new string[3,2]{   { "taxCity", "Stockholm" }, { "about", "aboutEx33" }, {"cities", "Stockholm|Ho Chi Minh" } }
                    //};
                    //await CSOMHelper.CsomCreateListItem(ctx, "CSOM Test", listItems4, termId);

                    ////Create new List type Document lib name “Document Test” 
                    //await CSOMHelper.CsomCreateList(ctx, "Document Test", 101);

                    ////add content type “CSOM Test content type” to this list.
                    //await CSOMHelper.CsomAddContentTypeToList(ctx, "CSOM Test content type", "Document Test");

                    ////Create Folder “Folder 1” in root of list “Document Test” 
                    //await CSOMHelper.CsomCreateListItemFolder(ctx, "Document Test", "Folder 1", "Root");

                    ////create “Folder 2” inside “Folder 1” 
                    //await CSOMHelper.CsomCreateListItemFolder(ctx, "Document Test", "Folder 2", "Folder 1");

                    ////Create 3 list items in “Folder 2” with value “Folder test” in field “about”. 
                    //string[][,] listItems5 = new string[][,]
                    //{
                    //    new string[2,2]{   { "file", "file1.txt" }, { "about", "Folder test"}},
                    //    new string[2,2]{   { "file", "file2.doc" }, { "about", "Folder test" } },
                    //    new string[2,2]{   { "file", "file3.docx" }, { "about", "Folder test" } }
                    //};
                    //await CSOMHelper.CsomCreateDocumentLibItem(ctx, "Document Test", "Folder 2", listItems5, termId);

                    ////Create 2 flies in “Folder 2” with value “Stockholm” in field “cities”.
                    //string[][,] listItems6 = new string[][,]
                    //{
                    //    new string[2,2]{   { "file", "file1Stockholm.txt" }, { "cities", "Stockholm" } },
                    //    new string[2,2]{   { "file", "file2Stockholm.doc" }, { "cities", "Stockholm" } }
                    //};
                    //await CSOMHelper.CsomCreateDocumentLibItem(ctx, "Document Test", "Folder 2", listItems6, termId);

                    //Create List Item in “Document Test” by upload a file Document.docx 
                    string[][,] listItems7 = new string[][,]
                    {
                        new string[3,2]{{ "file", "D:\\Exercise\\CSOM SharePoint\\doc\\Document.docx" }, { "cities", "Ho Chi Minh|Stockholm" }, { "about", "Folder test"} },
                    };
                    await CSOMHelper.CsomCreateDocumentLibItem(ctx, "Document Test", "Folder 1", listItems7, termId);

                    //Write CAML get all list item just in “Folder 2” and have value “Stockholm” in “cities” field
                    await CamlGetAllListItem(ctx, "Document Test");
                }
                Console.WriteLine($"Press Any Key To Stop!");
                Console.ReadKey();
            }
            catch (Exception ex) 
            {
                Console.WriteLine(ex.Message);
            }
        }

        static ClientContext GetContext(ClientContextHelper clientContextHelper)
        {
            var builder = new ConfigurationBuilder().AddJsonFile($"appsettings.json", true, true);
            IConfiguration config = builder.Build();
            var info = config.GetSection("SharepointInfo").Get<SharepointInfo>();
            return clientContextHelper.GetContext(new Uri(info.SiteUrl), info.Username, info.Password);
        }

        private static async Task SimpleCamlQueryAsync(ClientContext ctx)
        {
            try
            {

                var list = ctx.Web.Lists.GetByTitle("CSOM Test");

                var items = list.GetItems(new CamlQuery()
                {
                    ViewXml = "<View><Query><Where><Neq><FieldRef Name='about' /><Value Type='Text'>about default</Value></Neq></Where></Query><ViewFields><FieldRef Name='about' /></ViewFields><QueryOptions /></View>"
                });

                ctx.Load(items);
                await ctx.ExecuteQueryAsync();

                foreach (var item in items)
                {
                    Console.WriteLine(item["about"]);
                }
                    } catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }

        private static async Task CamlGetAllListItem (ClientContext ctx, string listName)
        {
            try
            {

                var list = ctx.Web.Lists.GetByTitle(listName);

                var items = list.GetItems(new CamlQuery()
                {
                    ViewXml = "<View Scope='RecursiveAll'><Query><Where><And><Eq><FieldRef Name='FileDirRef' /><Value Type='Text'>/sites/Naha/Document Test/Folder 1/Folder 2</Value></Eq><Eq><FieldRef Name='cities' /><Value Type='TaxonomyFieldTypeMulti'>Stockholm</Value></Eq></And></Where></Query><ViewFields><FieldRef Name='FileLeafRef' /></ViewFields><QueryOptions /></View>"
                });

                ctx.Load(items);
                await ctx.ExecuteQueryAsync();

                foreach (var item in items)
                {
                    Console.WriteLine(item["FileLeafRef"]);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }
    }
}
