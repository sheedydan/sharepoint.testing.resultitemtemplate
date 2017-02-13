using Microsoft.Office.Server.Search.Administration;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace sharepoint.testing.resultitemtype
{
    class Program
    {
        static void Main(string[] args)
        {
            using (SPSite site = new SPSite("http://portal.spdev16.com/one-stop-shop"))
            {
                using (SPWeb web = site.OpenWeb())
                {
                    SPServiceContext serviceContext = SPServiceContext.GetContext(site);

                    var searchApplicationProxy = (SearchServiceApplicationProxy)serviceContext.GetDefaultProxy(typeof(SearchServiceApplicationProxy));
                    //var owner = new SearchObjectOwner(SearchObjectLevel.SPSite, web);
                    var owner = new SearchObjectOwner(SearchObjectLevel.SPWeb, web);

                    ICollection<ResultItemType> itemTypes=  searchApplicationProxy.GetResultItemTypes(null, null, owner, true);

                    ResultItemType item = CreateResultType(web, "Internal Resource - Deployed", "/_catalogs/masterpage/Display Templates/Search/Item_InternalResource.js",
                        new PropertyRule[] {
                            CustomPropertyRule("ContentTypeId", PropertyRuleOperator.DefaultOperator.Contains, new string[] { "41E5CAB13DEA48B8A69A7A47EA4F4EE4" })
                        }, false);

                    searchApplicationProxy.AddResultItemType(item);
                }
            }
        }
        public static PropertyRule CustomPropertyRule(string propertyName, PropertyRuleOperator.DefaultOperator propertyOperator, string[] values)
        {
            Type type = typeof(PropertyRuleOperator);
            PropertyInfo info = type.GetProperty("DefaultOperators", BindingFlags.NonPublic | BindingFlags.Static);
            object value = info.GetValue(null);
            var DefaultOperators = (Dictionary<PropertyRuleOperator.DefaultOperator, PropertyRuleOperator>)value;
            PropertyRule rule = new PropertyRule(propertyName, DefaultOperators[propertyOperator]);
            rule.PropertyValues = new List<string>(values);
            return rule;
        }
       public static ResultItemType CreateResultType(SPWeb web, string name, string displayTemplateUrl, PropertyRule[] rules, bool optimizeForFrequentUse)
        {
            ResultItemType resType = new ResultItemType(new SearchObjectOwner(SearchObjectLevel.SPWeb, web));
            resType.Name = name;
            resType.SourceID = new Guid();
            resType.DisplayTemplateUrl = "~sitecollection" + displayTemplateUrl;
            SPFile file = web.GetFile(SPUtility.ConcatUrls(web.ServerRelativeUrl, displayTemplateUrl));
            //resType.DisplayProperties = ParseManagePropertyMappings(file.ListItemAllFields["Managed Property Mappings"].ToString());
            resType.Rules = new PropertyRuleCollection(new List<PropertyRule>(rules));
            typeof(ResultItemType).GetProperty("OptimizeForFrequentUse").SetValue(resType, optimizeForFrequentUse);
            return resType;
        }
        private static string ParseManagePropertyMappings(string mappings)
        {
            string[] propArray = mappings.Replace("'", "").Replace("\"", "").Split(',');
            for (int i = 0; i < propArray.Length; i++)
            {
                if (propArray[i].Contains(":"))
                {
                    int n = propArray[i].LastIndexOf(':');
                    if ((n > 0) && (n < (propArray[i].Length - 1)))
                    {
                        propArray[i] = propArray[i].Substring(n + 1);
                    }
                }
                propArray[i] = propArray[i].Replace(";", ",");
            }
            return string.Join(",", propArray);
        }

    }
}
