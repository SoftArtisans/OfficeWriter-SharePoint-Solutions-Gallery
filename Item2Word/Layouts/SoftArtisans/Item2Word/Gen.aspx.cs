using System;
using Microsoft.SharePoint;
using Microsoft.SharePoint.WebControls;
using System.Data;
using SoftArtisans.OfficeWriter.WordWriter;
using System.Collections.Generic;
using System.Linq;

namespace Item2Word.Layouts.SoftArtisans.Item2Word
{
    public partial class Gen : LayoutsPageBase
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            try
            {
                //load the web object for the site that the page is now in context of
                using (SPWeb web = SPControl.GetContextWeb(Context))
                {
                    //load the list that was passed in the 'list' querystring parameter to the page
                    SPList list = web.Lists[new Guid(Page.Request.QueryString["list"])];
                    SPFile template = Web.GetFile(Page.Request.QueryString["TemplateLocation"]);
                    SPListItem item = list.GetItemById(int.Parse(Page.Request.QueryString["Item"]));
                    Dictionary<string, object> dat = getData(item);
                    WordTemplate wt = new WordTemplate();


                    
                    wt.Open(template.OpenBinaryStream());
                    

                    wt.SetDataSource(dat.Values.ToArray(), dat.Keys.ToArray(), "data");
                    wt.Process();
                    wt.Save(Page.Response, template.Name, false);



                }
            }
            catch (Exception ex)
            {
                ErrorText.Text = ex.Message;
            }


        }

        private Dictionary<string, object> getData(SPListItem item)
        {
            
            
            Dictionary<string, object> dat = new Dictionary<string, object>();

            foreach (SPField field in item.Fields)
            {

                string colName = field.Title;
                if (dat.ContainsKey(colName))
                {
                    int i = 1;
                    while (dat.ContainsKey(colName + "_" + i.ToString()))
                    {
                        i++;
                    }
                    colName = colName + "_" + i.ToString();
                }
                dat.Add(colName,item[field.Id]);


            }


            return dat;


        }


    }
}
