using System;
using Microsoft.SharePoint;
using Microsoft.SharePoint.WebControls;
using SoftArtisans.OfficeWriter.ExcelWriter;
using System.Data;

namespace List2Excel.Layouts.SoftArtisans.List2Excel
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
                    DataTable dat = getData(list);
                    ExcelTemplate xlt = new ExcelTemplate();
                    xlt.Open(template.OpenBinaryStream());
                    xlt.BindData(dat, "data", xlt.CreateDataBindingProperties());
                    xlt.Process();
                    xlt.Save(Page.Response, template.Name, false);
                    

                }
            }
            catch(Exception ex)
            {
                ErrorText.Text = ex.Message;
            }


        }

        private DataTable getData(SPList list)
        {
            DataTable dat = list.Items.GetDataTable();
            dat = setColumnsToDisplayName(dat, list);
            return dat;
        }
        private DataTable setColumnsToDisplayName(DataTable dat, SPList list)
        {

            foreach (DataColumn dc in dat.Columns)
            {

                dc.ColumnName = list.Fields.GetFieldByInternalName(dc.ColumnName).Id.ToString();


            }
            foreach (DataColumn dc in dat.Columns)
            {

                string colName = list.Fields[new Guid(dc.ColumnName)].Title;
                if (dat.Columns.Contains(colName))
                {
                    int i = 1;
                    while (dat.Columns.Contains(colName + "_" + i.ToString()))
                    {
                        i++;
                    }
                    colName = colName + "_" + i.ToString();
                }
                dc.ColumnName = colName;


            }


            return dat;
        }
    }
}
