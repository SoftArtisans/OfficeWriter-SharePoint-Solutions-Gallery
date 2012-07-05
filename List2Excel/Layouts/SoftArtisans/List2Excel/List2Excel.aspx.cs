using System;
using Microsoft.SharePoint;
using Microsoft.SharePoint.WebControls;
using System.Web.UI;
using SoftArtisans.OfficeWriter.ExcelWriter;
using System.Web.UI.WebControls;
using System.Collections.Specialized;
using System.IO;

namespace List2Excel.Layouts.SoftArtisans.List2Excel
{
    public partial class List2Excel : LayoutsPageBase
    {
        protected void Page_Load(object sender, EventArgs e)
        {





            if (Page.IsPostBack)
            {

                Control c = GetPostBackControl(this.Page);
                if (c == List_AssetUrlSelector)
                {

                    try
                    {
                        using (SPSite siteCollection = new SPSite(SPContext.Current.Web.Url))
                        {
                            using (SPWeb web = siteCollection.OpenWeb())
                            {

                                SPList list = web.GetList(List_AssetUrlSelector.AssetUrl);
                                StringCollection defaultViewFields = list.DefaultView.ViewFields.ToStringCollection();
                                new_template_col_selection_default_CheckBoxList.Items.Clear();
                                new_template_col_selection_hidden_CheckBoxList.Items.Clear();
                                ListItem item;
                                foreach (SPField field in list.Fields)
                                {
                                    item = new ListItem(field.Title, field.StaticName);

                                    if (list.DefaultView.ViewFields.Exists(field.StaticName))
                                    {
                                        new_template_col_selection_default_CheckBoxList.Items.Add(item);
                                    }
                                    else
                                    {
                                        new_template_col_selection_hidden_CheckBoxList.Items.Add(item);
                                    }
                                }
                                List_AssetUrlSelector.AssetUrl = list.ParentWeb.Url + "/" + list.RootFolder.Url;
                                templateSelection_RadioButtonList.Enabled = true;
                            }


                        }
                    }
                    catch (Exception ex)
                    {
                        Message_Literal.Text = "<div  class='ui-state-error'><h2>Error</h2>";
                        Message_Literal.Text += "<p>" + ex.Message + "</p></div>";
                        Message_Literal.Visible = true;
                        List_AssetUrlSelector.Focus();
                    }

                }
                else if (c == templateSelection_RadioButtonList)
                {
                    if (templateSelection_RadioButtonList.SelectedValue == "existingTemplate")
                    {
                        templateSelection_MultiView.SetActiveView(existingTemplate_View);
                    }
                    else
                    {
                        templateSelection_MultiView.SetActiveView(newTemplate_View);
                    }
                }
                else if (c == newTemplateLocation_AssetUrlSelector)
                {

                    try
                    {
                        using (SPSite siteCollection = new SPSite(SPContext.Current.Web.Url))
                        {
                            using (SPWeb web = siteCollection.OpenWeb())
                            {
                                SPFolder folder = web.GetFolder(newTemplateLocation_AssetUrlSelector.AssetUrl);
                                if (folder.DocumentLibrary == null)
                                {
                                    newTemplateLocation_AssetUrlSelector.AssetUrl = "";
                                }
                                else
                                {
                                    while (!folder.Exists)
                                    {
                                        folder = folder.ParentFolder;
                                    }
                                    newTemplateLocation_AssetUrlSelector.AssetUrl = folder.ParentWeb.Url + "/" + folder.Url;
                                }


                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Message_Literal.Text = "<div  class='ui-state-error'><h2>Error</h2>";
                        Message_Literal.Text += "<p>" + ex.Message + "</p></div>";
                        Message_Literal.Visible = true;
                        List_AssetUrlSelector.Focus();
                    }

                }


            }







        }

        protected void createButtonAction(object sender, EventArgs e)
        {
            try
            {
                if (validate())
                {
                    using (SPSite siteCollection = new SPSite(SPContext.Current.Web.Url))
                    {
                        using (SPWeb web = siteCollection.OpenWeb())
                        {

                            SPList list = web.GetList(List_AssetUrlSelector.AssetUrl);
                            SPFile template;
                            SPUserCustomAction action = list.UserCustomActions.Add();
                            action.Title = Title_TextBox.Text;
                            action.Description = Discription_TextBox.Text;




                            if (templateSelection_RadioButtonList.SelectedValue == "existingTemplate")
                            {

                                template = web.GetFile(Template_AssetUrlSelector.AssetUrl);

                            }
                            else
                            {
                                ExcelApplication xla = new ExcelApplication();
                                Workbook wb = xla.Open(Page.MapPath("template/OfficeWriterStarterTemplate.xls"));
                                wb.DocumentProperties.Title = Title_TextBox.Text + " Template";

                                wb.GetNamedRange("Title").Areas[0][0, 0].Value = Title_TextBox.Text + " Template";


                                int startRow = wb.GetNamedRange("dataMarkerStart").Areas[0].FirstRow;
                                int startCol = wb.GetNamedRange("dataMarkerStart").Areas[0].FirstColumn;
                                Worksheet ws = wb[wb.GetNamedRange("dataMarkerStart").Areas[0].WorksheetIndex];
                                int col = 0;

                                foreach (ListItemCollection collection in new ListItemCollection[2] { new_template_col_selection_default_CheckBoxList.Items, new_template_col_selection_hidden_CheckBoxList.Items })
                                {

                                    foreach (ListItem item in collection)
                                    {
                                        if (item.Selected)
                                        {
                                            ws.Cells[startRow, startCol + col].Value = item.Text;
                                            ws.Cells[startRow + 1, startCol + col].Value = "%%=[data].[" + item.Text + "]";
                                            

                                            col++;
                                        }

                                    }


                                }


                                Stream mStream = new MemoryStream();

                                xla.Save(wb, mStream);


                                template = web.GetFolder(newTemplateLocation_AssetUrlSelector.AssetUrl).Files.Add(newTemplateFileName_TextBox.Text + ".xls", mStream);


                            }

                            action.Url = "~site/_layouts/SoftArtisans/List2Excel/Gen.aspx?List={ListId}&amp;TemplateLocation=" + (template.Web.Url + "/" + template.Url); 
                            action.ImageUrl = "~site/_layouts/images/SoftArtisans/icons/excelwriter.png";
                            action.Location = "CommandUI.Ribbon.ListView";

                            string extXml = " <CommandUIExtension xmlns='http://schemas.microsoft.com/sharepoint/'><CommandUIDefinitions><CommandUIDefinition Location='Ribbon.ListItem.Actions.Controls._children'>";
                            extXml += "<Button Id='" + action.Id.ToString() + "' ";
                            extXml += "Command='" + action.Id.ToString() + "action" + @"' ";
                            extXml += "Image32by32='~site/_layouts/images/SoftArtisans/icons/excelwriter.png' Image16by16='~site/_layouts/images/SoftArtisans/icons/xlsm.png' Sequence='0' ";
                            extXml += "LabelText='" + action.Title + "' ";
                            extXml += "Description='" + action.Description + "' ";
                            extXml += "TemplateAlias='o1' /></CommandUIDefinition></CommandUIDefinitions><CommandUIHandlers>";
                            extXml += "<CommandUIHandler Command='" + action.Id.ToString() + "action" + @"' ";
                            extXml += "CommandAction='" + action.Url + "' />";
                            extXml += "</CommandUIHandlers></CommandUIExtension>";
                            action.CommandUIExtension = extXml;
                            action.Update();
                            web.AllowUnsafeUpdates = false;
                            Message_Literal.Text = "<div class='ui-state-highlight'><h2>'" + action.Title + "' Button Sucessfully Created</h2>";
                            Message_Literal.Text += "<p>Your newly created button should now be available in the Ribbon on the ";
                            Message_Literal.Text += "<a href='" + list.DefaultViewUrl + "'>" + list.Title + "</a> list</p></div>";
                            Message_Literal.Visible = true;

                            resetForm();
                        }

                    }
                }
            }
            catch (Exception ex)
            {
                Message_Literal.Text = "<div  class='ui-state-error'><h2>Error</h2>";
                Message_Literal.Text += "<p>" + ex.Message + "</p></div>";
                Message_Literal.Visible = true;
                List_AssetUrlSelector.Focus();
            }


        }
        public static Control GetPostBackControl(Page page)
        {
            Control control = null;

            string ctrlname = page.Request.Params.Get("__EVENTTARGET");
            if (ctrlname != null && ctrlname != string.Empty)
            {
                control = page.FindControl(ctrlname);
            }
            else
            {
                foreach (string ctl in page.Request.Form)
                {
                    Control c = page.FindControl(ctl);
                    if (c is System.Web.UI.WebControls.Button)
                    {
                        control = c;
                        break;
                    }
                }
            }
            return control;
        }

        private void resetForm()
        {
            //list info
            List_AssetUrlSelector.AssetUrl = "";

            //action info
            Title_TextBox.Text = "";
            Discription_TextBox.Text = "";

            //select template type
            templateSelection_RadioButtonList.SelectedIndex = 0;
            templateSelection_RadioButtonList.Enabled = false;
            templateSelection_MultiView.SetActiveView(existingTemplate_View);

            //existing template
            Template_AssetUrlSelector.AssetUrl = "";

            //new template
            newTemplateFileName_TextBox.Text = "";
            newTemplateLocation_AssetUrlSelector.AssetUrl = "";
            new_template_col_selection_default_CheckBoxList.Items.Clear();
            new_template_col_selection_hidden_CheckBoxList.Items.Clear();


        }

        #region Validation

        bool validate()
        {
            col_selection_CheckBoxList_Validator.Text = "";
            List_AssetUrlSelector_Validator.Text = "";
            newTemplateFileName_Validator.Text = "";
            newTemplateLocation_AssetUrlSelector_Validator.Text = "";
            Template_AssetUrlSelector_Validator.Text = "";
            Title_TextBox_Validator.Text = "";



            if (List_AssetUrlSelector.AssetUrl == "")
            {
                List_AssetUrlSelector_Validator.Text = "A list is required";
                return false;
            }
            else if (Title_TextBox.Text == "")
            {
                Title_TextBox_Validator.Text = "An action name is required";
                return false;
            }
            else
            {
                return (RequireExistingTemplateFile() && RequireNewTemplateName() && RequireNewTemplateLocation() && RequireNewTemplateCols());
            }


        }



        private bool RequireExistingTemplateFile()
        {

            if (templateSelection_RadioButtonList.SelectedValue == "existingTemplate")
            {
                if (Template_AssetUrlSelector.AssetUrl == "")
                {

                    Template_AssetUrlSelector_Validator.Text = "A template is required";
                    return false;
                }
                else
                {
                    try
                    {
                        using (SPSite siteCollection = new SPSite(SPContext.Current.Web.Url))
                        {
                            using (SPWeb web = siteCollection.OpenWeb())
                            {
                                SPFile template = web.GetFile(Template_AssetUrlSelector.AssetUrl);
                                if (template.Exists)
                                {
                                    string ext = new FileInfo(template.Name).Extension;
                                    if ((new StringCollection() { ".xlsx", ".xlsm", ".xls", ".xlst"}).Contains(ext))
                                    {
                                        return true;
                                    }
                                    else
                                    {

                                        Template_AssetUrlSelector_Validator.Text = "A valid Excel file is required";
                                        return false;

                                    }
                                }
                                else
                                {

                                    Template_AssetUrlSelector_Validator.Text = "A template is required";
                                    return false;

                                }
                            }
                        }

                    }
                    catch
                    {

                        Template_AssetUrlSelector_Validator.Text = "A template is required";
                        return false;

                    }
                }
            }
            else
            {
                return true;
            }


        }
        private bool RequireNewTemplateName()
        {
            if (templateSelection_RadioButtonList.SelectedValue == "newTemplate")
            {
                if (newTemplateFileName_TextBox.Text == "")
                {
                    newTemplateFileName_Validator.Text = "A template name is required";
                    return false;
                }
                else
                {
                    return true;
                }

            }
            else
            {
                return true;
            }



        }

        private bool RequireNewTemplateLocation()
        {
            if (templateSelection_RadioButtonList.SelectedValue == "newTemplate")
            {

                if (newTemplateLocation_AssetUrlSelector.AssetUrl == "")
                {
                    newTemplateLocation_AssetUrlSelector_Validator.Text = "A location is required";
                    return false;
                }
                else
                {
                    return true;
                }

            }
            else
            {
                return true;
            }
        }

        private bool RequireNewTemplateCols()
        {
            if (templateSelection_RadioButtonList.SelectedValue == "newTemplate")
            {
                if (new_template_col_selection_default_CheckBoxList.SelectedIndex + new_template_col_selection_hidden_CheckBoxList.SelectedIndex == -2)
                {
                    col_selection_CheckBoxList_Validator.Text = "At least one column is required";
                    return false;
                }
                else
                {
                    return true;
                }

            }
            else
            {
                return true;
            }
        }



        #endregion

    }
}
