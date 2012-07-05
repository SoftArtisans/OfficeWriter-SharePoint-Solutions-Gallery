using System;
using Microsoft.SharePoint;
using Microsoft.SharePoint.WebControls;
using System.Web.UI;
using System.Collections.Specialized;
using System.Web.UI.WebControls;
using OfficeWriter = SoftArtisans.OfficeWriter.WordWriter;
using System.IO;

namespace Item2Word.Layouts.SoftArtisans.Item2Word
{
    public partial class Item2Word : LayoutsPageBase
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
                if(validate())
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
                                OfficeWriter.WordApplication wApp = new OfficeWriter.WordApplication();
                                OfficeWriter.Document doc = wApp.Open(Page.MapPath("template/OfficeWriterStarterTemplate.doc"));
                                doc.DocumentProperties.Title = Title_TextBox.Text + " Template";

                                doc.GetBookmark("ActionName").InsertTextAfter(Title_TextBox.Text + " Template", false).Font.FontName = "Calibri";
                                doc.GetBookmark("ActionName").DeleteElement();

                                OfficeWriter.Table tab = (OfficeWriter.Table)doc.GetElements(OfficeWriter.Element.Type.Table)[0];

                                
                                foreach (ListItemCollection collection in new ListItemCollection[2] { new_template_col_selection_default_CheckBoxList.Items, new_template_col_selection_hidden_CheckBoxList.Items })
                                {

                                    foreach (ListItem item in collection)
                                    {
                                        if (item.Selected)
                                        {
                                            tab.AddRows(1);
                                            tab[tab.NumRows - 1, 0].InsertTextAfter(item.Text, true);
                                            tab[tab.NumRows - 1, 1].InsertMergeFieldAfter("\"[data].[" + item.Text + "]", "[data].[" + item.Text + "]\"");
                                        }

                                    }


                                }

                                tab[0, 0].Shading.BackgroundColor = System.Drawing.Color.LightGray;
                                tab[0, 1].Shading.BackgroundColor = System.Drawing.Color.LightGray;


                                Stream mStream = new MemoryStream();

                                wApp.Save(doc, mStream);


                                template = web.GetFolder(newTemplateLocation_AssetUrlSelector.AssetUrl).Files.Add(newTemplateFileName_TextBox.Text + ".doc", mStream);


                            }

                            action.Url = "~site/_layouts/SoftArtisans/Item2Word/Gen.aspx?List={ListId}&amp;Item={ItemId}&amp;TemplateLocation=" + (template.Web.Url + "/" + template.Url);
                            action.ImageUrl = "~site/_layouts/images/SoftArtisans/icons/wwsm.png";
                            action.Location = "EditControlBlock";
                            action.Update();
                            web.AllowUnsafeUpdates = false;

                            Message_Literal.Text = "<div class='ui-state-highlight'><h2>'" + action.Title + "' Action Sucessfully Created</h2>";
                            Message_Literal.Text += "<p>Your newly created action should now be available in the item drop down menu on the ";
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
                                        if ((new StringCollection() { ".doc", ".docx", ".docm", ".dotx", ".dotm" }).Contains(ext))
                                        {
                                            return true;
                                        }
                                        else
                                        {
                                            
                                            Template_AssetUrlSelector_Validator.Text = "A valid Word file is required";
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
