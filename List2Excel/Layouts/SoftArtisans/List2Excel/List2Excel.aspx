<%@ Assembly Name="$SharePoint.Project.AssemblyFullName$" %>
<%@ Import Namespace="Microsoft.SharePoint.ApplicationPages" %>
<%@ Register TagPrefix="SharePoint" Namespace="Microsoft.SharePoint.WebControls"
    Assembly="Microsoft.SharePoint, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register TagPrefix="Utilities" Namespace="Microsoft.SharePoint.Utilities" Assembly="Microsoft.SharePoint, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register TagPrefix="asp" Namespace="System.Web.UI" Assembly="System.Web.Extensions, Version=3.5.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35" %>
<%@ Import Namespace="Microsoft.SharePoint" %>
<%@ Assembly Name="Microsoft.Web.CommandUI, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>

<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="List2Excel.aspx.cs" Inherits="List2Excel.Layouts.SoftArtisans.List2Excel.List2Excel"
    DynamicMasterPageFile="~masterurl/default.master" %>

<%@ Register Assembly="Microsoft.SharePoint.Publishing, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c"
    Namespace="Microsoft.SharePoint.Publishing.WebControls" TagPrefix="cc1" %>
<asp:Content ID="PageHead" ContentPlaceHolderID="PlaceHolderAdditionalPageHead" runat="server">
    <style type="text/css">
        .inner_content a
        {
            color: #06B;
            text-decoration: none;
        }
        
        .inner_content a:hover
        {
            text-decoration: underline;
        }
        
        .inner_content
        {
            margin-left: auto;
            margin-right: auto;
            padding: 5px;
            width: 600px;
        }
        
        .settings_link
        {
            float: left;
        }
        .ow
        {
            float: right;
            font-size: 24px;
        }
        .ow .office
        {
            color: #06B;
        }
        
        .ow .writer
        {
            color: #888;
            font-style: italic;
        }
        
        .form_title
        {
            text-align: center;
            width: 100%;
            margin-top: 30px;
            margin-bottom: 10px;
            font-size: 20px;
            font-weight: bold;
        }
        .form_subtitle
        {
            text-align: center;
            width: 100%;
            margin-top: 5px;
            margin-bottom: 10px;
            font-size: 16px;
        }
        .about_link
        {
            text-align: center;
            margin-bottom: 25px;
        }
        
        
        .section
        {
            border: 1px solid #DDD;
            padding: 10px;
            background: #FAFAFA;
            margin-top: 20px;
            margin-bottom: 20px;
        }
        .section_header
        {
            font-size: 20px;
            color: #444;
            margin-bottom: 3px;
        }
        .section_subheader
        {
            margin-bottom: 20px;
            color: #444;
        }
        .ui-state-highlight
        {
            border: 1px solid #fcefa1;
            background: #fbf9ee 50% 50% repeat-x;
            color: #363636;
            -moz-border-radius: 5px;
            -moz-border-radius: 5px;
            -webkit-border-radius: 5px;
            -khtml-border-radius: 5px;
            border-bottom-right-radius: 5px;
            border-top-right-radius: 5px;
            border-bottom-left-radius: 5px;
            border-top-left-radius: 5px;
            width: 450px;
            text-align: center;
            margin-top: 10px;
            margin-right: auto;
            margin-bottom: 5px;
            margin-left: auto;
        }
        .ui-state-error
        {
            border: 1px solid #cd0a0a;
            background: #fef1ec 50% 50% repeat-x;
            color: #cd0a0a;
            -moz-border-radius: 5px;
            -webkit-border-radius: 5px;
            -khtml-border-radius: 5px;
            border-bottom-right-radius: 5px;
            border-top-right-radius: 5px;
            border-bottom-left-radius: 5px;
            border-top-left-radius: 5px;
            width: 450px;
            text-align: center;
            margin-top: 10px;
            margin-right: auto;
            margin-bottom: 5px;
            margin-left: auto;
        }
        .form_div
        {
            border: 2px solid #AAA;
            -moz-border-radius: 10px;
            -webkit-border-radius: 10px;
            -khtml-border-radius: 10px;
            border-radius: 10px;
            padding: 10px;
            padding-left: 20px;
            padding-right: 20px;
            margin-top: 10px;
            margin-right: auto;
            margin-bottom: 5px;
            margin-left: auto;
        }
        
        .form_control
        {
            display: block;
            margin-top: 2px;
            width: 98%;
        }
        .AssetTxtBox
        {
            width: 79%;
        }
        .labels
        {
            font-weight: bold;
            color: #888;
            margin-top: 12px;
            margin-bottom: 3px;
        }
        .help_text
        {
            font-style: italic;
            font-size: .9em;
        }
        
        .form_control input[type=button]
        {
            width: 19%;
            background-color: #DDD;
            color: #666;
            margin-top: 0px;
        }
        .inner_label
        {
            font-weight: bold;
            margin-bottom: 5px;
        }
        .show_hide
        {
            font-weight: normal;
            color: #666;
        }
        .show_hide a
        {
            color: #06B;
        }
        
        .inner_content input[type=text], .inner_content input[type=button]
        {
            border: 1px solid #AAA;
            margin: 0px;
            padding: 2px;
        }
        
        
        #new_template_col_selection
        {
            background-color: white;
            border: 1px solid #AAA;
            padding: 5px;
            margin-top: 5px;
        }
        .button_div
        {
            text-align: center;
            margin-bottom: 10px;
        }
        
        .button
        {
            color: white;
            font-weight: bold;
            background-color: #06B;
            text-align: center;
            padding: 4px 10px 4px 10px;
            font-size: 130%;
            border: 1px solid #AAA;
        }
        .validationError
        {
            color: #e05656 !important;
            margin-left: 10px;
            font-weight: normal;
        }
    </style>
    <script type="text/javascript">

        function toggle_visibility(id) {
            var e = document.getElementById(id);
            var a = document.getElementById(id + "_a")
            if (e.style.display == 'block') {
                e.style.display = 'none';
                a.innerHTML = "show";
            }
            else {
                e.style.display = 'block';
                a.innerHTML = "hide";
            }
        }

        

    </script>
</asp:Content>
<asp:Content ID="Main" ContentPlaceHolderID="PlaceHolderMain" runat="server">
    <div class="inner_content">
        <asp:Literal ID="Message_Literal" runat="server"></asp:Literal>
        <div class="form_div">
            <div class="header">
                <a class="settings_link" href="..\..\settings.aspx">Back to Site Settings</a> <span
                    class="ow"><span class="office">Office</span><span class="writer">Writer</span></span>
            </div>
            <div class="form_title">
                Excel Export Plus</div>
            <div class="form_subtitle">
                Create a custom action that takes an item from a SharePoint list and exports it
                into a pre-formatted Excel template.</div>
            <div class="about_link">
                <a href="http://www.officewriter.com/sharepoint/solutions/ExcelExportPlus/docs">More
                    About This Solution</a></div>
            <div class="section">
                <div class="section_header">
                    SharePoint List</div>
                <div class="section_subheader">
                    Where do you want to get your data from?</div>
                <div class="labels">
                    <span class="validationError">
                        <asp:Literal ID="List_AssetUrlSelector_Validator" runat="server" /></span></div>
                <cc1:AssetUrlSelector ID="List_AssetUrlSelector" runat="server" CssClass="form_control"
                    CssTextBox="AssetTxtBox" OverrideDialogTitle="Select a List" AutoPostBack="True" />
                <span class="help_text">You can choose any list in the site</span>
            </div>
            <div class="section">
                <div class="section_header">
                    Excel Template</div>
                <div class="section_subheader">
                    Where do you want to export your data?</div>
                <asp:RadioButtonList ID="templateSelection_RadioButtonList" runat="server" class="template_radiobuttons"
                    AutoPostBack="True" Enabled="False">
                    <asp:ListItem Selected="True" Text="Use an existing template file" Value="existingTemplate" />
                    <asp:ListItem Text="Create a new template file" Value="newTemplate" />
                </asp:RadioButtonList>
                <asp:MultiView ID="templateSelection_MultiView" runat="server" ActiveViewIndex="0">
                    <asp:View ID="existingTemplate_View" runat="server">
                        <div id="existing_template">
                            <div class="labels">
                                Template File <span class="validationError">
                                    <asp:Literal ID="Template_AssetUrlSelector_Validator" runat="server" /></span>
                            </div>
                            <cc1:AssetUrlSelector ID="Template_AssetUrlSelector" runat="server" CssClass="form_control"
                                CssTextBox="AssetTxtBox" OverrideDialogTitle="Select a Template" />
                            <div class="help_text">
                                .XLS or .XLSX template file
                            </div>
                        </div>
                    </asp:View>
                    <asp:View ID="newTemplate_View" runat="server">
                        <div id="new_template">
                            <div class="labels">
                                Columns to Include<span class="validationError"><asp:Literal ID="col_selection_CheckBoxList_Validator"
                                    runat="server" /></span>
                            </div>
                            <div id="new_template_col_selection">
                                <div class="inner_label">
                                    Default Columns<span class="show_hide"> (<a id="new_template_col_selection_default_a"
                                        href="#" onclick="toggle_visibility('new_template_col_selection_default');">hide</a>)</span></div>
                                <div id="new_template_col_selection_default" style="display: block">
                                    <asp:CheckBoxList ID="new_template_col_selection_default_CheckBoxList" runat="server">
                                    </asp:CheckBoxList>
                                </div>
                                <div class="inner_label">
                                    Other Columns<span class="show_hide"> (<a id="new_template_col_selection_hidden_a"
                                        href="#" onclick="toggle_visibility('new_template_col_selection_hidden');">show</a>)</span></div>
                                <div id="new_template_col_selection_hidden" style="display: none;">
                                    <asp:CheckBoxList ID="new_template_col_selection_hidden_CheckBoxList" runat="server">
                                    </asp:CheckBoxList>
                                </div>
                            </div>
                            <div class="labels">
                                Template File Location<span class="validationError"><asp:Literal ID="newTemplateLocation_AssetUrlSelector_Validator"
                                    runat="server" /></span>
                            </div>
                            <cc1:AssetUrlSelector ID="newTemplateLocation_AssetUrlSelector" runat="server" CssClass="form_control"
                                CssTextBox="AssetTxtBox" AutoPostBack="True" />
                            <div class="labels">
                                Template File Name<span class="validationError"><asp:Literal ID="newTemplateFileName_Validator"
                                    runat="server" /></span>
                            </div>
                            <asp:TextBox ID="newTemplateFileName_TextBox" runat="server" TextMode="SingleLine"
                                CssClass="form_control"></asp:TextBox>
                        </div>
                    </asp:View>
                </asp:MultiView>
            </div>
            <div class="section">
                <div class="section_header">
                    Action Name</div>
                <div class="section_subheader">
                    What do you to call your new custom action?</div>
                <div class="labels">
                    Name<span class="validationError"><asp:Literal runat="server" ID="Title_TextBox_Validator" /></span></div>
                <asp:TextBox ID="Title_TextBox" runat="server" TextMode="SingleLine" CssClass="form_control"></asp:TextBox>
                <div class="help_text">
                    This name that will appear in the context menu for an item</div>
                <div class="labels">
                    Description</div>
                <asp:TextBox ID="Discription_TextBox" runat="server" TextMode="SingleLine" CssClass="form_control"></asp:TextBox>
                <div class="help_text">
                    (Optional)</div>
            </div>
            <div class="button_div">
                <asp:Button ID="Button1" runat="server" OnClientClick="scroll(0,0);" OnClick="createButtonAction"
                    Text="Create New Action" CssClass="button" />
            </div>
        </div>
    </div>
</asp:Content>
<asp:Content ID="PageTitle" ContentPlaceHolderID="PlaceHolderPageTitle" runat="server">
    Excel Export Plus
</asp:Content>
<asp:Content ID="PageTitleInTitleArea" ContentPlaceHolderID="PlaceHolderPageTitleInTitleArea"
    runat="server">
    Excel Export Plus
</asp:Content>
