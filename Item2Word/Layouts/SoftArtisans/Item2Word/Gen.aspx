<%@ Assembly Name="$SharePoint.Project.AssemblyFullName$" %>
<%@ Import Namespace="Microsoft.SharePoint.ApplicationPages" %>
<%@ Register Tagprefix="SharePoint" Namespace="Microsoft.SharePoint.WebControls" Assembly="Microsoft.SharePoint, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register Tagprefix="Utilities" Namespace="Microsoft.SharePoint.Utilities" Assembly="Microsoft.SharePoint, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register Tagprefix="asp" Namespace="System.Web.UI" Assembly="System.Web.Extensions, Version=3.5.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35" %>
<%@ Import Namespace="Microsoft.SharePoint" %>
<%@ Assembly Name="Microsoft.Web.CommandUI, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Gen.aspx.cs" Inherits="Item2Word.Layouts.SoftArtisans.Item2Word.Gen" DynamicMasterPageFile="~masterurl/default.master" %>

<asp:Content ID="PageHead" ContentPlaceHolderID="PlaceHolderAdditionalPageHead" runat="server">
       <style type="text/css">
        .ui-state-error
        {
            border: 1px solid #cd0a0a;
            background: #fef1ec 50% 50% repeat-x;
            color: #cd0a0a;
            -moz-border-radius: 4px;
            -webkit-border-radius: 4px;
            -khtml-border-radius: 4px;
            border-bottom-right-radius: 4px;
            border-top-right-radius: 4px;
            border-bottom-left-radius: 4px;
            border-top-left-radius: 4px;
            width: 400px;
            text-align: center;
            margin-top: 10px;
            margin-right: auto;
            margin-bottom: 4px;
            margin-left: auto;
        }
        </style>
</asp:Content>

<asp:Content ID="Main" ContentPlaceHolderID="PlaceHolderMain" runat="server">

<div  class='ui-state-error'><h2>Error</h2><p><asp:Label ID="ErrorText" runat="server" Text="Label"></asp:Label></p></div>

</asp:Content>

<asp:Content ID="PageTitle" ContentPlaceHolderID="PlaceHolderPageTitle" runat="server">
Application Page
</asp:Content>

<asp:Content ID="PageTitleInTitleArea" ContentPlaceHolderID="PlaceHolderPageTitleInTitleArea" runat="server" >
My Application Page
</asp:Content>
