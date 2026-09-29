<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="HeaderControl.ascx.vb" Inherits="Atom.Web.HeaderControl" %>
<div id="divHeader" runat="server" class="atom-header">
    <asp:Label ID="lblTitle" runat="server" CssClass="atom-header-title" />
    <div class="atom-header-right">
        <asp:Label ID="lblName" runat="server" CssClass="atom-header-name" />
        <asp:PlaceHolder ID="phButtons" runat="server" />
    </div>
</div>
