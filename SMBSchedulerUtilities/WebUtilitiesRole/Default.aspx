<%@ Page Title="SMB Scheduler DL verification tool" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="WebUtilitiesRole._Default" %>

<asp:Content runat="server" ID="BodyContent" ContentPlaceHolderID="MainContent">
   
<script type="text/ecmascript">
        $(document).ready(function () {            
            var docHeight = $(document).height();
            var winHeight = $(window).height();
            var divBodyWid = $('#body').width();
            var contentWrapperHeight = $('section.content-wrapper').height();
            
            $(function () {
                
                if (docHeight > winHeight) {
                    
                    $('#Footer_Container').parent('footer').css('position', 'relative');
                }
                else {
                    $('#Footer_Container').parent('footer').css({ "position": "absolute", "bottom": "0" });
                }
                $('.container').css('height', docHeight);
                $('section.content-wrapper').css('height', winHeight - 180);
                $('footer').css('width', divBodyWid - 80);
                $('#MainContent_dvCustomerEntry').parent('.main-content').addClass('boxCenter');
            });
        });       


 </script>
 <div>
        <div style="float: left; width:220px;">
            <span style="font-weight:bold;">Choose DL : </span>
            <asp:DropDownList ID="ddlDLNames" runat="server" DataSourceID="xmlDLListDataSource" DataTextField="TeamName" DataValueField="DLName"></asp:DropDownList>
        </div>
        <div style="float: left; width:380px;">
            <span style="font-weight:bold;">Choose Base Timezone : </span>
            <asp:DropDownList ID="ddlTimeZones" runat="server"></asp:DropDownList>
        </div>
        <div style="float:left; padding-top: 15px;">           
            <asp:Button ID="btnSubmit" runat="server" Text=" Verify " OnClick="btnSubmit_Click" style="padding: 0 0 0 0 ;"/>
        </div>
        <asp:XmlDataSource ID="xmlDLListDataSource" runat="server" DataFile="~/App_Data/DLList.xml"></asp:XmlDataSource>

        <div>
            <asp:GridView ID="grdDLDetails" runat="server" AutoGenerateColumns="False" CellPadding="20">
                <Columns>
                    <asp:BoundField DataField="DLName" HeaderText="DL Name" ReadOnly="True"/>
                    <asp:BoundField DataField="EmailId" HeaderText="Agent Email ID" ReadOnly="True" />
                    <asp:BoundField DataField="AgentOffset" HeaderText="Agent Offset" ReadOnly="True" />
                    <asp:BoundField DataField="DLOffset" HeaderText="DL Offset" ReadOnly="True" />
                    <asp:BoundField DataField="IsBaseOffset" HeaderText="Is Offset Matched" ReadOnly="True" />
                    <asp:BoundField DataField="WorkhoursStartTime" HeaderText="Work hours Start Time" ReadOnly="True" />
                    <asp:BoundField DataField="WorkhoursEndTime" HeaderText="Work hours End Time" ReadOnly="True" />
                </Columns>
            </asp:GridView>

        </div>
        <div style="float: left;"><asp:Label runat="server" ID="lblError" Visible="false"></asp:Label></div>
    </div>
    
</asp:Content>
