  <%@ Page Title="" Language="C#" MasterPageFile="~/main.master" AutoEventWireup="true" CodeFile="ExcelMapping5_final21.aspx.cs" Inherits="ExcelMapping5_final21" %>

<asp:Content ID="Content1" ContentPlaceHolderID="title" runat="Server">
   <style>
        .cb label {
            margin-left: 5px;
        }
    </style>   
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="head" runat="Server">
</asp:Content>
<asp:Content ID="Content3" ContentPlaceHolderID="pagecontents" runat="Server">
    <section class="content-header">
        
        <%--<ol class="breadcrumb">
        <li><a href="#"><i class="fa fa-dashboard"></i> Home</a></li>
        <li><a href="#">CRM</a></li>
        <li class="active">Create</li>
      </ol>
        --%> 
    </section>

    <section class="content">
        <div class="row">
            <!-- left column -->
            <div class="col-md-12">
                <!-- /.box-header -->
                <!-- TABLE: Search Criteria -->
                <div class="box box-info">
                    <div class="box-header with-border">
                        <h2 style="color: #5D7B9D; font-weight: bold; margin-left: 10px;">Excel Data Mapping</h2>
                        

                    </div>
                   
                    <div class="box-body">
                          <div style="backdrop-filter: blur(10px); background-color: rgba(255, 255, 255, 0.2); border: 1px solid rgba(255, 255, 255, 0.3); padding: 20px; margin: 20px; border-radius: 15px; box-shadow: 0 8px 32px rgba(0, 0, 0, 0.1);">
                        <div class="row">
                          
                            <div class="col-md-3">
                                <h3 style="color: #5D7B9D;font-family:'Arial Rounded MT'">Select Company</h3>
                                <asp:DropDownList ID="Ddlcompany" CssClass="form-control" runat="server"></asp:DropDownList>

                            </div>
                            <div class="col-md-3">
                                <h3 style="color: #5D7B9D;font-family:'Arial Rounded MT'">Select Act</h3>
                                <asp:DropDownList ID="DdlAct" CssClass="form-control" runat="server"></asp:DropDownList>

                            </div>

                            <div class="col-md-6">
                                <h3 style="color: #5D7B9D;font-family:'Arial Rounded MT'">Upload File</h3>
                                <asp:FileUpload ID="FileUpload1" CssClass="form-control" runat="server"></asp:FileUpload>

                            </div>  
                           

                        </div>
                        <div class="row">
                            <div class="col-md-6 col-xs-6">
                                <asp:Button runat="server" ID="Button1" UseSubmitBehavior="false" OnClientClick="this.disabled=true;" CssClass="btn bg-green margin" Text="Upload File" ValidationGroup="vgregister" OnClick="btnUpload_Click" />
                            </div>
                        </div>
                              </div> 
                        <div class="row">
                            <div class="col-md-12">
                                <div class="dashboard_graph">
                                    <div class="row">
                                        <div class="col-xs-12">

                                         
                                      <%--      <asp:GridView ID="grdview" class="table no-margin" runat="server" AutoGenerateColumns="false"  OnRowCommand="grdview_RowCommand">
                                               
                                                <Columns>
                                                    <asp:TemplateField HeaderText="ID" Visible="false">
                                                        <ItemTemplate>

                                                            <asp:Label ID="lblpk" Text='<%# Eval("pk") %>' runat="server" />

                                                        </ItemTemplate>
                                                    </asp:TemplateField>
                                                    <asp:TemplateField HeaderText="Excel Column">
                                                        <ItemTemplate>

                                                            <asp:Label Text='<%# Eval("ExcelColumn") %>' runat="server" />

                                                        </ItemTemplate>
                                                    </asp:TemplateField>
                                                    <asp:TemplateField HeaderText="DB Column">
                                                        <ItemTemplate>

                                                            <asp:Label Text='<%# Eval("DBColumn") %>' runat="server" />

                                                        </ItemTemplate>

                                                    </asp:TemplateField>
                                                    <asp:TemplateField HeaderText="Delete">
                                                        <ItemTemplate>

                                                            <asp:Button Text="Delete" runat="server" CommandName="DeleteRow" CommandArgument='<%# Eval("pk")%>' />

                                                        </ItemTemplate>


                                                    </asp:TemplateField>
                                                  

                                                </Columns>
                                            </asp:GridView>--%>
                                                 
                                        </div>
                                    </div>
                                   
                                  <div style="backdrop-filter: blur(10px); background-color: rgba(255, 255, 255, 0.2); border: 1px solid rgba(255, 255, 255, 0.3); padding: 20px; margin: 20px; border-radius: 15px; box-shadow: 0 8px 32px rgba(0, 0, 0, 0.1);">
                                    <%--<div style="border: 1px solid #ddd; padding: 20px; margin: 20px; border-radius: 6px; background-color: #ffffff;">--%>
                                   <%-- <div style="border: none; padding: 25px; margin: 20px; border-radius: 15px; background: linear-gradient(135deg, #e0f7fa, #ffffff); box-shadow: 0 6px 20px rgba(0, 0, 0, 0.1);">--%>
                                   <%-- <div style="border: 3px solid #7ab0d4; padding: 20px; margin: 20px; border-radius: 10px; box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1); background-color: #f4f6f8;">--%>
                                         <h3 style="color: #5D7B9D; font-weight: bold; margin-left: 10px;">Mapped Fields Summary</h3>
                                      <asp:ListView ID="grdview" runat="server" OnItemCommand="grdview_ItemCommand"   >
    <LayoutTemplate>
        <table class="table no-margin table-bordered table-hover " id="tblkb" >
          <thead style="background: #5D7B9D; color: white;">
                <tr>
                     <th style="width: 1%; text-align: center; white-space: nowrap;">
                    <th  style= "display:none;" >ID</th>
                    <th>Excel Column</th>
                    <th>DB Column</th>
                    <th style="display:none">Delete</th>
                    
                </tr>
            </thead>
            <tr id="itemPlaceholder" runat="server"></tr>
        </table>
    </LayoutTemplate>
    <ItemTemplate>
        <tr>
            <td style="width: 1%; text-align: center; white-space: nowrap;">
                <asp:CheckBox ID="chkSelect" runat="server" />
            </td>
            <td style= "display:none;">
                <asp:Label ID="lblexcelcolumns"  runat="server" Text='<%# Eval("pk") %>'></asp:Label></td>

            <td>
                <asp:Label ID="Label1" runat="server" Text='<%# Eval("ExcelColumn") %>'></asp:Label></td>
            <td>  <asp:Label ID="Label2" runat="server" Text='<%# Eval("DBColumn") %>'></asp:Label></td>
              </td>
           

              <td style="display:none">

                 <asp:Button Text="Delete" CssClass="btn-danger" runat="server" CommandName="DeleteRow" CommandArgument='<%# Eval("pk")%>' />
        </tr>
    </ItemTemplate>
    

    <EmptyDataTemplate>
        <h2>No Records!</h2>
    </EmptyDataTemplate>
</asp:ListView>
<asp:Button ID ="btnDeleteMAPPING" Visible="false"  CssClass="btn-danger" Text="Delete" runat="server" CommandName="DeleteRow" OnClick="btnDeleteMAPPING_Click" />
                                        </div>








                                       <div style="backdrop-filter: blur(10px); background-color: rgba(255, 255, 255, 0.2); border: 1px solid rgba(255, 255, 255, 0.3); padding: 20px; margin: 20px; border-radius: 15px; box-shadow: 0 8px 32px rgba(0, 0, 0, 0.1);">

                                    <h3 style="color: #5D7B9D; font-weight: bold; margin-left: 10px;">Excel Mapping</h3>

                                    <asp:ListView ID="lstdc" runat="server" OnItemDataBound="lstdc_ItemDataBound"  >
                                        <LayoutTemplate>
                                            <table class="table no-margin" >
                                                <thead>
                                                    <tr>
                                                        <th>Excel Column</th>
                                                        <th>DB Column</th>
                                                        <th>Sequence No</th>
                                                        <th>Check</th>
                                                        <th style="display:none">Update</th>
                                                    </tr>
                                                </thead>
                                                <tr id="itemPlaceholder" runat="server"></tr>
                                            </table>
                                        </LayoutTemplate>
                                        <ItemTemplate>
                                            <tr>
                                                <td>
                                                    <asp:Label ID="lblexcelcolumns" runat="server" Text='<%#Eval("Name")%>'></asp:Label></td>

                                                <td>
                                                    <asp:DropDownList ID="ddlDbColumn" CssClass="form-control" runat="server"></asp:DropDownList></td>
                                                <td> <asp:Textbox  ID="txtSequenceNo" TextMode="Number"   runat="server"></asp:Textbox>
                                                    </td>
                                                <td>
                                                      <asp:CheckBox  ID="chkMapping"  runat="server"></asp:CheckBox>
                                                </td>

                                                  <td style="display:none">

                                                    <asp:Button Text="Update" ID ="btnUpdate" runat="server" OnClick="btnUpdate_Click" /></td>
                                            </tr>
                                        </ItemTemplate>
                                        

                                        <EmptyDataTemplate>
                                            <h2>No Records!</h2>
                                        </EmptyDataTemplate>
                                    </asp:ListView>
                                </div>


                                <div class="clearfix"></div>

                            </div>
                        </div>


                        
            

               
                 <div class="row">
                     
                     <div class="col-md-6 col-xs-6">
                         <asp:Button runat="server" ID="btnSubmit"   UseSubmitBehavior="false" OnClientClick="this.disabled=true;" CssClass="btn bg-green margin" Text="Submit" ValidationGroup="vgregister" OnClick="btnSubmit_Click" />
                     </div>
                 </div>
                    </div>
                </div>
                         
            </div>
             </div>        
  
     <!-- top tiles -->
</section>
</asp:Content>

