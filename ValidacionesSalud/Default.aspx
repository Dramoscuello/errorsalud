<%@ Page Title="Home Page" Language="VB" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Default.aspx.vb" Inherits="ValidacionesSalud._Default" %>


<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">

    <div class="jumbotron">
        <h1>Validador</h1>
    </div>
    <div class="row container" >
        <asp:Label ID="Label1" runat="server" Text="Archivo"></asp:Label>
        <asp:FileUpload ID="FileUpload1" runat="server" />
        <br />
        <asp:Button ID="Button1" runat="server" Text="Guardar" ToolTip="Guardar en la base de datos..." CssClass="btn btn-success" />
        <asp:Button ID="Button3" runat="server" Text="Validar" CssClass="btn btn-success" />
         <br />

        <asp:Label ID="Label2" runat="server" CssClass="lead" ForeColor="Red"></asp:Label>
      <br />

        <asp:GridView ID="GridView1"  runat="server" BackColor="White" BorderColor="#6699FF"  CssClass="table table-bordered bs-table table-responsive " 
                               
                AutoGenerateColumns="False" 
                        allowpaging="true" >
                <HeaderStyle BackColor="#337ab7" Font-Bold="True" ForeColor="White" />
        <EditRowStyle BackColor="#ffffcc" />
        <EmptyDataRowStyle forecolor="Red" CssClass="table table-bordered table-responsive" />
          
                <Columns >
            <asp:BoundField DataField="Nombre Archivo" HeaderText="Nombre Archivo" InsertVisible="False" ReadOnly="True" SortExpression="CustomerID" ControlStyle-Width="70px" />
            <asp:BoundField DataField="Numero de Registros" HeaderText="Numero de Registros" InsertVisible="False" ReadOnly="True" SortExpression="CustomerID" ControlStyle-Width="70px" />
            <asp:BoundField DataField="Registros Erroneos" HeaderText="Registros Erroneos" ReadOnly="True" SortExpression="CompanyName" ControlStyle-Width="300px" />
                        <asp:TemplateField ItemStyle-HorizontalAlign="Center" HeaderStyle-Width="200px">
        <ItemTemplate>
            <asp:Button ID="btnDescarga" runat="server" Text="Descargar Reporte" CssClass="btn btn-success" CommandName="Edit" />

        </ItemTemplate>
    </asp:TemplateField>
         
                </Columns>
            </asp:GridView>
    </div>

</asp:Content>
