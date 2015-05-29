<%@ Page Language="C#" %>

<%@ Import Namespace="System.Globalization" %>
<%@ Import Namespace="DevExpress.XtraPivotGrid" %>

<%@ Register Assembly="DevExpress.Web.ASPxPivotGrid.v13.1, Version=13.1.5.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxPivotGrid" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.ASPxPivotGrid.v13.1, Version=13.1.5.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxPivotGrid.Export" TagPrefix="dx" %>
<script runat="server">
   
    protected void Page_Load(object sender, EventArgs e)
    {
        //foreach (var v in ASPxPivotGrid1.Fields)
        //{
        //    ((PivotGridFieldBase)v).CollapseAll();
        //}
        if (!Page.IsPostBack)
        {
            DateTime max = new DateTime();
            foreach (var data in ASPxPivotGrid1.Fields[7].FilterValues.ValuesIncluded)
            {
                var d = DateTime.ParseExact(data.ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
                if (d > max)
                    max = d;
            }
            ASPxPivotGrid1.Fields[7].FilterValues.ValuesIncluded = new object[] { max.ToString("dd/MM/yyyy") };

        }

    }
    private void pivotGrid_FieldValueDisplayText(object sender, PivotFieldDisplayTextEventArgs e)
    {
        if (e.ValueType == PivotGridValueType.Total)
            e.DisplayText = "Average value";
        if (e.ValueType == PivotGridValueType.GrandTotal)
                e.DisplayText = " Total Average";

    }

    protected void Button1_Click(object sender, EventArgs e)
    {
        ASPxPivotGridExporter1.ExportXlsxToResponse("export");
    }
</script>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">



<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Title</title>
</head>
<body>
    <form id="HtmlForm" runat="server">
        <div>
            <dx:ASPxPivotGrid ID="ASPxPivotGrid1" runat="server" ClientIDMode="AutoID" DataSourceID="SqlDataSource1" OnFieldValueDisplayText="pivotGrid_FieldValueDisplayText">
                <OptionsView DataHeadersDisplayMode="Popup" ShowColumnGrandTotals="False" ShowColumnTotals="False"></OptionsView>
                <Fields>
                    <dx:PivotGridField ID="fieldGiocatore" Area="RowArea" AreaIndex="0" FieldName="Giocatore">
                    </dx:PivotGridField>
                    <dx:PivotGridField ID="fieldTipoAllenamento" AreaIndex="3" FieldName="Tipo Allenamento" Area="RowArea">
                    </dx:PivotGridField>

                    <dx:PivotGridField ID="fieldMese" AreaIndex="1" FieldName="Mese">
                    </dx:PivotGridField>
                    <dx:PivotGridField ID="fieldN" AreaIndex="0" FieldName="N" Caption="#">
                    </dx:PivotGridField>
                    <dx:PivotGridField ID="fieldValue" Area="DataArea" AreaIndex="0"
                        CellFormat-FormatString="N2" CellFormat-FormatType="Numeric" FieldName="Value" GrandTotalCellFormat-FormatString="N2" GrandTotalCellFormat-FormatType="Numeric" SummaryType="Average" TotalCellFormat-FormatString="N2" TotalCellFormat-FormatType="Numeric" Caption="Value Avg">
                    </dx:PivotGridField>
                    <dx:PivotGridField ID="fieldField" Area="ColumnArea" AreaIndex="1" FieldName="Field">
                    </dx:PivotGridField>
                    <dx:PivotGridField ID="fieldGruppo" Area="ColumnArea" AreaIndex="0" FieldName="Gruppo">
                    </dx:PivotGridField>
                    <dx:PivotGridField ID="fieldData" Area="RowArea" AreaIndex="2" FieldName="Data" CellFormat-FormatString="d" CellFormat-FormatType="DateTime">
                    </dx:PivotGridField>
                    <dx:PivotGridField ID="fieldValue1" AreaIndex="2" Caption="Value Sum" CellFormat-FormatString="N2" CellFormat-FormatType="Numeric" FieldName="Value" GrandTotalCellFormat-FormatString="N2" GrandTotalCellFormat-FormatType="Numeric" SummaryType="Sum" TotalCellFormat-FormatString="N2" TotalCellFormat-FormatType="Numeric">
                    </dx:PivotGridField>
                    <dx:PivotGridField ID="fieldRuolo" Area="RowArea" AreaIndex="1" FieldName="Ruolo">
                    </dx:PivotGridField>
                </Fields>
                <OptionsPager RowsPerPage="0">
                </OptionsPager>
                <Groups>
                    <dx:PivotGridWebGroup Caption="Absolute values" ShowNewValues="True" />
                    <dx:PivotGridWebGroup Caption="Relative values" ShowNewValues="True" />
                    <dx:PivotGridWebGroup Caption=" Energy expenditure " ShowNewValues="True" />
                </Groups>
            </dx:ASPxPivotGrid>
            <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>" ProviderName="<%$ ConnectionStrings:ConnectionString.ProviderName %>" SelectCommand="SELECT * FROM [vw_dataPivot]"></asp:SqlDataSource>
            <br />
            <asp:Button ID="Button1" runat="server" OnClick="Button1_Click" Text="Export XLSX" />
            <dx:ASPxPivotGridExporter ID="ASPxPivotGridExporter1" runat="server">
            </dx:ASPxPivotGridExporter>
        </div>
    </form>

</body>
</html>
