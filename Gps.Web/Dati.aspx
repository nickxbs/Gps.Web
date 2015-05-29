<%@ Page Language="C#" %>

<%@ Import Namespace="DevExpress.XtraPivotGrid" %>

<%@ Register Assembly="DevExpress.Web.ASPxPivotGrid.v13.1, Version=13.1.5.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxPivotGrid" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.ASPxPivotGrid.v13.1, Version=13.1.5.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxPivotGrid.Export" TagPrefix="dx" %>
<script runat="server">
   
    protected void Page_Load(object sender, EventArgs e)
    {
        foreach (var v in ASPxPivotGrid1.Fields)
        {
            ((PivotGridFieldBase)v).CollapseAll();
        }
        ASPxPivotGrid1.Fields[1].FilterValues.ValuesIncluded = new object[] { "Full Training_Primo giorno" };

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
            <dx:ASPxPivotGrid ID="ASPxPivotGrid1" runat="server" ClientIDMode="AutoID" DataSourceID="SqlDataSource1">
                <OptionsView DataHeadersDisplayMode="Popup"  ></OptionsView>
                <Fields>
                    <dx:PivotGridField ID="fieldGiocatore" Area="RowArea" AreaIndex="0" FieldName="Giocatore">
                    </dx:PivotGridField>
                    <dx:PivotGridField ID="fieldTipoAllenamento" AreaIndex="0" FieldName="Tipo Allenamento" Area="ColumnArea">
                    </dx:PivotGridField>
                    <dx:PivotGridField ID="fieldData" AreaIndex="1" FieldName="Data">
                    </dx:PivotGridField>
                    <dx:PivotGridField ID="fieldRuolo" AreaIndex="0" FieldName="Ruolo">
                    </dx:PivotGridField>
                    <dx:PivotGridField ID="fieldDTotmin" Area="DataArea" AreaIndex="1" FieldName="D_Tot/min" SummaryType="Average" Caption="D_Tot/min"
                        CellFormat-FormatString="N2"
                        CellFormat-FormatType="Numeric"
                        TotalCellFormat-FormatType="Numeric"
                        TotalCellFormat-FormatString="N2"
                        GrandTotalCellFormat-FormatType="Numeric"
                        GrandTotalCellFormat-FormatString="N2" GroupIndex="1" InnerGroupIndex="0">
                    </dx:PivotGridField>

                    <dx:PivotGridField ID="fieldMese" Area="ColumnArea" AreaIndex="1" FieldName="Mese">
                    </dx:PivotGridField>
                    <dx:PivotGridField ID="fieldPartitadagiocareCincasaFFuoricasa" Area="ColumnArea" AreaIndex="2" FieldName="Partita da giocare (C= in casa; F= Fuori casa)">
                    </dx:PivotGridField>
                    <dx:PivotGridField ID="fieldNumerodisedutadellasettimana" Area="ColumnArea" AreaIndex="3" FieldName="Numero di seduta della settimana">
                    </dx:PivotGridField>
                    <dx:PivotGridField ID="fieldN" Area="DataArea" AreaIndex="0" FieldName="N" Caption="#">
                    </dx:PivotGridField>
                    <dx:PivotGridField ID="fieldSpeed15Kmhm" AreaIndex="9" FieldName="Speed&gt;15 Km#h (m)" SummaryType="Average"
                        CellFormat-FormatString="N2"
                        CellFormat-FormatType="Numeric"
                        TotalCellFormat-FormatType="Numeric"
                        TotalCellFormat-FormatString="N2"
                        GrandTotalCellFormat-FormatType="Numeric"
                        GrandTotalCellFormat-FormatString="N" GroupIndex="0" InnerGroupIndex="1" TotalValueFormat-FormatString="N2" TotalValueFormat-FormatType="Numeric" Area="DataArea">
                    </dx:PivotGridField>
                    <dx:PivotGridField ID="fieldSpeed15Kmhmmin" Area="DataArea" AreaIndex="2" FieldName="Speed&gt;15 Km#h (m/min)" SummaryType="Average" CellFormat-FormatString="N2"
                        CellFormat-FormatType="Numeric"
                        TotalCellFormat-FormatType="Numeric"
                        TotalCellFormat-FormatString="N2"
                        GrandTotalCellFormat-FormatType="Numeric"
                        GrandTotalCellFormat-FormatString="N" GroupIndex="1" InnerGroupIndex="1" TotalValueFormat-FormatString="N2" TotalValueFormat-FormatType="Numeric">
                    </dx:PivotGridField>
                    <dx:PivotGridField ID="fieldDTot" AreaIndex="8" CellFormat-FormatString="N2" CellFormat-FormatType="Numeric" FieldName="D_Tot" GrandTotalCellFormat-FormatString="N" GrandTotalCellFormat-FormatType="Numeric" GroupIndex="0" InnerGroupIndex="0" SummaryType="Average" TotalCellFormat-FormatString="N2" TotalCellFormat-FormatType="Numeric" TotalValueFormat-FormatString="N2" TotalValueFormat-FormatType="Numeric" Area="DataArea">
                    </dx:PivotGridField>
                    <dx:PivotGridField ID="fieldMP20Wm" AreaIndex="11" CellFormat-FormatString="N2" CellFormat-FormatType="Numeric" FieldName="MP&gt;20W (m)" GrandTotalCellFormat-FormatString="N" GrandTotalCellFormat-FormatType="Numeric" GroupIndex="0" InnerGroupIndex="3" SummaryType="Average" TotalCellFormat-FormatString="N2" TotalCellFormat-FormatType="Numeric" TotalValueFormat-FormatString="N2" TotalValueFormat-FormatType="Numeric" Area="DataArea">
                    </dx:PivotGridField>
                    <dx:PivotGridField ID="fieldAccDecm" AreaIndex="12" CellFormat-FormatString="N2" CellFormat-FormatType="Numeric" FieldName="Acc+Dec (m)" GrandTotalCellFormat-FormatString="N" GrandTotalCellFormat-FormatType="Numeric" GroupIndex="0" InnerGroupIndex="4" SummaryType="Average" TotalCellFormat-FormatString="N2" TotalCellFormat-FormatType="Numeric" TotalValueFormat-FormatString="N2" TotalValueFormat-FormatType="Numeric" Area="DataArea">
                    </dx:PivotGridField>
                    <dx:PivotGridField ID="fieldEDm" AreaIndex="13" CellFormat-FormatString="N2" CellFormat-FormatType="Numeric" FieldName="ED (m)" GrandTotalCellFormat-FormatString="N" GrandTotalCellFormat-FormatType="Numeric" GroupIndex="0" InnerGroupIndex="5" SummaryType="Average" TotalCellFormat-FormatString="N2" TotalCellFormat-FormatType="Numeric" TotalValueFormat-FormatString="N2" TotalValueFormat-FormatType="Numeric" Area="DataArea">
                    </dx:PivotGridField>
                    <dx:PivotGridField ID="fieldSprint24kmhm" AreaIndex="10" CellFormat-FormatString="N2" CellFormat-FormatType="Numeric" FieldName="Sprint&gt;24 km#h (m)" GrandTotalCellFormat-FormatString="N" GrandTotalCellFormat-FormatType="Numeric" GroupIndex="0" InnerGroupIndex="2" SummaryType="Average" TotalCellFormat-FormatString="N2" TotalCellFormat-FormatType="Numeric" TotalValueFormat-FormatString="N2" TotalValueFormat-FormatType="Numeric" Area="DataArea">
                    </dx:PivotGridField>
                    <dx:PivotGridField ID="fieldSprint24Kmhmmin" Area="DataArea" AreaIndex="3" CellFormat-FormatString="N2" CellFormat-FormatType="Numeric" FieldName="Sprint&gt;24 Km#h (m/min)" GrandTotalCellFormat-FormatString="N" GrandTotalCellFormat-FormatType="Numeric" GroupIndex="1" InnerGroupIndex="2" SummaryType="Average" TotalValueFormat-FormatString="N2" TotalValueFormat-FormatType="Numeric">
                    </dx:PivotGridField>
                    <dx:PivotGridField ID="fieldMP20Wmmin" Area="DataArea" AreaIndex="4" CellFormat-FormatString="N2" CellFormat-FormatType="Numeric" FieldName="MP&gt;20W (m/min)" GrandTotalCellFormat-FormatString="N" GrandTotalCellFormat-FormatType="Numeric" GroupIndex="1" InnerGroupIndex="3" SummaryType="Average" TotalValueFormat-FormatString="N2" TotalValueFormat-FormatType="Numeric">
                    </dx:PivotGridField>
                    <dx:PivotGridField ID="fieldAccDecmmin" Area="DataArea" AreaIndex="5" CellFormat-FormatString="N2" CellFormat-FormatType="Numeric" FieldName="Acc+Dec (m/min)" GrandTotalCellFormat-FormatString="N" GrandTotalCellFormat-FormatType="Numeric" GroupIndex="1" InnerGroupIndex="4" SummaryType="Average" TotalValueFormat-FormatString="N2" TotalValueFormat-FormatType="Numeric">
                    </dx:PivotGridField>
                    <dx:PivotGridField ID="fieldAcc3Dec3mmin" Area="DataArea" AreaIndex="6" CellFormat-FormatString="N2" CellFormat-FormatType="Numeric" FieldName="Acc3+Dec3 (m/min)" GrandTotalCellFormat-FormatString="N" GrandTotalCellFormat-FormatType="Numeric" GroupIndex="1" InnerGroupIndex="5" SummaryType="Average" TotalValueFormat-FormatString="N2" TotalValueFormat-FormatType="Numeric">
                    </dx:PivotGridField>
                    <dx:PivotGridField ID="fieldEDRelmmin" Area="DataArea" AreaIndex="7" CellFormat-FormatString="N2" CellFormat-FormatType="Numeric" FieldName="EDRel (m/min)" GrandTotalCellFormat-FormatString="N" GrandTotalCellFormat-FormatType="Numeric" GroupIndex="1" InnerGroupIndex="6" SummaryType="Average" TotalValueFormat-FormatString="N2" TotalValueFormat-FormatType="Numeric">
                    </dx:PivotGridField>
                    <dx:PivotGridField ID="fieldEEEMPHIKjKg" Area="DataArea" AreaIndex="14" CellFormat-FormatString="N2" CellFormat-FormatType="Numeric" FieldName="EEE_MPHI (Kj/Kg)" GrandTotalCellFormat-FormatString="N" GrandTotalCellFormat-FormatType="Numeric" GroupIndex="2" InnerGroupIndex="0" SummaryType="Average" TotalValueFormat-FormatString="N2" TotalValueFormat-FormatType="Numeric">
                    </dx:PivotGridField>
                    <dx:PivotGridField ID="fieldAMPWkg" Area="DataArea" AreaIndex="15" CellFormat-FormatString="N2" CellFormat-FormatType="Numeric" FieldName="AMP (W/kg)" GrandTotalCellFormat-FormatString="N" GrandTotalCellFormat-FormatType="Numeric" GroupIndex="2" InnerGroupIndex="1" SummaryType="Average" TotalValueFormat-FormatString="N2" TotalValueFormat-FormatType="Numeric">
                    </dx:PivotGridField>
                    <dx:PivotGridField ID="fieldEEE" Area="DataArea" AreaIndex="16" CellFormat-FormatString="N2" CellFormat-FormatType="Numeric" FieldName="EEE" GrandTotalCellFormat-FormatString="N" GrandTotalCellFormat-FormatType="Numeric" GroupIndex="2" InnerGroupIndex="2" SummaryType="Average" TotalValueFormat-FormatString="N2" TotalValueFormat-FormatType="Numeric">
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
            <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>" ProviderName="<%$ ConnectionStrings:ConnectionString.ProviderName %>" SelectCommand="SELECT * FROM [vw_data]"></asp:SqlDataSource>
            <br />
            <asp:Button ID="Button1" runat="server" OnClick="Button1_Click" Text="Export XLSX" />
            <dx:ASPxPivotGridExporter ID="ASPxPivotGridExporter1" runat="server">
            </dx:ASPxPivotGridExporter>
        </div>
    </form>

</body>
</html>
