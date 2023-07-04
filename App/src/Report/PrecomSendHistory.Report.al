report 50009 "PrecomSendHistory"
{
    Caption = 'Precom Send Service Order History';
    ProcessingOnly = true;
    ApplicationArea = All;
    UsageCategory = ReportsAndAnalysis;

    dataset
    {
        dataitem(ServiceInvoiceHeader; "Service Invoice Header")
        {
            RequestFilterFields = "No.", "Posting Date";

            trigger OnAfterGetRecord()
            var
                PrecomUpdateManagement: Codeunit "PreCom Update Management";

            begin
                PrecomUpdateManagement.WriteServiceOrderHistory(ServiceInvoiceHeader, false);
            end;
        }
    }
}
