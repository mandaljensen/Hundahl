report 50007 "Precom Act. Customers"
{
    ApplicationArea = All;
    Caption = 'Precom Act. Customer';
    UsageCategory = ReportsAndAnalysis;
    ProcessingOnly = true;

    dataset
    {
        dataitem(Customer; Customer)
        {
            RequestFilterFields = "No.", "Customer Posting Group";

            trigger OnAfterGetRecord()
            var
                PrecomUpdateManagement: Codeunit "PreCom Update Management";
                RecRef: RecordRef;

            begin
                RecRef.GetTable(Customer);
                PrecomUpdateManagement.OnInsert(RecRef);
            end;
        }

    }
}
