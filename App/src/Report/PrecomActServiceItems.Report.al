report 50006 "Precom Act. Service Items"
{
    ApplicationArea = All;
    Caption = 'Precom Act. Service Items';
    UsageCategory = ReportsAndAnalysis;
    ProcessingOnly = true;

    dataset
    {
  /*       dataitem(ServiceItem; "Service Item")
        {
            RequestFilterFields = "No.", Closed;

            trigger OnAfterGetRecord()
            var
                PrecomUpdateManagement: Codeunit "PreCom Update Management";
                RecRef: RecordRef;

            begin
                RecRef.GetTable(ServiceItem);
                PrecomUpdateManagement.OnInsert(RecRef);
            end;
        } */
    }

        trigger OnPreReport()
        var
            PrecomUpdateDispatcher: Codeunit "PreCom Update Dispatcher";
        begin
            PrecomUpdateDispatcher.Run();
            CurrReport.Quit();
        end;
}
