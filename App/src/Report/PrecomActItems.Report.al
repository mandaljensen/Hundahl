report 50008 "Precom Act. Items"
{
    ApplicationArea = All;
    Caption = 'Precom Act. Items';
    UsageCategory = ReportsAndAnalysis;
    ProcessingOnly = true;

    dataset
    {
        dataitem(Item; Item)
        {
            RequestFilterFields = "No.", "Inventory Posting Group";

            trigger OnAfterGetRecord()
            var
                PrecomUpdateManagement: Codeunit "PreCom Update Management";
                RecRef: RecordRef;

            begin
                if CheckItem() then begin
                    RecRef.GetTable(Item);
                    PrecomUpdateManagement.OnInsert(RecRef);
                end;
            end;
        }
    }

    local procedure CheckItem(): Boolean
    begin
        if StrPos(Item."No.", 'BS') = 1 then
            exit(false);

        if StrPos(Item."No.", 'JO') = 1 then
            exit(false);

        if StrPos(Item."No.", 'JD') = 1 then
            exit(false);

        if StrPos(Item."No.", 'CL') = 1 then
            exit(false);

        if StrPos(Item."No.", 'WP') = 1 then
            exit(false);

        if StrPos(Item."No.", 'AG') = 1 then
            exit(false);

        if StrPos(Item."No.", 'WS') = 1 then
            exit(false);

        if StrPos(Item."No.", 'X45VA') = 1 then
            exit(false);

        if StrPos(Item."No.", 'ÅÅ') = 1 then
            exit(false);

        if StrPos(Item."No.", 'VV') = 1 then
            exit(false);

        exit(true);
    end;
}
