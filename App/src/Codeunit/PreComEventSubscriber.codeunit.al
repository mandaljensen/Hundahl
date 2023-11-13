/// <summary>
/// 14-01-2022 HMJ  Codeunit contains all event subscribers for Precom Integration.
/// </summary>
codeunit 50500 "PreCom Event Subscriber"
{
    [EventSubscriber(ObjectType::Table, database::Customer, 'OnAfterInsertEvent', '', true, true)]
    local procedure InsertCustomer(var Rec: Record Customer; RunTrigger: Boolean)
    var
        PrecomUpdateSetup: Record "PreCom Update Setup";
        PreComUpdateManagement: Codeunit "PreCom Update Management";
        RecRef: RecordRef;

    begin
        PrecomUpdateSetup.Get();
        if not PrecomUpdateSetup."Use Precom" then
            exit;

        if Rec.IsTemporary then
            exit;

        RecRef.GetTable(Rec);
        PreComUpdateManagement.OnInsert(RecRef);
    end;

    [EventSubscriber(ObjectType::Table, database::Customer, 'OnAfterModifyEvent', '', true, true)]
    local procedure ModifyCustomer(var Rec: Record Customer; var xRec: Record Customer; RunTrigger: Boolean)
    var
        PrecomUpdateSetup: Record "PreCom Update Setup";
        PreComUpdateManagement: Codeunit "PreCom Update Management";
        RecRef: RecordRef;
    begin
        PrecomUpdateSetup.Get();
        if not PrecomUpdateSetup."Use Precom" then
            exit;

        if Rec.IsTemporary then
            exit;

        RecRef.GetTable(Rec);
        PreComUpdateManagement.OnUpdate(RecRef);
    end;

    [EventSubscriber(ObjectType::Table, database::Customer, 'OnAfterDeleteEvent', '', true, true)]
    local procedure DeleteCustomer(var Rec: Record Customer; RunTrigger: Boolean)
    var
        PrecomUpdateSetup: Record "PreCom Update Setup";
        PreComUpdateManagement: Codeunit "PreCom Update Management";
        RecRef: RecordRef;
    begin
        PrecomUpdateSetup.Get();
        if not PrecomUpdateSetup."Use Precom" then
            exit;

        if Rec.IsTemporary then
            exit;

        RecRef.GetTable(Rec);
        PreComUpdateManagement.OnDelete(RecRef);
    end;

    [EventSubscriber(ObjectType::Table, database::Customer, 'OnAfterRenameEvent', '', true, true)]
    local procedure RenameCustomer(var Rec: Record Customer; var xRec: Record Customer; RunTrigger: Boolean)
    var
        PrecomUpdateSetup: Record "PreCom Update Setup";
        PreComUpdateManagement: Codeunit "PreCom Update Management";
        RecRef: RecordRef;
        xRecRef: RecordRef;
    begin
        PrecomUpdateSetup.Get();
        if not PrecomUpdateSetup."Use Precom" then
            exit;

        if Rec.IsTemporary then
            exit;

        RecRef.GetTable(Rec);
        xRecRef.GetTable(xRec);
        PreComUpdateManagement.OnRename(RecRef, xRecRef);
    end;

    [EventSubscriber(ObjectType::Table, database::Item, 'OnAfterInsertEvent', '', true, true)]
    local procedure InsertItem(var Rec: Record Item; RunTrigger: Boolean)
    var
        PrecomUpdateSetup: Record "PreCom Update Setup";
        PreComUpdateManagement: Codeunit "PreCom Update Management";
        RecRef: RecordRef;
    begin
        PrecomUpdateSetup.Get();
        if not PrecomUpdateSetup."Use Precom" then
            exit;

        if Rec.IsTemporary then
            exit;

        if Rec.Blocked then
            exit;

        RecRef.GetTable(Rec);
        PreComUpdateManagement.OnInsert(RecRef);
    end;

    [EventSubscriber(ObjectType::Table, database::Item, 'OnAfterModifyEvent', '', true, true)]
    local procedure ModifyItem(var Rec: Record Item; var xRec: Record Item; RunTrigger: Boolean)
    var
        PrecomUpdateSetup: Record "PreCom Update Setup";
        PreComUpdateManagement: Codeunit "PreCom Update Management";
        RecRef: RecordRef;
    begin
        PrecomUpdateSetup.Get();
        if not PrecomUpdateSetup."Use Precom" then
            exit;

        if Rec.IsTemporary then
            exit;

        if Rec.Blocked then
            exit;

        if not RunTrigger then
            exit;

        RecRef.GetTable(Rec);
        PreComUpdateManagement.OnUpdate(RecRef);
    end;

    [EventSubscriber(ObjectType::Table, database::Item, 'OnAfterDeleteEvent', '', true, true)]
    local procedure DeleteItem(var Rec: Record Item; RunTrigger: Boolean)
    var
        PrecomUpdateSetup: Record "PreCom Update Setup";
        PreComUpdateManagement: Codeunit "PreCom Update Management";
        RecRef: RecordRef;
    begin
        PrecomUpdateSetup.Get();
        if not PrecomUpdateSetup."Use Precom" then
            exit;

        if Rec.IsTemporary then
            exit;

        if Rec.Blocked then
            exit;

        RecRef.GetTable(Rec);
        PreComUpdateManagement.OnDelete(RecRef);
    end;

    [EventSubscriber(ObjectType::Table, database::Item, 'OnAfterRenameEvent', '', true, true)]
    local procedure RenameItem(var Rec: Record Item; var xRec: Record Item; RunTrigger: Boolean)
    var
        PrecomUpdateSetup: Record "PreCom Update Setup";
        PreComUpdateManagement: Codeunit "PreCom Update Management";
        RecRef: RecordRef;
        xRecRef: RecordRef;
    begin
        PrecomUpdateSetup.Get();
        if not PrecomUpdateSetup."Use Precom" then
            exit;

        if Rec.IsTemporary then
            exit;

        if Rec.Blocked then
            exit;

        RecRef.GetTable(Rec);
        xRecRef.GetTable(xRec);
        PreComUpdateManagement.OnRename(RecRef, xRecRef);
    end;

    [EventSubscriber(ObjectType::Table, database::"Service Item", 'OnAfterInsertEvent', '', true, true)]
    local procedure InsertServiceItem(var Rec: Record "Service Item"; RunTrigger: Boolean)
    var
        PrecomUpdateSetup: Record "PreCom Update Setup";
        PreComUpdateManagement: Codeunit "PreCom Update Management";
        RecRef: RecordRef;
    begin
        PrecomUpdateSetup.Get();
        if not PrecomUpdateSetup."Use Precom" then
            exit;

        if Rec.IsTemporary then
            exit;

        if RunTrigger then begin
            RecRef.GetTable(Rec);
            PreComUpdateManagement.OnInsert(RecRef);
        end;
    end;

    [EventSubscriber(ObjectType::Table, database::"Service Item", 'OnAfterModifyEvent', '', true, true)]
    local procedure ModifyServiceItem(var Rec: Record "Service Item"; var xRec: Record "Service Item"; RunTrigger: Boolean)
    var
        PrecomUpdateSetup: Record "PreCom Update Setup";
        PreComUpdateManagement: Codeunit "PreCom Update Management";
        RecRef: RecordRef;
    begin
        PrecomUpdateSetup.Get();
        if not PrecomUpdateSetup."Use Precom" then
            exit;

        if Rec.IsTemporary then
            exit;

        if RunTrigger then begin
            RecRef.GetTable(Rec);
            PreComUpdateManagement.OnUpdate(RecRef);
        end;
    end;

    [EventSubscriber(ObjectType::Table, database::"Service Item", 'OnAfterDeleteEvent', '', true, true)]
    local procedure DeleteServiceItem(var Rec: Record "Service Item"; RunTrigger: Boolean)
    var
        PrecomUpdateSetup: Record "PreCom Update Setup";
        PreComUpdateManagement: Codeunit "PreCom Update Management";
        RecRef: RecordRef;
    begin
        PrecomUpdateSetup.Get();
        if not PrecomUpdateSetup."Use Precom" then
            exit;

        if Rec.IsTemporary then
            exit;

        if RunTrigger then begin
            RecRef.GetTable(Rec);
            PreComUpdateManagement.OnDelete(RecRef);
        end;
    end;

    [EventSubscriber(ObjectType::Table, database::"Service Item", 'OnAfterRenameEvent', '', true, true)]
    local procedure RenameServiceItem(var Rec: Record "Service Item"; var xRec: Record "Service Item"; RunTrigger: Boolean)
    var
        PrecomUpdateSetup: Record "PreCom Update Setup";
        PreComUpdateManagement: Codeunit "PreCom Update Management";
        RecRef: RecordRef;
        xRecRef: RecordRef;
    begin
        PrecomUpdateSetup.Get();
        if not PrecomUpdateSetup."Use Precom" then
            exit;

        if Rec.IsTemporary then
            exit;

        RecRef.GetTable(Rec);
        xRecRef.GetTable(xRec);
        PreComUpdateManagement.OnRename(RecRef, xRecRef);
    end;

    [EventSubscriber(ObjectType::Table, database::"Service Line", 'OnAfterInsertEvent', '', true, true)]
    local procedure InsertServiceLine(var Rec: Record "Service Line"; RunTrigger: Boolean)
    var
        PrecomUpdateSetup: Record "PreCom Update Setup";
        ServiceHeader: Record "Service Header";
        PreComUpdateManagement: Codeunit "PreCom Update Management";
        RecRef: RecordRef;
    begin
        PrecomUpdateSetup.Get();
        if not PrecomUpdateSetup."Use Precom" then
            exit;

        if Rec.IsTemporary then
            exit;

        if Rec.Type <> Rec.Type::Item then
            exit;

        if not ServiceHeader.GET(Rec."Document Type", Rec."Document No.") then
            CLEAR(ServiceHeader);

        if RunTrigger then begin
            RecRef.GetTable(Rec);
            PreComUpdateManagement.OnInsert(RecRef);
        end;
    end;

    [EventSubscriber(ObjectType::Table, database::"Service Line", 'OnAfterModifyEvent', '', true, true)]
    local procedure ModifyServiceLine(var Rec: Record "Service Line"; var xRec: Record "Service Line"; RunTrigger: Boolean)
    var
        PrecomUpdateSetup: Record "PreCom Update Setup";
        ServiceHeader: Record "Service Header";
        PreComUpdateManagement: Codeunit "PreCom Update Management";
        RecRef: RecordRef;
    begin
        PrecomUpdateSetup.Get();
        if not PrecomUpdateSetup."Use Precom" then
            exit;

        if Rec.IsTemporary then
            exit;

        if Rec.Type <> Rec.Type::Item then
            exit;

        if not ServiceHeader.GET(Rec."Document Type", Rec."Document No.") then
            CLEAR(ServiceHeader);

        if RunTrigger then begin
            RecRef.GetTable(Rec);
            PreComUpdateManagement.OnUpdate(RecRef);
        end;
    end;

    [EventSubscriber(ObjectType::Table, database::"Service Line", 'OnAfterDeleteEvent', '', true, true)]
    local procedure DeleteServiceLine(var Rec: Record "Service Line"; RunTrigger: Boolean)
    var
        PrecomUpdateSetup: Record "PreCom Update Setup";
        ServiceHeader: Record "Service Header";
        PreComUpdateQueue: Record "PreCom Update Queue";
        PreComUpdateManagement: Codeunit "PreCom Update Management";
    begin
        PrecomUpdateSetup.Get();
        if not PrecomUpdateSetup."Use Precom" then
            exit;

        if Rec.IsTemporary then
            exit;

        if Rec.Type <> Rec.Type::Item then
            exit;

        if not ServiceHeader.GET(Rec."Document Type", Rec."Document No.") then
            CLEAR(ServiceHeader);

        if RunTrigger then begin
            PreComUpdateManagement.SetGlobalServiceLine(Rec);
            PreComUpdateQueue.Init();
            PreComUpdateQueue."Command Type" := PreComUpdateQueue."Command Type"::Delete;
            PreComUpdateManagement.WriteServiceInvLine(PreComUpdateQueue);
        end;
    end;

    [EventSubscriber(ObjectType::Table, database::"Service Header", 'OnAfterDeleteEvent', '', true, true)]
    local procedure DeleteServiceHeader(var Rec: Record "Service Header"; RunTrigger: Boolean)
    var
        PrecomUpdateSetup: Record "PreCom Update Setup";
        PreComUpdateManagement: Codeunit "PreCom Update Management";
        RecRef: RecordRef;
    begin
        PrecomUpdateSetup.Get();
        if not PrecomUpdateSetup."Use Precom" then
            exit;

        if Rec.IsTemporary then
            exit;

        if RunTrigger then begin
            RecRef.GetTable(Rec);
            PreComUpdateManagement.OnDelete(RecRef);
        end;
    end;

    /*
    [EventSubscriber(ObjectType::Table, database::"Service Header", 'OnAfterModifyEvent', '', true, true)]
    local procedure ModifyServiceHeader(var Rec: Record "Service Header"; var xRec: Record "Service Header"; RunTrigger: Boolean)
    var
        PreComUpdateManagement: Codeunit "PreCom Update Management";
    begin
        if Rec.IsTemporary then
            exit;

        if Rec.Status = xRec.Status then
            exit;

        IF Rec.Status = Rec.Status::"4" then
            PreComUpdateManagement.WriteServiceOrder(Rec, FALSE);
    end;
    */

    /*[EventSubscriber(ObjectType::Table, 352, 'OnAfterInsertEvent', '', true, true)]
    local procedure InsertDefaultDimension(var Rec: Record "Default Dimension"; RunTrigger: Boolean)
    var
    Item: Record Item;
    GeneralLedgerSetup: Record "General Ledger Setup";
    RecRef: RecordRef;
    PreComUpdateManagement: Codeunit "PreCom Update Management";
    begin
        IF Rec.IsTemporary then
         EXIT;
        
        IF Rec."Table ID" <> DATABASE::Item then
         EXIT;
        
        GeneralLedgerSetup.GET;
        IF Rec."Dimension Code" <> GeneralLedgerSetup."Department Dimension Code" then
         EXIT;
        
        IF Item.GET(Rec."No.") then
         IF NOT Item."Precom Item" then
           EXIT;
        
        RecRef.GetTable(Item);
        PreComUpdateManagement.OnUpdate(RecRef);
    end;*/

    /*[EventSubscriber(ObjectType::Table, 352, 'OnAfterModifyEvent', '', true, true)]
    local procedure ModifyDefaultDimension(var Rec: Record "Default Dimension"; var xRec: Record "Default Dimension"; RunTrigger: Boolean)
    var
    Item: Record Item;
    GeneralLedgerSetup: Record "General Ledger Setup";
    RecRef: RecordRef;
    PreComUpdateManagement: Codeunit "PreCom Update Management";
    begin
        IF Rec.IsTemporary then
         EXIT;
        
        IF Rec."Table ID" <> DATABASE::Item then
         EXIT;
        
        GeneralLedgerSetup.GET;
        IF Rec."Dimension Code" <> GeneralLedgerSetup."Department Dimension Code" then
         EXIT;
        
        IF Item.GET(Rec."No.") then
         IF NOT Item."Precom Item" then
           EXIT;
        
        RecRef.GetTable(Item);
        PreComUpdateManagement.OnUpdate(RecRef);
    end;*/

    /*[EventSubscriber(ObjectType::Table, 352, 'OnAfterDeleteEvent', '', true, true)]
    local procedure DeleteDefaultDimension(var Rec: Record "Default Dimension"; RunTrigger: Boolean)
    var
    Item: Record Item;
    GeneralLedgerSetup: Record "General Ledger Setup";
    RecRef: RecordRef;
    PreComUpdateManagement: Codeunit "PreCom Update Management";
    begin
        IF Rec.IsTemporary then
         EXIT;
        
        IF Rec."Table ID" <> DATABASE::Item then
         EXIT;
        
        GeneralLedgerSetup.GET;
        IF Rec."Dimension Code" <> GeneralLedgerSetup."Department Dimension Code" then
         EXIT;
        
        IF Item.GET(Rec."No.") then
         IF NOT Item."Precom Item" then
           EXIT;
        
        RecRef.GetTable(Item);
        PreComUpdateManagement.OnUpdate(RecRef);
    end;*/
}

