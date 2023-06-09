/// <summary>
/// 14-01-2022 HMJ  Handles all communication with Precom.
/// </summary>
codeunit 50501 "PreCom Update Management"
{
    TableNo = "PreCom Update Queue";

    trigger OnRun()
    begin
        CASE "Table ID" OF
            -30:
                ImportPreComOrderNos();
            -90:
                ImportServiceOrders(Rec);
            -100:
                ImportInvoiceInfo(Rec);
            -110:
                ImportItemInfo(Rec);
            -120:
                ImportTimeInfo(Rec);
            -130:
                ImportCostInfo(Rec);
            -140:
                ImportWorkTimeInfo(Rec);
            DATABASE::Customer:
                WriteCustomer(Rec);
            DATABASE::Item:
                WriteItem(Rec);
            DATABASE::"Service Item":
                WriteMachine(Rec);
            DATABASE::"Service Line":
                WriteServiceInvLine(Rec);
            else
                ERROR(Text001Err);
        end;
    end;

    var
        GlobalServiceLine: Record "Service Line";
        Text001Err: Label 'Table not supported!';
        //"ServiceItemLineNo.": Integer;
        DescriptionLines: array[50] of Text[50];
        GlobalWorkToDo: Text;

    procedure OnInsert(PrecomRecordRef: RecordRef)
    var
        CommandType: Option Insert,Update,Delete,Rename,CreateTable;
    begin
        CreateSyncQueueEntry(PrecomRecordRef, CommandType::Insert);
    end;

    procedure OnDelete(PrecomRecordRef: RecordRef)
    var
        CommandType: Option Insert,Update,Delete,Rename,CreateTable;
    begin
        CreateSyncQueueEntry(PrecomRecordRef, CommandType::Delete);
    end;

    procedure OnUpdate(PrecomRecordRef: RecordRef)
    var
        PrecomUpdateQueue: Record "PreCom Update Queue";
        RecID: RecordID;
        CommandType: Option Insert,Update,Delete,Rename,CreateTable;
    begin
        PrecomUpdateQueue.SetCurrentKey(RecordID, "Command Type");
        RecID := PrecomRecordRef.RECORDID;
        PrecomUpdateQueue.SetRange(RecordID, RecID);
        PrecomUpdateQueue.SetFilter("Command Type", '%1|%2', CommandType::Insert, CommandType::Delete);
        if not PrecomUpdateQueue.IsEmpty then
            Exit;

        CreateSyncQueueEntry(PrecomRecordRef, CommandType::Update);
    end;

    procedure OnRename(PrecomRecordRef: RecordRef; xPrecomRecordRef: RecordRef)
    var
        CommandType: Option Insert,Update,Delete,Rename,CreateTable;
    begin
        CreateSyncQueueEntry(xPrecomRecordRef, CommandType::Delete);
        CreateSyncQueueEntry(PrecomRecordRef, CommandType::Insert);
    end;

    local procedure CreateSyncQueueEntry(PrecomRecordRef: RecordRef; CommandType: Option Insert,Update,Delete,Rename,CreateTable)
    var
        PrecomUpdateQueue: Record "PreCom Update Queue";
        RecID: RecordID;
    begin
        PrecomUpdateQueue.SetCurrentKey(RecordID, "Command Type");
        RecID := PrecomRecordRef.RECORDID;
        PrecomUpdateQueue.SetRange(RecordID, RecID);
        PrecomUpdateQueue.SetRange("Command Type", CommandType);
        PrecomUpdateQueue.SetRange("Process Error", False);
        if PrecomUpdateQueue.FindFirst() then
            exit;

        PrecomUpdateQueue.Reset();
        PrecomUpdateQueue.LockTable(true);
        if PrecomUpdateQueue.FindLast() then;

        PrecomUpdateQueue.Init();
        PrecomUpdateQueue."Update Message ID" := PrecomUpdateQueue."Update Message ID" + 1;
        PrecomUpdateQueue.RecordID := PrecomRecordRef.RECORDID;
        PrecomUpdateQueue."Command Type" := CommandType;
        PrecomUpdateQueue."Table ID" := PrecomRecordRef.Number;
        PrecomUpdateQueue.Insert();
    end;

    procedure ReturnConnString(): Text[1024]
    var
        PrecomUpdateSetup: Record "PreCom Update Setup";
        ConnString: Text[1024];
    begin
        PrecomUpdateSetup.Get();
        ConnString :=
            'SERVER=' + PrecomUpdateSetup."PreCom SQL Server" + ';'
            + 'DATABASE=' + PrecomUpdateSetup."PreCom SQL Database" + ';'
            + 'UID=' + PrecomUpdateSetup."PreCom SQL User" + ';'
            + 'PWD=' + PrecomUpdateSetup."PreCom SQL Password";
        Exit(ConnString);
    end;

    procedure WriteCustomer(PreComUpdateQueue: Record "PreCom Update Queue")
    var
        Customer: Record Customer;
        CountryRegion: Record "Country/Region";
        ContactRec: Record Contact;
        PrecomRecordRef: RecordRef;
        FieldRef: FieldRef;
        SQLConnection: DotNet NewSqlConnection;
        SQLCommand: DotNet NewSqlCommand;
        SQLParameter: DotNet NewSqlParameter;
        SQLCommandType: DotNet CommandType;
        SQLDBType: DotNet NewSqlDbType;
    begin
        PrecomRecordRef.Get(PreComUpdateQueue.RecordID);
        FieldRef := PrecomRecordRef.Field(1);
        Customer.Get(FieldRef.Value);
        if not ContactRec.Get(Customer."Primary Contact No.") then
            Clear(ContactRec);

        if not CountryRegion.Get(Customer."Country/Region Code") then
            Clear(CountryRegion);

        if IsNull(SQLConnection) then
            SQLConnection := SQLConnection.SqlConnection();
        SQLConnection.ConnectionString(ReturnConnString());
        SQLConnection.Open();
        if IsNull(SQLCommand) then
            SQLCommand := SQLCommand.SqlCommand();
        SQLCommand.Connection(SQLConnection);
        SQLCommand.CommandType(SQLCommandType.StoredProcedure);
        SQLCommand.CommandText('WriteCustomer');

        SQLParameter := SQLParameter.SqlParameter();
        SQLParameter.ParameterName := '@CustomerNo';
        SQLParameter.SqlDbType := SQLDBType.NVarChar;
        SQLParameter.Size := 255;
        SQLParameter.Value := Customer."No.";
        SQLCommand.Parameters.Add(SQLParameter);

        SQLParameter := SQLParameter.SqlParameter();
        SQLParameter.ParameterName := '@Name';
        SQLParameter.SqlDbType := SQLDBType.NVarChar;
        SQLParameter.Size := 255;
        SQLParameter.Value := Customer.Name;
        SQLCommand.Parameters.Add(SQLParameter);

        SQLParameter := SQLParameter.SqlParameter();
        SQLParameter.ParameterName := '@TelephoneNumber';
        SQLParameter.SqlDbType := SQLDBType.NVarChar;
        SQLParameter.Size := 255;
        SQLParameter.Value := Customer."Phone No.";
        SQLCommand.Parameters.Add(SQLParameter);

        SQLParameter := SQLParameter.SqlParameter();
        SQLParameter.ParameterName := '@MobileNumber';
        SQLParameter.SqlDbType := SQLDBType.NVarChar;
        SQLParameter.Size := 255;
        SQLParameter.Value := Customer."Mobile Phone No.";
        SQLCommand.Parameters.Add(SQLParameter);

        SQLParameter := SQLParameter.SqlParameter();
        SQLParameter.ParameterName := '@EmailAddress';
        SQLParameter.SqlDbType := SQLDBType.NVarChar;
        SQLParameter.Size := 255;
        SQLParameter.Value := Customer."E-Mail";
        SQLCommand.Parameters.Add(SQLParameter);

        SQLParameter := SQLParameter.SqlParameter();
        SQLParameter.ParameterName := '@Blocked';
        SQLParameter.SqlDbType := SQLDBType.Bit;
        if Customer.Blocked <> Customer.Blocked::" " then
            SQLParameter.Value := TRUE
        else
            SQLParameter.Value := False;
        SQLCommand.Parameters.Add(SQLParameter);

        SQLParameter := SQLParameter.SqlParameter();
        SQLParameter.ParameterName := '@Address';
        SQLParameter.SqlDbType := SQLDBType.NVarChar;
        SQLParameter.Size := 255;
        SQLParameter.Value := Customer.Address;
        SQLCommand.Parameters.Add(SQLParameter);

        SQLParameter := SQLParameter.SqlParameter();
        SQLParameter.ParameterName := '@Zipcode';
        SQLParameter.SqlDbType := SQLDBType.NVarChar;
        SQLParameter.Size := 255;
        SQLParameter.Value := Customer."Post Code";
        SQLCommand.Parameters.Add(SQLParameter);

        SQLParameter := SQLParameter.SqlParameter();
        SQLParameter.ParameterName := '@City';
        SQLParameter.SqlDbType := SQLDBType.NVarChar;
        SQLParameter.Size := 255;
        SQLParameter.Value := Customer.City;
        SQLCommand.Parameters.Add(SQLParameter);

        SQLParameter := SQLParameter.SqlParameter();
        SQLParameter.ParameterName := '@Country';
        SQLParameter.SqlDbType := SQLDBType.NVarChar;
        SQLParameter.Size := 255;
        SQLParameter.Value := CountryRegion.Name;
        SQLCommand.Parameters.Add(SQLParameter);

        SQLParameter := SQLParameter.SqlParameter();
        SQLParameter.ParameterName := '@ContactPerson';
        SQLParameter.SqlDbType := SQLDBType.NVarChar;
        SQLParameter.Size := 255;
        SQLParameter.Value := ContactRec.Name;
        SQLCommand.Parameters.Add(SQLParameter);

        SQLParameter := SQLParameter.SqlParameter();
        SQLParameter.ParameterName := '@ContactMobileNumber';
        SQLParameter.SqlDbType := SQLDBType.NVarChar;
        SQLParameter.Size := 255;
        SQLParameter.Value := ContactRec."Mobile Phone No.";
        SQLCommand.Parameters.Add(SQLParameter);

        SQLParameter := SQLParameter.SqlParameter();
        SQLParameter.ParameterName := '@ContactEmail';
        SQLParameter.SqlDbType := SQLDBType.NVarChar;
        SQLParameter.Size := 255;
        SQLParameter.Value := ContactRec."E-Mail";
        SQLCommand.Parameters.Add(SQLParameter);

        SQLParameter := SQLParameter.SqlParameter();
        SQLParameter.ParameterName := '@Deleted';
        SQLParameter.SqlDbType := SQLDBType.Bit;
        if PreComUpdateQueue."Command Type" = PreComUpdateQueue."Command Type"::Delete then
            SQLParameter.Value := TRUE
        else
            SQLParameter.Value := False;
        SQLCommand.Parameters.Add(SQLParameter);

        SQLCommand.ExecuteNonQuery();

        Clear(SQLParameter);
        Clear(SQLCommand);
        SQLConnection.Close();
        Clear(SQLConnection);
    end;

    procedure WriteItem(PreComUpdateQueue: Record "PreCom Update Queue")
    var
        Item: Record Item;
        ItemCrossReference: Record "Item Cross Reference";
        PrecomRecordRef: RecordRef;
        FieldRef: FieldRef;
        SQLConnection: DotNet NewSqlConnection;
        SQLCommand: DotNet NewSqlCommand;
        SQLParameter: DotNet NewSqlParameter;
        SQLCommandType: DotNet CommandType;
        SQLDBType: DotNet NewSqlDbType;
        Barcodes: array[10] of Text[50];
        i: Integer;

    begin
        PrecomRecordRef.Get(PreComUpdateQueue.RecordID);
        FieldRef := PrecomRecordRef.FIELD(1);
        Item.Get(FieldRef.Value);
        WITH Item DO begin
            if IsNull(SQLConnection) then
                SQLConnection := SQLConnection.SqlConnection();
            SQLConnection.ConnectionString(ReturnConnString());
            SQLConnection.Open();
            if IsNull(SQLCommand) then
                SQLCommand := SQLCommand.SqlCommand();
            SQLCommand.Connection(SQLConnection);
            SQLCommand.CommandType(SQLCommandType.StoredProcedure);
            SQLCommand.CommandText('WriteItem');

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@ItemNo';
            SQLParameter.SqlDbType := SQLDBType.NVarChar;
            SQLParameter.Size := 255;
            SQLParameter.Value := "No.";
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@Description';
            SQLParameter.SqlDbType := SQLDBType.NVarChar;
            SQLParameter.Size := 255;
            SQLParameter.Value := Description;
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@Price';
            SQLParameter.SqlDbType := SQLDBType.Float;
            SQLParameter.Value := "Unit Price";
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@UOM';
            SQLParameter.SqlDbType := SQLDBType.NVarChar;
            SQLParameter.Size := 255;
            SQLParameter.Value := "Base Unit of Measure";
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@ItemNo2';
            SQLParameter.SqlDbType := SQLDBType.NVarChar;
            SQLParameter.Size := 255;
            SQLParameter.Value := "No. 2";
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@VendorNo';
            SQLParameter.SqlDbType := SQLDBType.NVarChar;
            SQLParameter.Size := 255;
            SQLParameter.Value := "Vendor No.";
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@VendorItemNo';
            SQLParameter.SqlDbType := SQLDBType.NVarChar;
            SQLParameter.Size := 255;
            SQLParameter.Value := "Vendor Item No.";
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@Description2';
            SQLParameter.SqlDbType := SQLDBType.NVarChar;
            SQLParameter.Size := 255;
            SQLParameter.Value := "Description 2";
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@ArticleGroup';
            SQLParameter.SqlDbType := SQLDBType.NVarChar;
            SQLParameter.Size := 255;
            SQLParameter.Value := "Global Dimension 1 Code";
            SQLCommand.Parameters.Add(SQLParameter);

            /*
            GeneralLedgerSetup.Get();
            DefaultDimension.Reset();
            DefaultDimension.SetRange("Table ID",DATABASE::Item);
            DefaultDimension.SetRange("No.","No.");
            DefaultDimension.SetRange("Dimension Code",GeneralLedgerSetup."Department Dimension Code");
            if not DefaultDimension.FindFirst() then
            Clear(DefaultDimension);
        
            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@Organization';
            SQLParameter.SqlDbType := SQLDBType.NVarChar;
            SQLParameter.Size := 255;
            SQLParameter.Value := DefaultDimension."Dimension Value Code";
            SQLCommand.Parameters.Add(SQLParameter);
            */

            i := 1;
            ItemCrossReference.Reset();
            ItemCrossReference.SetRange("Item No.", Item."No.");
            ItemCrossReference.SetFilter("Cross-Reference Type", '%1|%2', ItemCrossReference."Cross-Reference Type"::Vendor, ItemCrossReference."Cross-Reference Type"::"Bar Code");
            if ItemCrossReference.FindSet(False, False) then
                repeat
                    Barcodes[i] := ItemCrossReference."Cross-Reference No.";
                    i += 1;
                until (ItemCrossReference.Next() <= 0);

            for i := 1 to 10 do begin
                SQLParameter := SQLParameter.SqlParameter();
                if StrLen(Format(i)) = 1 then
                    SQLParameter.ParameterName := '@Barcode0' + Format(i)
                else
                    SQLParameter.ParameterName := '@Barcode' + Format(i);
                SQLParameter.SqlDbType := SQLDBType.NVarChar;
                SQLParameter.Size := 255;
                SQLParameter.Value := Barcodes[i];
                SQLCommand.Parameters.Add(SQLParameter);
            end;

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@Deleted';
            SQLParameter.SqlDbType := SQLDBType.Bit;
            if PreComUpdateQueue."Command Type" = PreComUpdateQueue."Command Type"::Delete then
                SQLParameter.Value := TRUE
            else
                SQLParameter.Value := False;
            SQLCommand.Parameters.Add(SQLParameter);

            SQLCommand.ExecuteNonQuery();

            Clear(SQLParameter);
            Clear(SQLCommand);
            SQLConnection.Close();
            Clear(SQLConnection);
        end;
    end;

    procedure WriteMachine(PreComUpdateQueue: Record "PreCom Update Queue")
    var
        ServiceItem: Record "Service Item";
        ServiceContractLine: Record "Service Contract Line";
        ServiceItemGroup: Record "Service Item Group";
        ContactRec: Record Contact;
        PrecomRecordRef: RecordRef;
        FieldRef: FieldRef;
        SQLConnection: DotNet NewSqlConnection;
        SQLCommand: DotNet NewSqlCommand;
        SQLParameter: DotNet NewSqlParameter;
        SQLCommandType: DotNet CommandType;
        SQLDBType: DotNet NewSqlDbType;
    begin
        PrecomRecordRef.Get(PreComUpdateQueue.RecordID);
        FieldRef := PrecomRecordRef.FIELD(1);
        ServiceItem.Get(FieldRef.Value);

        if not ContactRec.Get(ServiceItem.Contact) then
            Clear(ContactRec);

        ServiceContractLine.Reset();
        ServiceContractLine.SetRange("Service Item No.", ServiceItem."No.");
        if not ServiceContractLine.FindLast() then
            Clear(ServiceContractLine);

        ServiceItem.CalcFields("Item Description", "Ship-to Address", Address, Contact);
        if ServiceItem."Ship-to Address" = '' then
            ServiceItem."Ship-to Address" := ServiceItem.Address;

        if not ServiceItemGroup.Get(ServiceItem."Service Item Group Code") then
            Clear(ServiceItemGroup);

        if IsNull(SQLConnection) then
            SQLConnection := SQLConnection.SqlConnection();
        SQLConnection.ConnectionString(ReturnConnString());
        SQLConnection.Open();
        if IsNull(SQLCommand) then
            SQLCommand := SQLCommand.SqlCommand();
        SQLCommand.Connection(SQLConnection);
        SQLCommand.CommandType(SQLCommandType.StoredProcedure);
        SQLCommand.CommandText('WriteMachine');

        SQLParameter := SQLParameter.SqlParameter();
        SQLParameter.ParameterName := '@CustomerNumber';
        SQLParameter.SqlDbType := SQLDBType.NVarChar;
        SQLParameter.Size := 255;
        SQLParameter.Value := ServiceItem."Customer No.";
        SQLCommand.Parameters.Add(SQLParameter);

        SQLParameter := SQLParameter.SqlParameter();
        SQLParameter.ParameterName := '@EquipmentNumber';
        SQLParameter.SqlDbType := SQLDBType.NVarChar;
        SQLParameter.Size := 255;
        SQLParameter.Value := ServiceItem."No.";
        SQLCommand.Parameters.Add(SQLParameter);

        ServiceItem.CalcFields("Item Description");
        SQLParameter := SQLParameter.SqlParameter();
        SQLParameter.ParameterName := '@EquipmentName';
        SQLParameter.SqlDbType := SQLDBType.NVarChar;
        SQLParameter.Size := 255;
        SQLParameter.Value := ServiceItem.Description;
        SQLCommand.Parameters.Add(SQLParameter);

        SQLParameter := SQLParameter.SqlParameter();
        SQLParameter.ParameterName := '@CustomerContact';
        SQLParameter.SqlDbType := SQLDBType.NVarChar;
        SQLParameter.Size := 255;
        SQLParameter.Value := ContactRec.Name;
        SQLCommand.Parameters.Add(SQLParameter);

        SQLParameter := SQLParameter.SqlParameter();
        SQLParameter.ParameterName := '@CustomerPhone';
        SQLParameter.SqlDbType := SQLDBType.NVarChar;
        SQLParameter.Size := 255;
        SQLParameter.Value := ContactRec."Phone No.";
        SQLCommand.Parameters.Add(SQLParameter);

        SQLParameter := SQLParameter.SqlParameter();
        SQLParameter.ParameterName := '@CustomerEmail';
        SQLParameter.SqlDbType := SQLDBType.NVarChar;
        SQLParameter.Size := 255;
        SQLParameter.Value := ContactRec."E-Mail";
        SQLCommand.Parameters.Add(SQLParameter);

        SQLParameter := SQLParameter.SqlParameter();
        SQLParameter.ParameterName := '@SerialNumber';
        SQLParameter.SqlDbType := SQLDBType.NVarChar;
        SQLParameter.Size := 255;
        SQLParameter.Value := ServiceItem."Serial No.";
        SQLCommand.Parameters.Add(SQLParameter);

        SQLParameter := SQLParameter.SqlParameter();
        SQLParameter.ParameterName := '@ExternalSerialNumber';
        SQLParameter.SqlDbType := SQLDBType.NVarChar;
        SQLParameter.Size := 255;
        SQLParameter.Value := ServiceItem."External Serial No.";
        SQLCommand.Parameters.Add(SQLParameter);

        SQLParameter := SQLParameter.SqlParameter();
        SQLParameter.ParameterName := '@WarrantyStartDate';
        SQLParameter.SqlDbType := SQLDBType.Date;
        if ServiceItem."Warranty Starting Date (Parts)" <> 0D then
            SQLParameter.Value := ServiceItem."Warranty Starting Date (Parts)"
        else
            SQLParameter.Value := '01-01-1960';
        SQLCommand.Parameters.Add(SQLParameter);

        SQLParameter := SQLParameter.SqlParameter();
        SQLParameter.ParameterName := '@WarrantyEndDate';
        SQLParameter.SqlDbType := SQLDBType.Date;
        if ServiceItem."Warranty ending Date (Parts)" <> 0D then
            SQLParameter.Value := ServiceItem."Warranty ending Date (Parts)"
        else
            SQLParameter.Value := '01-01-1960';
        SQLCommand.Parameters.Add(SQLParameter);

        SQLParameter := SQLParameter.SqlParameter();
        SQLParameter.ParameterName := '@Type';
        SQLParameter.SqlDbType := SQLDBType.NVarChar;
        SQLParameter.Size := 255;
        SQLParameter.Value := ServiceItem."Service Item Group Code";
        SQLCommand.Parameters.Add(SQLParameter);

        SQLParameter := SQLParameter.SqlParameter();
        SQLParameter.ParameterName := '@Reading';
        SQLParameter.SqlDbType := SQLDBType.Int;
        SQLParameter.Value := ServiceItem."Counter (Last)";
        SQLCommand.Parameters.Add(SQLParameter);

        SQLParameter := SQLParameter.SqlParameter();
        SQLParameter.ParameterName := '@EquipmentType';
        SQLParameter.SqlDbType := SQLDBType.NVarChar;
        SQLParameter.Size := 255;
        SQLParameter.Value := ServiceItem."Model Class Code";
        SQLCommand.Parameters.Add(SQLParameter);

        SQLParameter := SQLParameter.SqlParameter();
        SQLParameter.ParameterName := '@Manufacturer';
        SQLParameter.SqlDbType := SQLDBType.NVarChar;
        SQLParameter.Size := 255;
        SQLParameter.Value := ServiceItem."Make Code";
        SQLCommand.Parameters.Add(SQLParameter);

        SQLParameter := SQLParameter.SqlParameter();
        SQLParameter.ParameterName := '@DeliveryPlace';
        SQLParameter.SqlDbType := SQLDBType.NVarChar;
        SQLParameter.Size := 255;
        SQLParameter.Value := ServiceItem."Ship-to Post Code";
        SQLCommand.Parameters.Add(SQLParameter);

        SQLParameter := SQLParameter.SqlParameter();
        SQLParameter.ParameterName := '@DeliveryDateCustomer';
        SQLParameter.SqlDbType := SQLDBType.Date;
        SQLParameter.Size := 255;
        if ServiceItem."Installation Date" <> 0D then
            SQLParameter.Value := ReturnLocalDateTime(ServiceItem."Installation Date", 000000T)
        else
            SQLParameter.Value := '01-01-1960';
        SQLCommand.Parameters.Add(SQLParameter);

        SQLParameter := SQLParameter.SqlParameter();
        SQLParameter.ParameterName := '@Application';
        SQLParameter.SqlDbType := SQLDBType.NVarChar;
        SQLParameter.Size := 255;
        SQLParameter.Value := ServiceItem."Description 2";
        SQLCommand.Parameters.Add(SQLParameter);

        SQLParameter := SQLParameter.SqlParameter();
        SQLParameter.ParameterName := '@LastServicingDate';
        SQLParameter.SqlDbType := SQLDBType.Date;
        SQLParameter.Size := 255;
        if ServiceItem."Last Service Date" <> 0D then
            SQLParameter.Value := ServiceItem."Last Service Date"
        else
            SQLParameter.Value := '01-01-1960';
        SQLCommand.Parameters.Add(SQLParameter);

        SQLParameter := SQLParameter.SqlParameter();
        SQLParameter.ParameterName := '@NextServicingDate';
        SQLParameter.SqlDbType := SQLDBType.Date;
        SQLParameter.Size := 255;
        if ServiceContractLine."Next Planned Service Date" <> 0D then
            SQLParameter.Value := ServiceContractLine."Next Planned Service Date"
        else
            SQLParameter.Value := '01-01-1960';
        SQLCommand.Parameters.Add(SQLParameter);

        SQLParameter := SQLParameter.SqlParameter();
        SQLParameter.ParameterName := '@Department';
        SQLParameter.SqlDbType := SQLDBType.NVarChar;
        SQLParameter.Size := 255;
        SQLParameter.Value := ServiceItem."Salesperson Code";
        SQLCommand.Parameters.Add(SQLParameter);

        SQLParameter := SQLParameter.SqlParameter();
        SQLParameter.ParameterName := '@State';
        SQLParameter.SqlDbType := SQLDBType.NVarChar;
        SQLParameter.Size := 255;
        SQLParameter.Value := ServiceItem."State Code";
        SQLCommand.Parameters.Add(SQLParameter);

        SQLParameter := SQLParameter.SqlParameter();
        SQLParameter.ParameterName := '@WarrantyType';
        SQLParameter.SqlDbType := SQLDBType.NVarChar;
        SQLParameter.Size := 255;
        SQLParameter.Value := ServiceItem."Warrenty Type Code";
        SQLCommand.Parameters.Add(SQLParameter);

        SQLParameter := SQLParameter.SqlParameter();
        SQLParameter.ParameterName := '@WarrantyAdditionalDate';
        SQLParameter.SqlDbType := SQLDBType.Date;
        SQLParameter.Size := 255;
        if ServiceItem."Warranty ending Date (Parts)" <> 0D then
            SQLParameter.Value := ServiceItem."Warranty ending Date (Parts)"
        else
            SQLParameter.Value := '01-01-1960';
        SQLCommand.Parameters.Add(SQLParameter);

        SQLParameter := SQLParameter.SqlParameter();
        SQLParameter.ParameterName := '@MotorId';
        SQLParameter.SqlDbType := SQLDBType.NVarChar;
        SQLParameter.Size := 255;
        SQLParameter.Value := ServiceItem."Engine No.";
        SQLCommand.Parameters.Add(SQLParameter);

        SQLParameter := SQLParameter.SqlParameter();
        SQLParameter.ParameterName := '@Blocked';
        SQLParameter.SqlDbType := SQLDBType.Bit;
        SQLParameter.Value := ServiceItem.Closed;
        SQLCommand.Parameters.Add(SQLParameter);

        SQLParameter := SQLParameter.SqlParameter();
        SQLParameter.ParameterName := '@Deleted';
        SQLParameter.SqlDbType := SQLDBType.Bit;
        if PreComUpdateQueue."Command Type" = PreComUpdateQueue."Command Type"::Delete then
            SQLParameter.Value := TRUE
        else
            SQLParameter.Value := False;
        SQLCommand.Parameters.Add(SQLParameter);

        SQLCommand.ExecuteNonQuery();

        Clear(SQLParameter);
        Clear(SQLCommand);
        SQLConnection.Close();
        Clear(SQLConnection);
    end;

    procedure WriteServiceOrder(ServiceHeader: Record "Service Header"; Deleted: Boolean)
    var
        ServiceItemLine: Record "Service Item Line";
        ServiceLine: Record "Service Line";
        PreComUpdateQueue: Record "PreCom Update Queue";
        PrecomRecordRef: RecordRef;
        SQLConnection: DotNet NewSqlConnection;
        SQLCommand: DotNet NewSqlCommand;
        SQLParameter: DotNet NewSqlParameter;
        SQLCommandType: DotNet CommandType;
        SQLDBType: DotNet NewSqlDbType;
        WorkToDo: Text;
        StopWorkToDo: Boolean;
    begin
        WITH ServiceHeader DO begin
            ServiceItemLine.Reset();
            ServiceItemLine.SetRange("Document Type", "Document Type");
            ServiceItemLine.SetRange("Document No.", "No.");
            if not ServiceItemLine.FindFirst() then
                Clear(ServiceItemLine);

            if IsNull(SQLConnection) then
                SQLConnection := SQLConnection.SqlConnection();
            SQLConnection.ConnectionString(ReturnConnString());
            SQLConnection.Open();
            if IsNull(SQLCommand) then
                SQLCommand := SQLCommand.SqlCommand();
            SQLCommand.Connection(SQLConnection);
            SQLCommand.CommandType(SQLCommandType.StoredProcedure);
            SQLCommand.CommandText('WriteServiceOrder');

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@OrderNo';
            SQLParameter.SqlDbType := SQLDBType.NVarChar;
            SQLParameter.Size := 255;
            SQLParameter.Value := ''; // FIXME: Hvilket ordrenr.
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@OrderTitle';
            SQLParameter.SqlDbType := SQLDBType.NVarChar;
            SQLParameter.Size := 255;
            SQLParameter.Value := Description;
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@CustomerNo';
            SQLParameter.SqlDbType := SQLDBType.NVarChar;
            SQLParameter.Size := 255;
            SQLParameter.Value := "Customer No.";
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@Address';
            SQLParameter.SqlDbType := SQLDBType.NVarChar;
            SQLParameter.Size := 255;
            SQLParameter.Value := Address;
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@PostCode';
            SQLParameter.SqlDbType := SQLDBType.NVarChar;
            SQLParameter.Size := 255;
            SQLParameter.Value := "Post Code";
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@City';
            SQLParameter.SqlDbType := SQLDBType.NVarChar;
            SQLParameter.Size := 255;
            SQLParameter.Value := City;
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@Customerphonenumber';
            SQLParameter.SqlDbType := SQLDBType.NVarChar;
            SQLParameter.Size := 255;
            SQLParameter.Value := "Phone No.";
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@Name';
            SQLParameter.SqlDbType := SQLDBType.NVarChar;
            SQLParameter.Size := 255;
            SQLParameter.Value := Name;
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@EmailAddress';
            SQLParameter.SqlDbType := SQLDBType.NVarChar;
            SQLParameter.Size := 255;
            SQLParameter.Value := "E-Mail";
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@InvoiceCustomerNo';
            SQLParameter.SqlDbType := SQLDBType.NVarChar;
            SQLParameter.Size := 255;
            SQLParameter.Value := "Bill-to Customer No.";
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@Startdatetime';
            SQLParameter.SqlDbType := SQLDBType.DateTime;
            if "Starting Date" <> 0D then
                SQLParameter.Value := ReturnLocalDateTime("Starting Date", "Starting Time") //FORMAT(CREATEDATETIME("Starting Date","Starting Time"),0,'<Year4>/<Month,2>/<Day,2> <Hours24>:<Minutes,2>:<Seconds,2>')
            else
                SQLParameter.Value := CREATEDATETIME(19600101D, 000000T); //FORMAT(CREATEDATETIME(010160D,000000T),0,'<Year4>-<Month,2>-<Day,2> <Hours24>:<Minutes,2>:<Seconds,2>');
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@PlannedDuration';
            SQLParameter.SqlDbType := SQLDBType.Int;
            SQLParameter.Value := 0;
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@ERPReference';
            SQLParameter.SqlDbType := SQLDBType.NVarChar;
            SQLParameter.Size := 255;
            SQLParameter.Value := "No.";
            SQLCommand.Parameters.Add(SQLParameter);

            StopWorkToDo := False;
            //WorkToDo := Description;
            ServiceLine.Reset();
            ServiceLine.SetRange("Document Type", "Document Type");
            ServiceLine.SetRange("Document No.", "No.");
            if ServiceLine.FindSet() then
                Repeat
                    if ServiceLine.Type = ServiceLine.Type::" " then
                        WorkToDo := WorkToDo + ' ' + ServiceLine.Description;

                    if ServiceLine.Type > ServiceLine.Type::" " then
                        StopWorkToDo := TRUE;
                Until ((ServiceLine.Next() <= 0) OR (StopWorkToDo));

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@WorkToDo';
            SQLParameter.SqlDbType := SQLDBType.NVarChar;
            SQLParameter.Size := 1024;
            SQLParameter.Value := COPYSTR(WorkToDo, 1, 1024);
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@EquipmentNo';
            SQLParameter.SqlDbType := SQLDBType.NVarChar;
            SQLParameter.Size := 255;
            SQLParameter.Value := ServiceItemLine."Service Item No.";
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@OrderTypeID';
            SQLParameter.SqlDbType := SQLDBType.NVarChar;
            SQLParameter.Size := 255;
            SQLParameter.Value := ServiceHeader."Service Order Type";
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@UserGroupID';
            SQLParameter.SqlDbType := SQLDBType.NVarChar;
            SQLParameter.Size := 255;
            SQLParameter.Value := ServiceHeader."Location Code";
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@Contactname';
            SQLParameter.SqlDbType := SQLDBType.NVarChar;
            SQLParameter.Size := 255;
            SQLParameter.Value := ServiceHeader."Contact Name";
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@Reference';
            SQLParameter.SqlDbType := SQLDBType.NVarChar;
            SQLParameter.Size := 255;
            SQLParameter.Value := ServiceHeader."Your Reference";
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@RepairStatus';
            SQLParameter.SqlDbType := SQLDBType.NVarChar;
            SQLParameter.Size := 255;
            SQLParameter.Value := ServiceItemLine."Repair Status Code";
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@DispatcherExternalId';
            SQLParameter.SqlDbType := SQLDBType.NVarChar;
            SQLParameter.Size := 255;
            SQLParameter.Value := ServiceHeader."Salesperson Code";
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@PriorityId';
            SQLParameter.SqlDbType := SQLDBType.NVarChar;
            SQLParameter.Size := 255;
            SQLParameter.Value := FORMAT(ServiceHeader.Priority);
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@CauseCodeId';
            SQLParameter.SqlDbType := SQLDBType.NVarChar;
            SQLParameter.Size := 255;
            SQLParameter.Value := "Reason Code";
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@ReceivedDate';
            SQLParameter.SqlDbType := SQLDBType.DateTime;
            //SQLParameter.Size := 255;
            SQLParameter.Value := ReturnLocalDateTime("Order Date", "Order Time");
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@ResponsibilityCenter';
            SQLParameter.SqlDbType := SQLDBType.NVarChar;
            SQLParameter.Size := 255;
            SQLParameter.Value := "Responsibility Center";
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@Deleted';
            SQLParameter.SqlDbType := SQLDBType.Bit;
            if Deleted then
                SQLParameter.Value := TRUE
            else
                SQLParameter.Value := False;
            SQLCommand.Parameters.Add(SQLParameter);

            SQLCommand.ExecuteNonQuery();

            Clear(SQLParameter);
            Clear(SQLCommand);
            SQLConnection.Close();
            Clear(SQLConnection);

            ServiceLine.Reset();
            ServiceLine.SetRange("Document Type", ServiceHeader."Document Type");
            ServiceLine.SetRange("Document No.", ServiceHeader."No.");
            ServiceLine.SetRange(Type, ServiceLine.Type::Item);
            ServiceLine.SetRange("Work Type Code", '');
            if ServiceLine.FindSet() then
                Repeat
                    PrecomRecordRef.GetTable(ServiceLine);
                    PreComUpdateQueue.Init();
                    PreComUpdateQueue."Update Message ID" := 999999999;
                    PreComUpdateQueue.RecordID := PrecomRecordRef.RECORDID;
                    PreComUpdateQueue."Command Type" := PreComUpdateQueue."Command Type"::Insert;
                    PreComUpdateQueue."Table ID" := PrecomRecordRef.NUMBER;
                    WriteServiceInvLine(PreComUpdateQueue);
                Until (ServiceLine.Next() <= 0);
        end;
    end;

    procedure WriteServiceOrderHistory(ServiceHeader: Record "Service Header"; Deleted: Boolean)
    var
        ServiceItemLine: Record "Service Item Line";
        ServiceLine: Record "Service Line";
        PreComUpdateQueue: Record "PreCom Update Queue";
        PrecomRecordRef: RecordRef;
        SQLConnection: DotNet NewSqlConnection;
        SQLCommand: DotNet NewSqlCommand;
        SQLParameter: DotNet NewSqlParameter;
        SQLCommandType: DotNet CommandType;
        SQLDBType: DotNet NewSqlDbType;
        WorkToDo: Text;
        StopWorkToDo: Boolean;
    begin
        WITH ServiceHeader DO begin
            ServiceItemLine.Reset();
            ServiceItemLine.SetRange("Document Type", "Document Type");
            ServiceItemLine.SetRange("Document No.", "No.");
            if not ServiceItemLine.FindFirst() then
                Clear(ServiceItemLine);

            if IsNull(SQLConnection) then
                SQLConnection := SQLConnection.SqlConnection();
            SQLConnection.ConnectionString(ReturnConnString());
            SQLConnection.Open();
            if IsNull(SQLCommand) then
                SQLCommand := SQLCommand.SqlCommand();
            SQLCommand.Connection(SQLConnection);
            SQLCommand.CommandType(SQLCommandType.StoredProcedure);
            SQLCommand.CommandText('WriteServiceOrderHistory');

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@OrderNo';
            SQLParameter.SqlDbType := SQLDBType.NVarChar;
            SQLParameter.Size := 255;
            SQLParameter.Value := ''; // FIXME: Hvilket ordrenr.
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@OrderTitle';
            SQLParameter.SqlDbType := SQLDBType.NVarChar;
            SQLParameter.Size := 255;
            SQLParameter.Value := Description;
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@CustomerNo';
            SQLParameter.SqlDbType := SQLDBType.NVarChar;
            SQLParameter.Size := 255;
            SQLParameter.Value := "Customer No.";
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@Address';
            SQLParameter.SqlDbType := SQLDBType.NVarChar;
            SQLParameter.Size := 255;
            SQLParameter.Value := Address;
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@PostCode';
            SQLParameter.SqlDbType := SQLDBType.NVarChar;
            SQLParameter.Size := 255;
            SQLParameter.Value := "Post Code";
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@City';
            SQLParameter.SqlDbType := SQLDBType.NVarChar;
            SQLParameter.Size := 255;
            SQLParameter.Value := City;
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@Customerphonenumber';
            SQLParameter.SqlDbType := SQLDBType.NVarChar;
            SQLParameter.Size := 255;
            SQLParameter.Value := "Phone No.";
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@Name';
            SQLParameter.SqlDbType := SQLDBType.NVarChar;
            SQLParameter.Size := 255;
            SQLParameter.Value := Name;
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@EmailAddress';
            SQLParameter.SqlDbType := SQLDBType.NVarChar;
            SQLParameter.Size := 255;
            SQLParameter.Value := "E-Mail";
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@InvoiceCustomerNo';
            SQLParameter.SqlDbType := SQLDBType.NVarChar;
            SQLParameter.Size := 255;
            SQLParameter.Value := "Bill-to Customer No.";
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@Startdatetime';
            SQLParameter.SqlDbType := SQLDBType.DateTime;
            if "Starting Date" <> 0D then
                SQLParameter.Value := ReturnLocalDateTime("Starting Date", "Starting Time") //FORMAT(CREATEDATETIME("Starting Date","Starting Time"),0,'<Year4>/<Month,2>/<Day,2> <Hours24>:<Minutes,2>:<Seconds,2>')
            else
                SQLParameter.Value := CREATEDATETIME(19600101D, 000000T); //FORMAT(CREATEDATETIME(010160D,000000T),0,'<Year4>-<Month,2>-<Day,2> <Hours24>:<Minutes,2>:<Seconds,2>');
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@PlannedDuration';
            SQLParameter.SqlDbType := SQLDBType.Int;
            SQLParameter.Value := 0;
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@ERPReference';
            SQLParameter.SqlDbType := SQLDBType.NVarChar;
            SQLParameter.Size := 255;
            SQLParameter.Value := "No.";
            SQLCommand.Parameters.Add(SQLParameter);

            StopWorkToDo := False;
            //WorkToDo := Description;
            ServiceLine.Reset();
            ServiceLine.SetRange("Document Type", "Document Type");
            ServiceLine.SetRange("Document No.", "No.");
            if ServiceLine.FindSet() then
                Repeat
                    if ServiceLine.Type = ServiceLine.Type::" " then
                        WorkToDo := WorkToDo + ' ' + ServiceLine.Description;

                    if ServiceLine.Type > ServiceLine.Type::" " then
                        StopWorkToDo := TRUE;
                Until ((ServiceLine.Next() <= 0) OR (StopWorkToDo));

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@WorkToDo';
            SQLParameter.SqlDbType := SQLDBType.NVarChar;
            SQLParameter.Size := 1024;
            SQLParameter.Value := COPYSTR(WorkToDo, 1, 1024);
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@EquipmentNo';
            SQLParameter.SqlDbType := SQLDBType.NVarChar;
            SQLParameter.Size := 255;
            SQLParameter.Value := ServiceItemLine."Service Item No.";
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@OrderTypeID';
            SQLParameter.SqlDbType := SQLDBType.NVarChar;
            SQLParameter.Size := 255;
            SQLParameter.Value := ServiceHeader."Service Order Type";
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@UserGroupID';
            SQLParameter.SqlDbType := SQLDBType.NVarChar;
            SQLParameter.Size := 255;
            SQLParameter.Value := ServiceHeader."Location Code";
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@Contactname';
            SQLParameter.SqlDbType := SQLDBType.NVarChar;
            SQLParameter.Size := 255;
            SQLParameter.Value := ServiceHeader."Contact Name";
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@Reference';
            SQLParameter.SqlDbType := SQLDBType.NVarChar;
            SQLParameter.Size := 255;
            SQLParameter.Value := ServiceHeader."Your Reference";
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@RepairStatus';
            SQLParameter.SqlDbType := SQLDBType.NVarChar;
            SQLParameter.Size := 255;
            SQLParameter.Value := ServiceItemLine."Repair Status Code";
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@DispatcherExternalId';
            SQLParameter.SqlDbType := SQLDBType.NVarChar;
            SQLParameter.Size := 255;
            SQLParameter.Value := ServiceHeader."Salesperson Code";
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@PriorityId';
            SQLParameter.SqlDbType := SQLDBType.NVarChar;
            SQLParameter.Size := 255;
            SQLParameter.Value := FORMAT(ServiceHeader.Priority);
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@CauseCodeId';
            SQLParameter.SqlDbType := SQLDBType.NVarChar;
            SQLParameter.Size := 255;
            SQLParameter.Value := "Reason Code";
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@ReceivedDate';
            SQLParameter.SqlDbType := SQLDBType.DateTime;
            //SQLParameter.Size := 255;
            SQLParameter.Value := ReturnLocalDateTime("Order Date", "Order Time");
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@ResponsibilityCenter';
            SQLParameter.SqlDbType := SQLDBType.NVarChar;
            SQLParameter.Size := 255;
            SQLParameter.Value := "Responsibility Center";
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@Deleted';
            SQLParameter.SqlDbType := SQLDBType.Bit;
            if Deleted then
                SQLParameter.Value := TRUE
            else
                SQLParameter.Value := False;
            SQLCommand.Parameters.Add(SQLParameter);

            SQLCommand.ExecuteNonQuery();

            Clear(SQLParameter);
            Clear(SQLCommand);
            SQLConnection.Close();
            Clear(SQLConnection);

            ServiceLine.Reset();
            ServiceLine.SetRange("Document Type", ServiceHeader."Document Type");
            ServiceLine.SetRange("Document No.", ServiceHeader."No.");
            ServiceLine.SetRange(Type, ServiceLine.Type::Item);
            ServiceLine.SetRange("Work Type Code", '');
            if ServiceLine.FindSet() then
                Repeat
                    PrecomRecordRef.GetTable(ServiceLine);
                    PreComUpdateQueue.Init();
                    PreComUpdateQueue."Update Message ID" := 999999999;
                    PreComUpdateQueue.RecordID := PrecomRecordRef.RECORDID;
                    PreComUpdateQueue."Command Type" := PreComUpdateQueue."Command Type"::Insert;
                    PreComUpdateQueue."Table ID" := PrecomRecordRef.NUMBER;
                    WriteServiceInvLine(PreComUpdateQueue);
                Until (ServiceLine.Next() <= 0);
        end;
    end;

    procedure WriteServiceInvLine(PreComUpdateQueue: Record "PreCom Update Queue")
    var
        ServiceLine: Record "Service Line";
        ServiceLine2: Record "Service Line";
        ServiceHeader: Record "Service Header";
        PrecomOrderNumberLink: Record "Precom Order Number Link";
        PrecomRecordRef: RecordRef;
        FieldRef1: FieldRef;
        FieldRef2: FieldRef;
        FieldRef3: FieldRef;
        SQLConnection: DotNet NewSqlConnection;
        SQLCommand: DotNet NewSqlCommand;
        SQLParameter: DotNet NewSqlParameter;
        SQLCommandType: DotNet CommandType;
        SQLDBType: DotNet NewSqlDbType;
        TestDT: Text;
        LongDescription: Text;
    begin
        if PreComUpdateQueue."Command Type" <> PreComUpdateQueue."Command Type"::Delete then begin
            PrecomRecordRef.Get(PreComUpdateQueue.RecordID);
            FieldRef1 := PrecomRecordRef.Field(1);
            FieldRef2 := PrecomRecordRef.Field(3);
            FieldRef3 := PrecomRecordRef.Field(4);
            ServiceLine.Get(FieldRef1.Value, FieldRef2.Value, FieldRef3.Value);
            if not PrecomOrderNumberLink.Get(ServiceLine."Document No.") then
                Clear(PrecomOrderNumberLink);
            WITH ServiceLine DO begin
                if IsNull(SQLConnection) then
                    SQLConnection := SQLConnection.SqlConnection();

                SQLConnection.ConnectionString(ReturnConnString());
                SQLConnection.Open();

                if IsNull(SQLCommand) then
                    SQLCommand := SQLCommand.SqlCommand();
                SQLCommand.Connection(SQLConnection);
                SQLCommand.CommandType(SQLCommandType.StoredProcedure);
                SQLCommand.CommandText('WriteArticleIn');

                ServiceHeader.Get("Document Type", "Document No.");

                SQLParameter := SQLParameter.SqlParameter();
                SQLParameter.ParameterName := '@OrderNumber';
                SQLParameter.SqlDbType := SQLDBType.Int;
                SQLParameter.Size := 255;
                SQLParameter.Value := PrecomOrderNumberLink."Precom No.";
                SQLCommand.Parameters.Add(SQLParameter);

                SQLParameter := SQLParameter.SqlParameter();
                SQLParameter.ParameterName := '@ERPReference';
                SQLParameter.SqlDbType := SQLDBType.NVarChar;
                SQLParameter.Size := 255;
                SQLParameter.Value := "Document No.";
                SQLCommand.Parameters.Add(SQLParameter);

                SQLParameter := SQLParameter.SqlParameter();
                SQLParameter.ParameterName := '@ArticleNumber';
                SQLParameter.SqlDbType := SQLDBType.NVarChar;
                SQLParameter.Size := 255;
                SQLParameter.Value := "No.";
                SQLCommand.Parameters.Add(SQLParameter);

                LongDescription := Description;
                ServiceLine2.Reset();
                ServiceLine2.SetRange("Document Type", ServiceLine."Document Type");
                ServiceLine2.SetRange("Document No.", ServiceLine."Document No.");
                ServiceLine2.SetRange("Attached to Line No.", ServiceLine."Line No.");
                if ServiceLine2.FindSet() then
                    Repeat
                        LongDescription += ' ' + ServiceLine2.Description;
                    Until (ServiceLine2.Next() <= 0);

                SQLParameter := SQLParameter.SqlParameter();
                SQLParameter.ParameterName := '@Description';
                SQLParameter.SqlDbType := SQLDBType.NVarChar;
                SQLParameter.Size := 255;
                SQLParameter.Value := LongDescription;
                SQLCommand.Parameters.Add(SQLParameter);

                SQLParameter := SQLParameter.SqlParameter();
                SQLParameter.ParameterName := '@Quantity';
                SQLParameter.SqlDbType := SQLDBType.Float;
                SQLParameter.Value := Quantity;
                SQLCommand.Parameters.Add(SQLParameter);

                SQLParameter := SQLParameter.SqlParameter();
                SQLParameter.ParameterName := '@UOM';
                SQLParameter.SqlDbType := SQLDBType.NVarChar;
                SQLParameter.Size := 255;
                SQLParameter.Value := "Unit of Measure Code";
                SQLCommand.Parameters.Add(SQLParameter);

                SQLParameter := SQLParameter.SqlParameter();
                SQLParameter.ParameterName := '@Price';
                SQLParameter.SqlDbType := SQLDBType.Float;
                SQLParameter.Value := "Unit Price";
                SQLCommand.Parameters.Add(SQLParameter);

                SQLParameter := SQLParameter.SqlParameter();
                SQLParameter.ParameterName := '@RecordIDERP';
                SQLParameter.SqlDbType := SQLDBType.NVarChar;
                SQLParameter.Size := 255;
                SQLParameter.Value := Format("Line No.");
                SQLCommand.Parameters.Add(SQLParameter);

                SQLParameter := SQLParameter.SqlParameter();
                SQLParameter.ParameterName := '@RecordID';
                SQLParameter.SqlDbType := SQLDBType.Int;
                SQLParameter.Value := "Precom Record ID";
                SQLCommand.Parameters.Add(SQLParameter);

                SQLParameter := SQLParameter.SqlParameter();
                SQLParameter.ParameterName := '@Currency';
                SQLParameter.SqlDbType := SQLDBType.NVarChar;
                SQLParameter.Size := 255;
                SQLParameter.Value := "Currency Code";
                SQLCommand.Parameters.Add(SQLParameter);

                SQLParameter := SQLParameter.SqlParameter();
                SQLParameter.ParameterName := '@ResourceExternalID';
                SQLParameter.SqlDbType := SQLDBType.NVarChar;
                SQLParameter.Size := 255;
                SQLParameter.Value := "Location Code";
                SQLCommand.Parameters.Add(SQLParameter);

                SQLParameter := SQLParameter.SqlParameter();
                SQLParameter.ParameterName := '@CreatedDateTime';
                SQLParameter.SqlDbType := SQLDBType.DateTime;
                TestDT := FORMAT("Planned Delivery Date", 0, '<Year4>-<Month,2>-<Day,2>') + ' 00:00:00';
                SQLParameter.Value := FORMAT("Planned Delivery Date", 0, '<Year4>-<Month,2>-<Day,2>') + ' 00:00:00';
                SQLCommand.Parameters.Add(SQLParameter);

                SQLParameter := SQLParameter.SqlParameter();
                SQLParameter.ParameterName := '@Picked';
                SQLParameter.SqlDbType := SQLDBType.Bit;
                if Quantity = "Quantity Shipped" then
                    SQLParameter.Value := TRUE
                else
                    SQLParameter.Value := False;
                SQLCommand.Parameters.Add(SQLParameter);

                SQLParameter := SQLParameter.SqlParameter();
                SQLParameter.ParameterName := '@Deleted';
                SQLParameter.SqlDbType := SQLDBType.Bit;

                if PreComUpdateQueue."Command Type" = PreComUpdateQueue."Command Type"::Delete then
                    SQLParameter.Value := TRUE
                else
                    SQLParameter.Value := False;
                SQLCommand.Parameters.Add(SQLParameter);

                SQLCommand.ExecuteNonQuery();

                Clear(SQLParameter);
                Clear(SQLCommand);
                SQLConnection.Close();
                Clear(SQLConnection);
            end;
        end else begin
            if IsNull(SQLConnection) then
                SQLConnection := SQLConnection.SqlConnection();

            SQLConnection.ConnectionString(ReturnConnString());
            SQLConnection.Open();

            if IsNull(SQLCommand) then
                SQLCommand := SQLCommand.SqlCommand();
            SQLCommand.Connection(SQLConnection);
            SQLCommand.CommandType(SQLCommandType.StoredProcedure);
            SQLCommand.CommandText('WriteArticleIn');

            if not PrecomOrderNumberLink.Get(GlobalServiceLine."Document No.") then
                Clear(PrecomOrderNumberLink);

            WITH GlobalServiceLine DO begin
                ServiceHeader.Get("Document Type", "Document No.");

                SQLParameter := SQLParameter.SqlParameter();
                SQLParameter.ParameterName := '@OrderNumber';
                SQLParameter.SqlDbType := SQLDBType.Int;
                SQLParameter.Size := 255;
                SQLParameter.Value := PrecomOrderNumberLink."Precom No.";
                SQLCommand.Parameters.Add(SQLParameter);

                SQLParameter := SQLParameter.SqlParameter();
                SQLParameter.ParameterName := '@ERPReference';
                SQLParameter.SqlDbType := SQLDBType.NVarChar;
                SQLParameter.Size := 255;
                SQLParameter.Value := "Document No.";
                SQLCommand.Parameters.Add(SQLParameter);

                SQLParameter := SQLParameter.SqlParameter();
                SQLParameter.ParameterName := '@ArticleNumber';
                SQLParameter.SqlDbType := SQLDBType.NVarChar;
                SQLParameter.Size := 255;
                SQLParameter.Value := "No.";
                SQLCommand.Parameters.Add(SQLParameter);

                SQLParameter := SQLParameter.SqlParameter();
                SQLParameter.ParameterName := '@Description';
                SQLParameter.SqlDbType := SQLDBType.NVarChar;
                SQLParameter.Size := 255;
                SQLParameter.Value := LongDescription;
                SQLCommand.Parameters.Add(SQLParameter);

                SQLParameter := SQLParameter.SqlParameter();
                SQLParameter.ParameterName := '@Quantity';
                SQLParameter.SqlDbType := SQLDBType.Float;
                SQLParameter.Value := Quantity;
                SQLCommand.Parameters.Add(SQLParameter);

                SQLParameter := SQLParameter.SqlParameter();
                SQLParameter.ParameterName := '@UOM';
                SQLParameter.SqlDbType := SQLDBType.NVarChar;
                SQLParameter.Size := 255;
                SQLParameter.Value := "Unit of Measure Code";
                SQLCommand.Parameters.Add(SQLParameter);

                SQLParameter := SQLParameter.SqlParameter();
                SQLParameter.ParameterName := '@Price';
                SQLParameter.SqlDbType := SQLDBType.Float;
                SQLParameter.Value := "Unit Price";
                SQLCommand.Parameters.Add(SQLParameter);

                SQLParameter := SQLParameter.SqlParameter();
                SQLParameter.ParameterName := '@RecordIDERP';
                SQLParameter.SqlDbType := SQLDBType.NVarChar;
                SQLParameter.Size := 255;
                SQLParameter.Value := Format("Line No.");
                SQLCommand.Parameters.Add(SQLParameter);

                SQLParameter := SQLParameter.SqlParameter();
                SQLParameter.ParameterName := '@RecordID';
                SQLParameter.SqlDbType := SQLDBType.Int;
                SQLParameter.Value := "Precom Record ID";
                SQLCommand.Parameters.Add(SQLParameter);

                SQLParameter := SQLParameter.SqlParameter();
                SQLParameter.ParameterName := '@Currency';
                SQLParameter.SqlDbType := SQLDBType.NVarChar;
                SQLParameter.Size := 255;
                SQLParameter.Value := "Currency Code";
                SQLCommand.Parameters.Add(SQLParameter);

                SQLParameter := SQLParameter.SqlParameter();
                SQLParameter.ParameterName := '@ResourceExternalID';
                SQLParameter.SqlDbType := SQLDBType.NVarChar;
                SQLParameter.Size := 255;
                SQLParameter.Value := "Location Code";
                SQLCommand.Parameters.Add(SQLParameter);

                SQLParameter := SQLParameter.SqlParameter();
                SQLParameter.ParameterName := '@CreatedDateTime';
                SQLParameter.SqlDbType := SQLDBType.DateTime;
                TestDT := FORMAT("Planned Delivery Date", 0, '<Year4>-<Month,2>-<Day,2>') + ' 00:00:00';
                SQLParameter.Value := FORMAT("Planned Delivery Date", 0, '<Year4>-<Month,2>-<Day,2>') + ' 00:00:00';
                SQLCommand.Parameters.Add(SQLParameter);

                SQLParameter := SQLParameter.SqlParameter();
                SQLParameter.ParameterName := '@Picked';
                SQLParameter.SqlDbType := SQLDBType.Bit;
                if Quantity = "Quantity Shipped" then
                    SQLParameter.Value := TRUE
                else
                    SQLParameter.Value := False;
                SQLCommand.Parameters.Add(SQLParameter);

                SQLParameter := SQLParameter.SqlParameter();
                SQLParameter.ParameterName := '@Deleted';
                SQLParameter.SqlDbType := SQLDBType.Bit;

                if PreComUpdateQueue."Command Type" = PreComUpdateQueue."Command Type"::Delete then
                    SQLParameter.Value := TRUE
                else
                    SQLParameter.Value := False;
                SQLCommand.Parameters.Add(SQLParameter);

                SQLCommand.ExecuteNonQuery();

                Clear(SQLParameter);
                Clear(SQLCommand);
                SQLConnection.Close();
                Clear(SQLConnection);
            end;
        end;
    end;

    procedure WriteServiceInvLineHistory(PreComUpdateQueue: Record "PreCom Update Queue")
    var
        ServiceLine: Record "Service Line";
        ServiceLine2: Record "Service Line";
        ServiceHeader: Record "Service Header";
        PrecomOrderNumberLink: Record "Precom Order Number Link";
        PrecomRecordRef: RecordRef;
        FieldRef1: FieldRef;
        FieldRef2: FieldRef;
        FieldRef3: FieldRef;
        SQLConnection: DotNet NewSqlConnection;
        SQLCommand: DotNet NewSqlCommand;
        SQLParameter: DotNet NewSqlParameter;
        SQLCommandType: DotNet CommandType;
        SQLDBType: DotNet NewSqlDbType;
        TestDT: Text;
        LongDescription: Text;
    begin

        PrecomRecordRef.Get(PreComUpdateQueue.RecordID);
        FieldRef1 := PrecomRecordRef.Field(1);
        FieldRef2 := PrecomRecordRef.Field(3);
        FieldRef3 := PrecomRecordRef.Field(4);
        ServiceLine.Get(FieldRef1.Value, FieldRef2.Value, FieldRef3.Value);
        if not PrecomOrderNumberLink.Get(ServiceLine."Document No.") then
            Clear(PrecomOrderNumberLink);
        with ServiceLine do begin
            if IsNull(SQLConnection) then
                SQLConnection := SQLConnection.SqlConnection();

            SQLConnection.ConnectionString(ReturnConnString());
            SQLConnection.Open();

            if IsNull(SQLCommand) then
                SQLCommand := SQLCommand.SqlCommand();
            SQLCommand.Connection(SQLConnection);
            SQLCommand.CommandType(SQLCommandType.StoredProcedure);
            SQLCommand.CommandText('WriteArticleInHistory');

            ServiceHeader.Get("Document Type", "Document No.");

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@OrderNumber';
            SQLParameter.SqlDbType := SQLDBType.Int;
            SQLParameter.Size := 255;
            SQLParameter.Value := PrecomOrderNumberLink."Precom No.";
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@ERPReference';
            SQLParameter.SqlDbType := SQLDBType.NVarChar;
            SQLParameter.Size := 255;
            SQLParameter.Value := "Document No.";
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@ArticleNumber';
            SQLParameter.SqlDbType := SQLDBType.NVarChar;
            SQLParameter.Size := 255;
            SQLParameter.Value := "No.";
            SQLCommand.Parameters.Add(SQLParameter);

            LongDescription := Description;
            ServiceLine2.Reset();
            ServiceLine2.SetRange("Document Type", ServiceLine."Document Type");
            ServiceLine2.SetRange("Document No.", ServiceLine."Document No.");
            ServiceLine2.SetRange("Attached to Line No.", ServiceLine."Line No.");
            if ServiceLine2.FindSet() then
                Repeat
                    LongDescription += ' ' + ServiceLine2.Description;
                Until (ServiceLine2.Next() <= 0);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@Description';
            SQLParameter.SqlDbType := SQLDBType.NVarChar;
            SQLParameter.Size := 255;
            SQLParameter.Value := LongDescription;
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@Quantity';
            SQLParameter.SqlDbType := SQLDBType.Float;
            SQLParameter.Value := Quantity;
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@UOM';
            SQLParameter.SqlDbType := SQLDBType.NVarChar;
            SQLParameter.Size := 255;
            SQLParameter.Value := "Unit of Measure Code";
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@Price';
            SQLParameter.SqlDbType := SQLDBType.Float;
            SQLParameter.Value := "Unit Price";
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@RecordIDERP';
            SQLParameter.SqlDbType := SQLDBType.NVarChar;
            SQLParameter.Size := 255;
            SQLParameter.Value := Format("Line No.");
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@RecordID';
            SQLParameter.SqlDbType := SQLDBType.Int;
            SQLParameter.Value := "Precom Record ID";
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@Currency';
            SQLParameter.SqlDbType := SQLDBType.NVarChar;
            SQLParameter.Size := 255;
            SQLParameter.Value := "Currency Code";
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@ResourceExternalID';
            SQLParameter.SqlDbType := SQLDBType.NVarChar;
            SQLParameter.Size := 255;
            SQLParameter.Value := "Location Code";
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@CreatedDateTime';
            SQLParameter.SqlDbType := SQLDBType.DateTime;
            TestDT := FORMAT("Planned Delivery Date", 0, '<Year4>-<Month,2>-<Day,2>') + ' 00:00:00';
            SQLParameter.Value := FORMAT("Planned Delivery Date", 0, '<Year4>-<Month,2>-<Day,2>') + ' 00:00:00';
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@Picked';
            SQLParameter.SqlDbType := SQLDBType.Bit;
            if Quantity = "Quantity Shipped" then
                SQLParameter.Value := true
            else
                SQLParameter.Value := False;
            SQLCommand.Parameters.Add(SQLParameter);

            SQLParameter := SQLParameter.SqlParameter();
            SQLParameter.ParameterName := '@Deleted';
            SQLParameter.SqlDbType := SQLDBType.Bit;

            if PreComUpdateQueue."Command Type" = PreComUpdateQueue."Command Type"::Delete then
                SQLParameter.Value := true
            else
                SQLParameter.Value := false;
            SQLCommand.Parameters.Add(SQLParameter);

            SQLCommand.ExecuteNonQuery();

            Clear(SQLParameter);
            Clear(SQLCommand);
            SQLConnection.Close();
            Clear(SQLConnection);
        end;
    end;

    procedure ImportServiceOrders(PreComUpdateQueue: Record "PreCom Update Queue")
    var
        ServiceHeader: Record "Service Header";
        ServiceItemLine: Record "Service Item Line";
        ServiceLine: Record "Service Line";
        PreComUpdateSetup: Record "PreCom Update Setup";
        PrecomOrderNumberLink: Record "Precom Order Number Link";
        ServiceItem: Record "Service Item";
        Location: Record Location;
        SalesPerson: Record "Salesperson/Purchaser";
        GenLedgSetup: Record "General Ledger Setup";
        ReasonCode: Record "Reason Code";
        ServiceInvoiceHeader: Record "Service Invoice Header";
        GLSetup: Record "General Ledger Setup";
        Customer: Record Customer;
        RepairStatus: Record "Repair Status";
        ServiceContractLine: Record "Service Contract Line";
        SQLConnection: DotNet NewSqlConnection;
        SQLCommand: DotNet NewSqlCommand;
        SQLParameter: DotNet NewSqlParameter;
        SQLCommandType: DotNet CommandType;
        SQLDBType: DotNet NewSqlDbType;
        ImportDate: Date;
        DescriptionText: Text[1024];
        DeleteServiceOrder: Boolean;
        NoCustomerNoErr: Label 'Customer No. not entered!';
        LineNo: Integer;
        i: Integer;
        InsertWorkDoneLines: Boolean;
        WorkToDoLineInterval: Integer;
        UseEquipmentNumber: Code[20];
        ServiceDescriptionErr: Label 'Service Article %1 has more than one Service Contract!', Comment = '%1 = Service Article No.';

    begin
        GLSetup.Get();

        if IsNull(SQLConnection) then
            SQLConnection := SQLConnection.SqlConnection();

        SQLConnection.ConnectionString(ReturnConnString());
        SQLConnection.Open();

        DeleteServiceOrder := PreComUpdateQueue.Deleted;
        if (not DeleteServiceOrder) AND (PreComUpdateQueue.CustomerNumber <> '') then begin
            if not ServiceHeader.Get(ServiceHeader."Document Type"::Order, PreComUpdateQueue.ERPReference) then begin
                PrecomOrderNumberLink.Reset();
                PrecomOrderNumberLink.SetRange("Precom No.", PreComUpdateQueue.OrderNumber);
                if not PrecomOrderNumberLink.FindFirst() then
                    Clear(PrecomOrderNumberLink);

                if not ServiceHeader.Get(ServiceHeader."Document Type"::Order, PrecomOrderNumberLink."No.") then
                    Clear(ServiceHeader);

                if ServiceHeader."No." = '' then begin
                    if ServiceInvoiceHeader.Get(PreComUpdateQueue.ERPReference) then
                        Exit;

                    if PreComUpdateQueue.ERPReference <> '' then begin
                        ServiceInvoiceHeader.Reset();
                        ServiceInvoiceHeader.SetCurrentKey("Order No.");
                        ServiceInvoiceHeader.SetRange("Order No.", PreComUpdateQueue.ERPReference);
                        if not ServiceInvoiceHeader.IsEmpty() then
                            Exit;
                    end;
                end;
            end;

            if ServiceHeader."No." = '' then begin
                UseEquipmentNumber := CopyStr(PreComUpdateQueue.EquipmentNumber, 1, 20);
                ImportDate := ConvertDate(PreComUpdateQueue.PlannedStartDate);
                if ImportDate = 19000101D then
                    ImportDate := Today();

                ServiceContractLine.Reset();
                ServiceContractLine.SetRange("Contract Type", ServiceContractLine."Contract Type"::Contract);
                ServiceContractLine.SetRange("Service Item No.", UseEquipmentNumber);
                ServiceContractLine.SetRange("Contract Status", ServiceContractLine."Contract Status"::Signed);
                ServiceContractLine.SetFilter("Starting Date", '<=%1', ImportDate);
                ServiceContractLine.SetFilter("Contract Expiration Date", '>%1 | =%2', ImportDate, 0D);
                IF ServiceContractLine.Count() > 1 then
                    Error(ServiceDescriptionErr);

                ServiceHeader.SetHideValidationDialog(true);
                ServiceHeader.Init();
                ServiceHeader.Validate("Document Type", ServiceHeader."Document Type"::Order);
                ServiceHeader."No." := '';
                ServiceHeader.Insert(TRUE);
                ServiceHeader."Starting Date" := 0D;
                ServiceHeader."Starting Time" := 0T;
                ServiceHeader."Finishing Date" := 0D;
                ServiceHeader."Finishing Time" := 0T;
                ServiceHeader.SetHideValidationDialog(TRUE);
                ServiceHeader.Validate("Customer No.", PreComUpdateQueue.CustomerNumber);
                if (PreComUpdateQueue.BillingNumber <> '') AND (PreComUpdateQueue.BillingNumber <> PreComUpdateQueue.CustomerNumber) then
                    ServiceHeader.Validate("Bill-to Customer No.", PreComUpdateQueue.BillingNumber);
                //    if PreComUpdateQueue.PrimaryResource <> '' then
                //      if STRLEN(PreComUpdateQueue.PrimaryResource) <= 10 then
                //        if Location.Get(PreComUpdateQueue.PrimaryResource) then
                //          ServiceHeader.Validate("Location Code",PreComUpdateQueue.PrimaryResource);

                if GLSetup."Bill-to/Sell-to VAT Calc." = GLSetup."Bill-to/Sell-to VAT Calc."::"Sell-to/Buy-from No." then begin
                    Customer.Get(ServiceHeader."Customer No.");
                    if ServiceHeader."VAT Bus. Posting Group" <> Customer."VAT Bus. Posting Group" then
                        ServiceHeader.Validate("VAT Bus. Posting Group", Customer."VAT Bus. Posting Group");
                end else begin
                    Customer.Get(ServiceHeader."Bill-to Customer No.");
                    if ServiceHeader."VAT Bus. Posting Group" <> Customer."VAT Bus. Posting Group" then
                        ServiceHeader.Validate("VAT Bus. Posting Group", Customer."VAT Bus. Posting Group");
                end;

                ImportDate := ConvertDate(PreComUpdateQueue.ActualStartDate);
                if ImportDate <> 19000101D then
                    ServiceHeader.Validate("Posting Date", ImportDate);

                ImportDate := ConvertDate(PreComUpdateQueue.PlannedStartDate);
                if ImportDate <> 19000101D then begin
                    if ServiceHeader."Order Date" > ImportDate then
                        ServiceHeader.Validate("Order Date", ImportDate);
                    if ServiceHeader."Finishing Date" = 0D then
                        ServiceHeader.Validate("Starting Date", ImportDate);
                end;

                PrecomOrderNumberLink.Reset();
                PrecomOrderNumberLink.Init();
                PrecomOrderNumberLink.Validate("No.", ServiceHeader."No.");
                PrecomOrderNumberLink.Validate("Precom No.", PreComUpdateQueue.OrderNumber);
                PrecomOrderNumberLink.Insert(true);

                ServiceItemLine.Reset();

                PreComUpdateSetup.Get();
                if PreComUpdateSetup."Default Service Order Type" <> '' then
                    ServiceHeader.Validate("Service Order Type", PreComUpdateSetup."Default Service Order Type");
                if PreComUpdateSetup."Default Dimension 1 Code" <> '' then
                    ServiceHeader.Validate("Shortcut Dimension 1 Code", PreComUpdateSetup."Default Dimension 1 Code");
                ServiceHeader."Your Reference" := CopyStr(PreComUpdateQueue.Reference, 1, 35);
                if ReasonCode.Get(PreComUpdateQueue.CauseCode) then
                    ServiceHeader.Validate("Reason Code", PreComUpdateQueue.CauseCode);

                if SalesPerson.Get(PreComUpdateQueue.PrimaryResource) then
                    ServiceHeader.Validate("Salesperson Code", PreComUpdateQueue.PrimaryResource);
                GenLedgSetup.Get();
                if Location.Get(PreComUpdateQueue.PrimaryResource) then
                    ServiceHeader.Validate("Location Code", PreComUpdateQueue.PrimaryResource);
                if PreComUpdateQueue.Type <> '' then
                    ServiceHeader.Validate("Service Order Type", PreComUpdateQueue.Type);
                ServiceHeader.Validate("Responsibility Center", PreComUpdateQueue.ResponsibilityCenter);
                ServiceHeader.Validate("Service Order (Quick)", true);
                ServiceHeader.Validate(Description, PreComUpdateQueue.Description);

                ServiceHeader.Modify(true);
                ServiceHeader.Get(ServiceHeader."Document Type", ServiceHeader."No.");
                ServiceHeader.SetHideValidationDialog(true);

                ServiceItemLine.Init();
                ServiceItemLine.SetHideDialogBox(true);
                ServiceItemLine.Validate("Document Type", ServiceHeader."Document Type");
                ServiceItemLine.Validate("Document No.", ServiceHeader."No.");
                ServiceItemLine.Validate("Line No.", 10000);

                if not ServiceItem.Get(PreComUpdateQueue.EquipmentNumber) then begin
                    Clear(ServiceItem);
                    UseEquipmentNumber := '';
                end;
                if ServiceHeader."Customer No." <> ServiceItem."Customer No." then
                    UseEquipmentNumber := '';

                if UseEquipmentNumber <> '' then
                    ServiceItemLine.Validate("Service Item No.", UseEquipmentNumber)
                else
                    ServiceItemLine.Validate(Description, PreComUpdateQueue.Description);
                //      ServiceItemLine.Validate("Reason Code",PreComUpdateQueue.Type);
                if ServiceItemLine."Shortcut Dimension 1 Code" = '' then
                    if PreComUpdateSetup."Default Dimension 1 Code" <> '' then
                        ServiceItemLine.Validate("Shortcut Dimension 1 Code", PreComUpdateSetup."Default Dimension 1 Code");
                ServiceItemLine.Validate("Starting Date", ServiceHeader."Starting Date");
                //ServiceItemLine.Validate("Finishing Date",ServiceHeader."Finishing Date");
                if RepairStatus.Get(PreComUpdateQueue.RepairStatus) then
                    ServiceItemLine.Validate("Repair Status Code", PreComUpdateQueue.RepairStatus);
                ServiceItemLine.Insert(true);

                ServiceHeader.Get(ServiceHeader."Document Type", ServiceHeader."No.");
                ServiceHeader.SetHideValidationDialog(true);
                //    ServiceHeader."SA Item Description" := ServiceItemLine."SA Item Description";
                ServiceHeader."Your Reference" := CopyStr(PreComUpdateQueue.Reference, 1, 35);
                ServiceHeader.MODifY(TRUE);

                LineNo := 10000;
                //DescriptionText := PreComUpdateQueue.WorkToDo;
                DescriptionText := CopyStr(GlobalWorkToDo, 1, 1024);
                SplitStringToLines(DescriptionText);
                i := 1;
                FOR i := 1 TO 50 DO
                    if DescriptionLines[i] <> '' then begin
                        //        if i = 1 then begin
                        //          ServiceHeader.Validate(Description,DescriptionLines[i]);
                        //          ServiceHeader.MODifY(TRUE);
                        //        end else begin
                        ServiceLine.Reset();
                        ServiceLine.Init();
                        ServiceLine.Validate("Document Type", ServiceHeader."Document Type");
                        ServiceLine.Validate("Document No.", ServiceHeader."No.");
                        ServiceLine.Validate("Line No.", LineNo);
                        ServiceLine.Validate("Service Item Line No.", ServiceItemLine."Line No.");
                        LineNo += 10000;
                        ServiceLine.Validate(Type, ServiceLine.Type::" ");
                        ServiceLine.Description := DescriptionLines[i];
                        ServiceLine.Insert(TRUE);
                        //        end;
                    end;

                ServiceLine.Reset();
                ServiceLine.SetRange("Document Type", ServiceHeader."Document Type");
                ServiceLine.SetRange("Document No.", ServiceHeader."No.");
                if ServiceLine.FindLast() then
                    LineNo := ServiceLine."Line No." + 10000
                else
                    LineNo := 10000;

                DescriptionText := PreComUpdateQueue.WorkDoneExternal1 + ' ' + PreComUpdateQueue.WorkDoneExternal2 + ' ' + PreComUpdateQueue.WorkDoneExternal3;
                SplitStringToLines(DescriptionText);
                i := 1;
                FOR i := 1 TO 50 DO
                    if DescriptionLines[i] <> '' then begin
                        ServiceLine.Reset();
                        ServiceLine.Init();
                        ServiceLine.Validate("Document Type", ServiceHeader."Document Type");
                        ServiceLine.Validate("Document No.", ServiceHeader."No.");
                        ServiceLine.Validate("Line No.", LineNo);
                        ServiceLine.Validate("Service Item Line No.", ServiceItemLine."Line No.");
                        LineNo += 10000;
                        ServiceLine.Validate(Type, ServiceLine.Type::" ");
                        ServiceLine.Description := DescriptionLines[i];
                        ServiceLine.Insert(TRUE);
                    end;

                //    if (EVALUATE(ActualendDate,PreComUpdateQueue.ActualendDate)) AND (EVALUATE(ActualStartDate,PreComUpdateQueue.ActualStartDate)) then begin
                //      ServiceItemLine.Validate("Starting Date",DT2DATE(ActualStartDate));
                //      ServiceItemLine.Validate("Finishing Date",DT2DATE(ActualendDate));
                //    end;
                //    ServiceItemLine.MODifY(TRUE);

                if IsNull(SQLCommand) then
                    SQLCommand := SQLCommand.SqlCommand();
                SQLCommand.Connection(SQLConnection);
                SQLCommand.CommandType(SQLCommandType.StoredProcedure);
                SQLCommand.CommandText('WritePreComToNAVHandshake');

                SQLParameter := SQLParameter.SqlParameter();
                SQLParameter.ParameterName := '@OrderNumber';
                SQLParameter.SqlDbType := SQLDBType.NVarChar;
                SQLParameter.Size := 255;
                SQLParameter.Value := PreComUpdateQueue.OrderNumber;
                SQLCommand.Parameters.Add(SQLParameter);

                SQLParameter := SQLParameter.SqlParameter();
                SQLParameter.ParameterName := '@ERPReference';
                SQLParameter.SqlDbType := SQLDBType.NVarChar;
                SQLParameter.Size := 255;
                SQLParameter.Value := ServiceHeader."No.";
                SQLCommand.Parameters.Add(SQLParameter);

                SQLCommand.ExecuteNonQuery();

                Clear(SQLParameter);
                Clear(SQLCommand);
            end else begin
                ServiceHeader.SetHideValidationDialog(true);

                if (PreComUpdateQueue.BillingNumber <> '') AND (PreComUpdateQueue.BillingNumber <> PreComUpdateQueue.CustomerNumber) then
                    if ServiceHeader."Bill-to Customer No." <> PreComUpdateQueue.BillingNumber then
                        ServiceHeader.Validate("Bill-to Customer No.", PreComUpdateQueue.BillingNumber);

                //    if PreComUpdateQueue.PrimaryResource <> '' then
                //      if STRLEN(PreComUpdateQueue.PrimaryResource) <= 10 then
                //        if Location.Get(PreComUpdateQueue.PrimaryResource) then
                //          ServiceHeader.Validate("Location Code",PreComUpdateQueue.PrimaryResource);

                ImportDate := ConvertDate(PreComUpdateQueue.ActualStartDate);
                if ImportDate <> ServiceHeader."Posting Date" then begin
                    ServiceHeader.SetHideValidationDialog(true);
                    ServiceHeader.Validate("Posting Date", ImportDate);
                end;

                ImportDate := ConvertDate(PreComUpdateQueue.PlannedStartDate);
                if (ImportDate <> 19000101D) AND (ServiceHeader."Finishing Date" = 0D) then begin
                    if (ServiceHeader."Order Date" > ImportDate) then
                        ServiceHeader.Validate("Order Date", ImportDate);

                    ServItemLineStartingDateReset(ServiceHeader, False);
                    ServiceHeader.Validate("Starting Date", ImportDate);
                    ServItemLineStartingDateReset(ServiceHeader, TRUE);
                end;

                if GLSetup."Bill-to/Sell-to VAT Calc." = GLSetup."Bill-to/Sell-to VAT Calc."::"Sell-to/Buy-from No." then begin
                    Customer.Get(ServiceHeader."Customer No.");
                    if ServiceHeader."VAT Bus. Posting Group" <> Customer."VAT Bus. Posting Group" then
                        ServiceHeader.Validate("VAT Bus. Posting Group", Customer."VAT Bus. Posting Group");
                end else begin
                    Customer.Get(ServiceHeader."Bill-to Customer No.");
                    if ServiceHeader."VAT Bus. Posting Group" <> Customer."VAT Bus. Posting Group" then
                        ServiceHeader.Validate("VAT Bus. Posting Group", Customer."VAT Bus. Posting Group");
                end;

                /*
                ImportDate := ConvertDate(PreComUpdateQueue.PlannedendDate);
                if ImportDate <> 01011900D then begin
                  if ImportDate >= ServiceHeader."Starting Date" then
                    ServiceHeader.Validate("Finishing Date",ImportDate)
                  else
                    ServiceHeader.Validate("Finishing Date",ServiceHeader."Starting Date");
                end;
                */

                ServiceItemLine.Reset();
                ServiceItemLine.SetRange("Document Type", ServiceHeader."Document Type");
                ServiceItemLine.SetRange("Document No.", ServiceHeader."No.");
                if not ServiceItemLine.FindFirst() then
                    Clear(ServiceItemLine);

                //DescriptionText := PreComUpdateQueue.WorkToDo;
                DescriptionText := CopyStr(GlobalWorkToDo, 1, 1024);
                SplitStringToLines(DescriptionText);
                LineNo := RemoveWorkToDoLines(ServiceHeader);
                WorkToDoLineInterval := LineNo;
                i := 1;
                FOR i := 1 TO 50 DO
                    if DescriptionLines[i] <> '' then begin
                        //        if i = 1 then begin
                        //          ServiceHeader.Validate(Description,DescriptionLines[i]);
                        //          ServiceHeader.MODifY(TRUE);
                        //        end else begin
                        ServiceLine.Reset();
                        ServiceLine.HideShowDialog(true);
                        ServiceLine.Init();
                        ServiceLine.Validate("Document Type", ServiceHeader."Document Type");
                        ServiceLine.Validate("Document No.", ServiceHeader."No.");
                        ServiceLine.Validate("Line No.", LineNo);
                        LineNo += WorkToDoLineInterval;
                        ServiceLine.Validate("Service Item Line No.", ServiceItemLine."Line No.");
                        ServiceLine.Validate(Type, ServiceLine.Type::" ");
                        ServiceLine.Description := DescriptionLines[i];
                        ServiceLine.Insert(true);
                        //        end;
                    end;

                //    if STRLEN(PreComUpdateQueue.WorkToDo) > 50 then
                //      DescriptionText := COPYSTR(PreComUpdateQueue.WorkToDo,1,50)
                //    else
                //      DescriptionText := PreComUpdateQueue.WorkToDo;
                //    ServiceHeader.Validate(Description,DescriptionText);

                ServiceItemLine.Reset();

                PreComUpdateSetup.Get();
                if (PreComUpdateSetup."Default Service Order Type" <> '') AND (ServiceHeader."Service Order Type" = '') then
                    ServiceHeader.Validate("Service Order Type", PreComUpdateSetup."Default Service Order Type");

                if (PreComUpdateQueue.Type <> '') AND (PreComUpdateQueue.Type <> ServiceHeader."Service Order Type") then
                    ServiceHeader.Validate("Service Order Type", PreComUpdateQueue.Type);

                if not PrecomOrderNumberLink.Get(ServiceHeader."No.") then begin
                    PrecomOrderNumberLink.Reset();
                    PrecomOrderNumberLink.Init();
                    PrecomOrderNumberLink.Validate("No.", ServiceHeader."No.");
                    PrecomOrderNumberLink.Validate("Precom No.", PreComUpdateQueue.OrderNumber);
                    PrecomOrderNumberLink.Insert(true);
                end;


                if ReasonCode.Get(PreComUpdateQueue.CauseCode) then
                    ServiceHeader.Validate("Reason Code", PreComUpdateQueue.CauseCode);

                if PreComUpdateQueue.Reference <> '' then
                    ServiceHeader."Your Reference" := CopyStr(PreComUpdateQueue.Reference, 1, 35);

                if (SalesPerson.Get(PreComUpdateQueue.PrimaryResource)) AND (ServiceHeader."Salesperson Code" = '') then
                    ServiceHeader.Validate("Salesperson Code", PreComUpdateQueue.PrimaryResource);
                GenLedgSetup.Get();
                if (Location.Get(PreComUpdateQueue.PrimaryResource)) AND (ServiceHeader."Location Code" = '') then
                    ServiceHeader.Validate("Location Code", PreComUpdateQueue.PrimaryResource);

                if PreComUpdateQueue.Reference <> '' then
                    ServiceHeader."Your Reference" := CopyStr(PreComUpdateQueue.Reference, 1, 35);
                if ServiceHeader."Responsibility Center" = '' then
                    ServiceHeader.Validate("Responsibility Center", PreComUpdateQueue.ResponsibilityCenter);
                if ServiceHeader.Description <> PreComUpdateQueue.Description then
                    ServiceHeader.Validate(Description, PreComUpdateQueue.Description);
                ServiceHeader.Modify(true);

                ServiceItemLine.Reset();
                ServiceItemLine.SetRange("Document Type", ServiceHeader."Document Type");
                ServiceItemLine.SetRange("Document No.", ServiceHeader."No.");
                if ServiceItemLine.FindFirst() then begin
                    UseEquipmentNumber := CopyStr(PreComUpdateQueue.EquipmentNumber, 1, 20);
                    if not ServiceItem.Get(PreComUpdateQueue.EquipmentNumber) then begin
                        Clear(ServiceItem);
                        UseEquipmentNumber := '';
                    end;
                    if ServiceHeader."Customer No." <> ServiceItem."Customer No." then
                        UseEquipmentNumber := '';
                    if (ServiceItemLine."Service Item No." <> UseEquipmentNumber) AND (UseEquipmentNumber <> '') then begin
                        ServiceItemLine.Delete(true);
                        ServiceItemLine.Reset();
                        ServiceItemLine.SetHideDialogBox(true);
                        ServiceItemLine.Init();
                        ServiceItemLine.Validate("Document Type", ServiceHeader."Document Type");
                        ServiceItemLine.Validate("Document No.", ServiceHeader."No.");
                        ServiceItemLine.Validate("Line No.", 10000);
                        ServiceItemLine.Validate("Service Item No.", PreComUpdateQueue.EquipmentNumber);
                        ServiceItemLine.Validate("Starting Date", ServiceHeader."Starting Date");
                        //ServiceItemLine.Validate("Finishing Date",ServiceHeader."Finishing Date");
                        ServiceItemLine.Insert(true);
                        //end else begin
                        //ServiceItemLine.Validate("Reason Code",PreComUpdateQueue.Type);
                        //ServiceItemLine.MODifY(TRUE);
                    end;
                    if RepairStatus.Get(PreComUpdateQueue.RepairStatus) then
                        ServiceItemLine.Validate("Repair Status Code", PreComUpdateQueue.RepairStatus);
                    ServiceItemLine.Modify(true);
                end;

                ServiceLine.Reset();
                ServiceLine.SetRange("Document Type", ServiceHeader."Document Type");
                ServiceLine.SetRange("Document No.", ServiceHeader."No.");
                if ServiceLine.FindLast() then
                    LineNo := ServiceLine."Line No." + 10000
                else
                    LineNo := 10000;

                DescriptionText := PreComUpdateQueue.WorkDoneExternal1 + ' ' + PreComUpdateQueue.WorkDoneExternal2 + ' ' + PreComUpdateQueue.WorkDoneExternal3;
                SplitStringToLines(DescriptionText);
                i := 1;
                InsertWorkDoneLines := TRUE;
                if ServiceLine.FindFirst() then
                    if ServiceLine.Description = DescriptionLines[i] then
                        InsertWorkDoneLines := False;

                if InsertWorkDoneLines then
                    FOR i := 1 TO 50 DO
                        if DescriptionLines[i] <> '' then begin
                            ServiceLine.Reset();
                            ServiceLine.HideShowDialog(true);
                            ServiceLine.Init();
                            ServiceLine.Validate("Document Type", ServiceHeader."Document Type");
                            ServiceLine.Validate("Document No.", ServiceHeader."No.");
                            ServiceLine.Validate("Line No.", LineNo);
                            ServiceLine.Validate("Service Item Line No.", ServiceItemLine."Line No.");
                            LineNo += 10000;
                            ServiceLine.Validate(Type, ServiceLine.Type::" ");
                            ServiceLine.Description := DescriptionLines[i];
                            ServiceLine.Insert(true);
                        end;
            end;
        end else
            if DeleteServiceOrder then begin
                if ServiceHeader.Get(ServiceHeader."Document Type"::Order, PreComUpdateQueue.ERPReference) then begin
                    ServiceHeader.SetHideValidationDialog(true);
                    ServiceHeader.Delete(true);
                end;
            end else
                if PreComUpdateQueue.CustomerNumber = '' then
                    Error(NoCustomerNoErr);

        SQLConnection.Close();
        Clear(SQLConnection);

    end;

    procedure ImportInvoiceInfo(PreComUpdateQueue: Record "PreCom Update Queue")
    var
        ServiceHeader: Record "Service Header";
        ServiceItemLine: Record "Service Item Line";
        ServiceItem: Record "Service Item";
        RepairStatus: Record "Repair Status";
        //ReasonCode: Record "Reason Code";
        PreComUpdateSetup: Record "PreCom Update Setup";
        //TempNameValueBuffer: Record "Name/Value Buffer" temporary;
        ServiceLine: Record "Service Line";
        StandardText: Record "Standard Text";
        //FileManagement: Codeunit "File Management";
        TransferExtendedText: Codeunit "Transfer Extended Text";
        HourReading: Integer;
        LineNo: Integer;
    //ServiceHeaderChanged: Boolean;
    begin
        //ServiceHeaderChanged := False;
        ServiceHeader.Reset();
        if ServiceHeader.Get(ServiceHeader."Document Type"::Order, PreComUpdateQueue.ERPReference) then begin
            if not PreComUpdateSetup.Get() then
                Clear(PreComUpdateSetup);

            // TODO: Skal det med i standard løsning?
            /*if FileManagement.ServerDirectoryExists(PreComUpdateSetup."Document Path" + '\' + PreComUpdateQueue.ERPReference) then begin
                FileManagement.OpenZipFile(NameValueBuffer, PreComUpdateSetup."Document Path" + '\' + PreComUpdateQueue.ERPReference);
                if NameValueBuffer.FindSet then
                    Repeat
                        ServiceHeader.ADDLINK(NameValueBuffer.Name);
                    Until (NameValueBuffer.Next = 0);
            end;*/

            ServiceItemLine.Reset();
            ServiceItemLine.SetRange("Document Type", ServiceHeader."Document Type");
            ServiceItemLine.SetRange("Document No.", ServiceHeader."No.");
            if ServiceItemLine.FindFirst() then begin
                if not ServiceItem.Get(ServiceItemLine."Service Item No.") then
                    Clear(ServiceItem);

                if Evaluate(HourReading, PreComUpdateQueue.EquipmentReading) then
                    if HourReading > 0 then begin
                        ServiceItemLine.Validate(Counter, HourReading);
                        ServiceItemLine.Modify(true);
                    end;
            end;

            if PreComUpdateQueue.StatusID = '160' then begin
                ServiceHeader.Get(ServiceHeader."Document Type"::Order, PreComUpdateQueue.ERPReference);
                ServiceHeader.Validate(Status, ServiceHeader.Status::Finished);
                ServiceHeader.Modify(true);

                SortServiceLines(ServiceHeader);

                if StandardText.Get(ServiceHeader."Salesperson Code") then begin
                    ServiceLine.Reset();
                    ServiceLine.SetRange("Document Type", ServiceHeader."Document Type");
                    ServiceLine.SetRange("Document No.", ServiceHeader."No.");
                    if ServiceLine.FindLast() then
                        LineNo := ServiceLine."Line No." + 10000
                    else
                        LineNo := 10000;
                    ServiceLine.Reset();
                    ServiceLine.Init();
                    ServiceLine.Validate("Document Type", ServiceHeader."Document Type");
                    ServiceLine.Validate("Document No.", ServiceHeader."No.");
                    ServiceLine.Validate("Line No.", LineNo);
                    ServiceLine.Validate("Service Item Line No.", ServiceItemLine."Line No.");
                    ServiceLine.Validate(Type, ServiceLine.Type::" ");
                    ServiceLine.Validate("No.", ServiceHeader."Salesperson Code");
                    ServiceLine.Insert(TRUE);

                    if TransferExtendedText.ServCheckifAnyExtText(ServiceLine, False) then
                        TransferExtendedText.InsertServExtText(ServiceLine);
                end;

                if ServiceItemLine."Service Item No." <> '' then begin
                    RepairStatus.Reset();
                    RepairStatus.SetRange(Finished, TRUE);
                    if RepairStatus.FindFirst() then begin
                        ServiceItemLine.Validate("Repair Status Code", RepairStatus.Code);
                        ServiceItemLine.Modify(False);
                    end;
                end;
            end;

            if PreComUpdateQueue.StatusID = '9999' then begin
                ServiceHeader.Get(ServiceHeader."Document Type"::Order, PreComUpdateQueue.ERPReference);
                ServiceHeader.Validate(Status, ServiceHeader.Status::Finished);
                ServiceHeader.MODifY(TRUE);

                SortServiceLines(ServiceHeader);

                if StandardText.Get(ServiceHeader."Salesperson Code") then begin
                    ServiceLine.Reset();
                    ServiceLine.SetRange("Document Type", ServiceHeader."Document Type");
                    ServiceLine.SetRange("Document No.", ServiceHeader."No.");
                    if ServiceLine.FindLast() then
                        LineNo := ServiceLine."Line No." + 10000
                    else
                        LineNo := 10000;
                    ServiceLine.Reset();
                    ServiceLine.Init();
                    ServiceLine.Validate("Document Type", ServiceHeader."Document Type");
                    ServiceLine.Validate("Document No.", ServiceHeader."No.");
                    ServiceLine.Validate("Line No.", LineNo);
                    ServiceLine.Validate("Service Item Line No.", ServiceItemLine."Line No.");
                    ServiceLine.Validate(Type, ServiceLine.Type::" ");
                    ServiceLine.Validate("No.", ServiceHeader."Salesperson Code");
                    ServiceLine.Insert(TRUE);

                    if TransferExtendedText.ServCheckifAnyExtText(ServiceLine, False) then
                        TransferExtendedText.InsertServExtText(ServiceLine);
                end;
            end;
        end;
    end;

    procedure ImportItemInfo(PreComUpdateQueue: Record "PreCom Update Queue")
    var
        ServiceHeader: Record "Service Header";
        ServiceItemLine: Record "Service Item Line";
        ServiceLine: Record "Service Line";
        ServiceLine2: Record "Service Line";
        Item: Record Item;
        LineNo: Integer;
        NoServiceItemLineErr: Label 'There is no Service Item Line on the Service Order!';
        NegativeQtyErr: Label 'No line exist where negative quantity can be deducted!';
        ItemNotExistErr: label 'Item %1 does not exist!', Comment = '%1 is Item No.';
        NegativeQtyHandled: Boolean;
        QuantityRemaining: Decimal;
    begin
        if ServiceHeader.Get(ServiceHeader."Document Type"::Order, PreComUpdateQueue.ERPReference) then
            if Item.Get(PreComUpdateQueue.ArticleNumber) then begin
                NegativeQtyHandled := false;
                ServiceItemLine.Reset();
                ServiceItemLine.SetRange("Document Type", ServiceHeader."Document Type");
                ServiceItemLine.SetRange("Document No.", ServiceHeader."No.");
                if ServiceItemLine.FindFirst() then begin
                    if PreComUpdateQueue."Record ID ERP" = 0 then begin
                        if PreComUpdateQueue.Quantity < 0 then begin
                            if PreComUpdateQueue."PreCom Record ID" <> 0 then begin
                                ServiceLine.Reset();
                                ServiceLine.SetRange("Precom Record ID", PreComUpdateQueue."PreCom Record ID");
                                ServiceLine.SetRange("Document Type", ServiceHeader."Document Type");
                                ServiceLine.SetRange("Document No.", ServiceHeader."No.");
                                if ServiceLine.FindFirst() then begin
                                    if ServiceLine.Quantity + PreComUpdateQueue.Quantity <> 0 then begin
                                        ServiceLine.Validate(Quantity, ServiceLine.Quantity + PreComUpdateQueue.Quantity);
                                        ServiceLine.Modify(true);
                                    end else
                                        ServiceLine.Delete(true);

                                    NegativeQtyHandled := true;
                                END;
                            end;
                            if (PreComUpdateQueue."PreCom Record ID" = 0) or (not NegativeQtyHandled) then begin
                                ServiceLine.Reset();
                                ServiceLine.SetRange("Document Type", ServiceHeader."Document Type");
                                ServiceLine.SetRange("Document No.", ServiceHeader."No.");
                                ServiceLine.SetRange(Type, ServiceLine.Type::Item);
                                ServiceLine.SetRange("No.", PreComUpdateQueue.ArticleNumber);
                                ServiceLine.SetFilter(Quantity, '>=%1', -PreComUpdateQueue.Quantity);
                                if ServiceLine.FindFirst() then begin
                                    ServiceLine.Validate(Quantity, ServiceLine.Quantity + PreComUpdateQueue.Quantity);
                                    ServiceLine.Modify(true);
                                end else begin
                                    QuantityRemaining := -PreComUpdateQueue.Quantity;
                                    ServiceLine.SetRange(Quantity);
                                    ServiceLine.CalcSums(Quantity);
                                    if ServiceLine.Quantity >= QuantityRemaining then begin
                                        IF ServiceLine.FindSet(true, false) then
                                            repeat
                                                if QuantityRemaining >= ServiceLine.Quantity then begin
                                                    QuantityRemaining -= ServiceLine.Quantity;
                                                    ServiceLine2.Get(ServiceLine."Document Type", ServiceLine."Document No.", ServiceLine."Line No.");
                                                    ServiceLine2.Delete(true);
                                                end else begin
                                                    ServiceLine.Validate(Quantity, QuantityRemaining);
                                                    ServiceLine.Modify(true);
                                                    QuantityRemaining := 0;
                                                END;
                                            until ((ServiceLine.Next() <= 0) or (QuantityRemaining <= 0));
                                    end else
                                        ERROR(NegativeQtyErr);

                                end;
                            end;
                        end else
                            if PreComUpdateQueue.Quantity > 0 then begin
                                ServiceLine.Reset();
                                ServiceLine.SetRange("Document Type", ServiceHeader."Document Type");
                                ServiceLine.SetRange("Document No.", ServiceHeader."No.");
                                if ServiceLine.FindLast() then
                                    LineNo := ServiceLine."Line No." + 10000
                                else
                                    LineNo := 10000;

                                ServiceLine.Reset();
                                ServiceLine.SetHideReplacementDialog(TRUE);
                                ServiceLine.Init();
                                ServiceLine.Validate("Document Type", ServiceHeader."Document Type");
                                ServiceLine.Validate("Document No.", ServiceHeader."No.");
                                ServiceLine.Validate("Line No.", LineNo);
                                ServiceLine.Validate("Service Item Line No.", ServiceItemLine."Line No.");
                                ServiceLine.Insert(False);
                                LineNo += 10000;
                                ServiceLine.Validate(Type, ServiceLine.Type::Item);
                                ServiceLine.Validate("No.", PreComUpdateQueue.ArticleNumber);
                                ServiceLine.Validate("Location Code", PreComUpdateQueue.StorePlace);
                                ServiceLine.Validate(Quantity, PreComUpdateQueue.Quantity);
                                ServiceLine.Validate("Precom Line Type", ServiceLine."Precom Line Type"::Material);
                                ServiceLine.Validate("Precom Record ID", PreComUpdateQueue."PreCom Record ID");
                                ServiceLine.Modify(False);
                            end;
                    end else begin
                        ServiceLine.Get(ServiceHeader."Document Type", ServiceHeader."No.", PreComUpdateQueue."Record ID ERP");
                        ServiceLine.SetHideReplacementDialog(TRUE);
                        if ServiceLine."No." <> PreComUpdateQueue.ArticleNumber then
                            ServiceLine.Validate("No.", PreComUpdateQueue.ArticleNumber);
                        if ServiceLine."Location Code" <> PreComUpdateQueue.StorePlace then
                            ServiceLine.Validate("Location Code", PreComUpdateQueue.StorePlace);
                        if ServiceLine.Quantity <> PreComUpdateQueue.Quantity then
                            ServiceLine.Validate(Quantity, PreComUpdateQueue.Quantity);
                        ServiceLine.MODifY(TRUE);
                    end;
                end else
                    ERROR(NoServiceItemLineErr);
            end else
                Error(ItemNotExistErr);
    end;

    procedure ImportTimeInfo(PreComUpdateQueue: Record "PreCom Update Queue")
    var
        ServiceHeader: Record "Service Header";
        ServiceItemLine: Record "Service Item Line";
        ServiceLine: Record "Service Line";
        ServiceLine2: Record "Service Line";
        LineNo: Integer;
        NegativeQtyErr: Label 'No line exist where negative quantity can be deducted!';
        QuantityRemaining: Decimal;
        NegativeQtyHandled: Boolean;
    begin
        if ServiceHeader.Get(ServiceHeader."Document Type"::Order, PreComUpdateQueue.ERPReference) then begin
            NegativeQtyHandled := false;
            ServiceItemLine.Reset();
            ServiceItemLine.SetRange("Document Type", ServiceHeader."Document Type");
            ServiceItemLine.SetRange("Document No.", ServiceHeader."No.");
            if ServiceItemLine.FindFirst() then
                if PreComUpdateQueue.Quantity < 0 then begin
                    if PreComUpdateQueue."PreCom Record ID" <> 0 then begin
                        ServiceLine.Reset();
                        ;
                        ServiceLine.SetRange("Precom Record ID", PreComUpdateQueue."PreCom Record ID");
                        ServiceLine.SetRange("Document Type", ServiceHeader."Document Type");
                        ServiceLine.SetRange("Document No.", ServiceHeader."No.");
                        if ServiceLine.FindFirst() then begin
                            if ServiceLine.Quantity + PreComUpdateQueue.Quantity <> 0 then begin
                                ServiceLine.Validate(Quantity, ServiceLine.Quantity + PreComUpdateQueue.Quantity);
                                ServiceLine.Modify(true);
                            END ELSE
                                ServiceLine.Delete(true);

                            NegativeQtyHandled := true;
                        END;
                    end;
                    if (PreComUpdateQueue."PreCom Record ID" = 0) or (not NegativeQtyHandled) then begin
                        ServiceLine.Reset();
                        ServiceLine.SetRange("Document Type", ServiceHeader."Document Type");
                        ServiceLine.SetRange("Document No.", ServiceHeader."No.");
                        ServiceLine.SetRange(Type, ServiceLine.Type::Resource);
                        ServiceLine.SetRange("No.", PreComUpdateQueue.PrimaryResource);
                        ServiceLine.SETFILTER(Quantity, '>=%1', -PreComUpdateQueue.Quantity);
                        if ServiceLine.FindFirst() then begin
                            ServiceLine.Validate(Quantity, ServiceLine.Quantity + PreComUpdateQueue.Quantity);
                            ServiceLine.MODifY(TRUE);
                        end else begin
                            QuantityRemaining := -PreComUpdateQueue.Quantity;
                            ServiceLine.SetRange(Quantity);
                            ServiceLine.CALCSUMS(Quantity);
                            if ServiceLine.Quantity >= QuantityRemaining then begin
                                if ServiceLine.FindSet() then
                                    Repeat
                                        if QuantityRemaining >= ServiceLine.Quantity then begin
                                            QuantityRemaining -= ServiceLine.Quantity;
                                            ServiceLine2.Get(ServiceLine."Document Type", ServiceLine."Document No.", ServiceLine."Line No.");
                                            ServiceLine2.Delete(true);
                                        end else begin
                                            ServiceLine.Validate(Quantity, QuantityRemaining);
                                            ServiceLine.MODifY(TRUE);
                                            QuantityRemaining := 0;
                                        end;
                                    Until ((ServiceLine.Next() <= 0) OR (QuantityRemaining <= 0));
                            end else
                                ERROR(NegativeQtyErr);
                        end;
                    end;
                end else begin
                    ServiceLine.Reset();
                    ServiceLine.SetRange("Document Type", ServiceHeader."Document Type");
                    ServiceLine.SetRange("Document No.", ServiceHeader."No.");
                    if ServiceLine.FindLast() then
                        LineNo := ServiceLine."Line No." + 10000
                    else
                        LineNo := 10000;

                    ServiceLine.Reset();
                    ServiceLine.SetHideReplacementDialog(TRUE);
                    ServiceLine.Init();
                    ServiceLine.Validate("Document Type", ServiceHeader."Document Type");
                    ServiceLine.Validate("Document No.", ServiceHeader."No.");
                    ServiceLine.Validate("Line No.", LineNo);
                    ServiceLine.Validate("Service Item Line No.", ServiceItemLine."Line No.");
                    ServiceLine.Insert(False);
                    LineNo += 10000;
                    ServiceLine.Validate(Type, ServiceLine.Type::Resource);
                    ServiceLine.Validate("No.", PreComUpdateQueue.PrimaryResource);
                    ServiceLine.Validate(Quantity, PreComUpdateQueue.Quantity);
                    ServiceLine.Validate("Work Type Code", PreComUpdateQueue.Type);
                    ServiceLine.Validate("Precom Line Type", ServiceLine."Precom Line Type"::Hours);
                    ServiceLine.Validate("Precom Record ID", PreComUpdateQueue."PreCom Record ID");
                    ServiceLine.Modify(False);
                end;
        end;
    end;

    procedure ImportCostInfo(PreComUpdateQueue: Record "PreCom Update Queue")
    var
        ServiceHeader: Record "Service Header";
        ServiceItemLine: Record "Service Item Line";
        ServiceLine: Record "Service Line";
        LineNo: Integer;
        NoServiceItemLineErr: Label 'There is no Service Item Line on the Service Order!';
        NegativeQtyErr: Label 'No line exist where negative quantity can be deducted!';
    begin
        if ServiceHeader.Get(ServiceHeader."Document Type"::Order, PreComUpdateQueue.ERPReference) then begin
            ServiceHeader.MODifY(TRUE);

            ServiceItemLine.Reset();
            ServiceItemLine.SetRange("Document Type", ServiceHeader."Document Type");
            ServiceItemLine.SetRange("Document No.", ServiceHeader."No.");
            if ServiceItemLine.FindFirst() then begin
                if PreComUpdateQueue.Quantity < 0 then begin
                    ServiceLine.Reset();
                    ServiceLine.SetRange("Document Type", ServiceHeader."Document Type");
                    ServiceLine.SetRange("Document No.", ServiceHeader."No.");
                    ServiceLine.SetRange(Type, ServiceLine.Type::Cost);
                    ServiceLine.SetRange("No.", PreComUpdateQueue.ArticleNumber);
                    ServiceLine.SETFILTER(Quantity, '>=%1', -PreComUpdateQueue.Quantity);
                    if ServiceLine.FindFirst() then begin
                        ServiceLine.Validate(Quantity, ServiceLine.Quantity + PreComUpdateQueue.Quantity);
                        ServiceLine.Modify(true);
                    end else
                        ERROR(NegativeQtyErr);
                end else begin
                    ServiceLine.Reset();
                    ServiceLine.SetRange("Document Type", ServiceHeader."Document Type");
                    ServiceLine.SetRange("Document No.", ServiceHeader."No.");
                    if ServiceLine.FindLast() then
                        LineNo := ServiceLine."Line No." + 10000
                    else
                        LineNo := 10000;

                    ServiceLine.Reset();
                    ServiceLine.SetHideReplacementDialog(TRUE);
                    ServiceLine.Init();
                    ServiceLine.Validate("Document Type", ServiceHeader."Document Type");
                    ServiceLine.Validate("Document No.", ServiceHeader."No.");
                    ServiceLine.Validate("Line No.", LineNo);
                    ServiceLine.Validate("Service Item Line No.", ServiceItemLine."Line No.");
                    LineNo += 10000;
                    ServiceLine.Validate(Type, ServiceLine.Type::Cost);
                    ServiceLine.Validate("No.", PreComUpdateQueue.ArticleNumber);
                    ServiceLine.Validate(Quantity, PreComUpdateQueue.Quantity);
                    ServiceLine.Validate("Unit Price", PreComUpdateQueue.Price);
                    ServiceLine.Validate("Precom Line Type", ServiceLine."Precom Line Type"::Cost);
                    ServiceLine.Insert(False);
                end;
            end else
                ERROR(NoServiceItemLineErr);
        end;
    end;

    procedure ImportWorkTimeInfo(PreComUpdateQueue: Record "PreCom Update Queue")
    var
        PreComUpdateSetup: Record "PreCom Update Setup";
        TimeRegJournalLine: Record "Time Reg. Journal Line";
        LineNo: Integer;
        StartDateTime: DateTime;
        endDateTime: DateTime;
    begin
        PreComUpdateSetup.Get();
        PreComUpdateSetup.TESTFIELD("Time Reg. Journal");

        TimeRegJournalLine.Reset();
        TimeRegJournalLine.SetRange("Time Reg. Journal Batch", PreComUpdateSetup."Time Reg. Journal");
        if TimeRegJournalLine.FindLast() then
            LineNo := TimeRegJournalLine."Line No." + 10000
        else
            LineNo := 10000;

        Evaluate(StartDateTime, PreComUpdateQueue.PlannedStartDate);
        Evaluate(endDateTime, PreComUpdateQueue.PlannedendDate);

        TimeRegJournalLine.Reset();
        TimeRegJournalLine.Init();
        TimeRegJournalLine.Validate("Time Reg. Journal Batch", PreComUpdateSetup."Time Reg. Journal");
        TimeRegJournalLine.Validate("Line No.", LineNo);
        TimeRegJournalLine.Validate(Date, DT2Date(StartDateTime));
        TimeRegJournalLine.Validate("Start Time", DT2TIME(StartDateTime));
        TimeRegJournalLine.Validate("End Time", DT2TIME(endDateTime));
        TimeRegJournalLine.Validate("Resource Code", PreComUpdateQueue.PrimaryResource);
        TimeRegJournalLine.Validate("Service Order", PreComUpdateQueue.ERPReference);
        TimeRegJournalLine.Validate("Work Type Code", PreComUpdateQueue.Type);
        //TimeRegJournalLine.Validate("Entry Only", TRUE);
        if UpperCase(PreComUpdateQueue.ReferenceType) = 'CI' then
            TimeRegJournalLine.Validate("Time Reg. Type", TimeRegJournalLine."Time Reg. Type"::Arrival);
        if UpperCase(PreComUpdateQueue.ReferenceType) = 'CO' then
            TimeRegJournalLine.Validate("Time Reg. Type", TimeRegJournalLine."Time Reg. Type"::Departure);
        TimeRegJournalLine.Insert(true);
    end;

    procedure TransferItemInventory()
    var
        Item: Record Item;
        Location: Record Location;
        ItemRegister: Record "Item Register";
        TempItem: Record Item temporary;
        ItemLedgEntry: Record "Item Ledger Entry";
        SQLConnection: DotNet NewSqlConnection;
        SQLCommand: DotNet NewSqlCommand;
        SQLParameter: DotNet NewSqlParameter;
        SQLCommandType: DotNet CommandType;
        SQLDBType: DotNet NewSqlDbType;
    begin
        TempItem.DeleteAll();
        ItemRegister.SetRange("Creation Date", TODAY);
        if ItemRegister.FindSet() then
            Repeat
                ItemLedgEntry.SetRange("Entry No.", ItemRegister."From Entry No.", ItemRegister."To Entry No.");
                if ItemLedgEntry.FindSet() then
                    Repeat
                        TempItem.Init();
                        TempItem."No." := ItemLedgEntry."Item No.";
                        if TempItem.Insert(False) then;
                    Until (ItemLedgEntry.Next() <= 0);
            Until (ItemRegister.Next() <= 0);

        TempItem.Reset();
        if TempItem.FindSet() then begin
            if IsNull(SQLConnection) then
                SQLConnection := SQLConnection.SqlConnection();
            SQLConnection.ConnectionString(ReturnConnString());
            SQLConnection.Open();

            Repeat
                Item.Get(TempItem."No.");
                //    if Item."Service Spare Part" then begin
                if Location.FindSet() then
                    Repeat
                        Item.SetRange("Location Filter", Location.Code);
                        Item.CALCFIELDS(Inventory);

                        Clear(SQLCommand);
                        if IsNull(SQLCommand) then
                            SQLCommand := SQLCommand.SqlCommand();
                        SQLCommand.Connection(SQLConnection);
                        SQLCommand.CommandType(SQLCommandType.StoredProcedure);
                        SQLCommand.CommandText('WriteInventory');

                        SQLParameter := SQLParameter.SqlParameter();
                        SQLParameter.ParameterName := '@ArticleNumber';
                        SQLParameter.SqlDbType := SQLDBType.NVarChar;
                        SQLParameter.Size := 255;
                        SQLParameter.Value := Item."No.";
                        SQLCommand.Parameters.Add(SQLParameter);

                        SQLParameter := SQLParameter.SqlParameter();
                        SQLParameter.ParameterName := '@StorePlace';
                        SQLParameter.SqlDbType := SQLDBType.NVarChar;
                        SQLParameter.Size := 255;
                        SQLParameter.Value := Location.Code;
                        SQLCommand.Parameters.Add(SQLParameter);

                        SQLParameter := SQLParameter.SqlParameter();
                        SQLParameter.ParameterName := '@Quantity';
                        SQLParameter.SqlDbType := SQLDBType.Decimal;
                        SQLParameter.Value := Item.Inventory;
                        SQLCommand.Parameters.Add(SQLParameter);

                        SQLParameter := SQLParameter.SqlParameter();
                        SQLParameter.ParameterName := '@OrderPoint';
                        SQLParameter.SqlDbType := SQLDBType.Decimal;
                        SQLParameter.Value := Item."Reorder Point";
                        SQLCommand.Parameters.Add(SQLParameter);

                        SQLCommand.ExecuteNonQuery();
                    Until (Location.Next() <= 0);
            //    end;
            Until (TempItem.Next() <= 0);

            Clear(SQLParameter);
            Clear(SQLCommand);
            SQLConnection.Close();
            Clear(SQLConnection);
        end;
    end;

    procedure TransferItemInventoryTotal()
    var
        Item: Record Item;
        Location: Record Location;
        SQLConnection: DotNet NewSqlConnection;
        SQLCommand: DotNet NewSqlCommand;
        SQLParameter: DotNet NewSqlParameter;
        SQLCommandType: DotNet CommandType;
        SQLDBType: DotNet NewSqlDbType;
    begin
        Item.Reset();
        if Item.FindSet() then begin
            if IsNull(SQLConnection) then
                SQLConnection := SQLConnection.SqlConnection();
            SQLConnection.ConnectionString(ReturnConnString());
            SQLConnection.Open();

            Repeat
                if Location.FindSet() then
                    Repeat
                        Item.SetRange("Location Filter", Location.Code);
                        Item.CALCFIELDS(Inventory);
                        if Item.Inventory > 0 then begin
                            Clear(SQLCommand);
                            if IsNull(SQLCommand) then
                                SQLCommand := SQLCommand.SqlCommand();
                            SQLCommand.Connection(SQLConnection);
                            SQLCommand.CommandType(SQLCommandType.StoredProcedure);
                            SQLCommand.CommandText('WriteInventory');

                            SQLParameter := SQLParameter.SqlParameter();
                            SQLParameter.ParameterName := '@ArticleNumber';
                            SQLParameter.SqlDbType := SQLDBType.NVarChar;
                            SQLParameter.Size := 255;
                            SQLParameter.Value := Item."No.";
                            SQLCommand.Parameters.Add(SQLParameter);

                            SQLParameter := SQLParameter.SqlParameter();
                            SQLParameter.ParameterName := '@StorePlace';
                            SQLParameter.SqlDbType := SQLDBType.NVarChar;
                            SQLParameter.Size := 255;
                            SQLParameter.Value := Location.Code;
                            SQLCommand.Parameters.Add(SQLParameter);

                            SQLParameter := SQLParameter.SqlParameter();
                            SQLParameter.ParameterName := '@Quantity';
                            SQLParameter.SqlDbType := SQLDBType.Decimal;
                            SQLParameter.Value := Item.Inventory;
                            SQLCommand.Parameters.Add(SQLParameter);

                            SQLParameter := SQLParameter.SqlParameter();
                            SQLParameter.ParameterName := '@OrderPoint';
                            SQLParameter.SqlDbType := SQLDBType.Decimal;
                            SQLParameter.Value := Item."Reorder Point";
                            SQLCommand.Parameters.Add(SQLParameter);

                            SQLCommand.ExecuteNonQuery();
                        end;
                    Until (Location.Next() <= 0);
            Until (Item.Next() <= 0);

            Clear(SQLParameter);
            Clear(SQLCommand);
            SQLConnection.Close();
            Clear(SQLConnection);
        end;
    end;

    procedure ImportPreComOrderNos()
    var
        ServiceHeader: Record "Service Header";
        //ServiceItemLine: Record "Service Item Line";
        //ServiceLine: Record "Service Line";
        PrecomOrderNumberLink: Record "Precom Order Number Link";
        SQLConnection: DotNet NewSqlConnection;
        SQLConnection2: DotNet NewSqlConnection;
        SQLCommand: DotNet NewSqlCommand;
        SQLCommand2: DotNet NewSqlCommand;
        SQLDataReader: DotNet NewSqlDataReader;
    //DeleteServiceOrder: Boolean;
    //LineNo: Integer;
    begin
        if IsNull(SQLConnection) then
            SQLConnection := SQLConnection.SqlConnection();
        SQLConnection.ConnectionString(ReturnConnString());
        SQLConnection.Open();

        if IsNull(SQLCommand) then
            SQLCommand := SQLCommand.SqlCommand();
        SQLCommand.Connection(SQLConnection);
        SQLCommand.CommandText('SELECT * FROM INT_Order_Number_OUT WHERE (IntegrationHandleDate is null) Order By IntegrationCreateDate;');
        SQLDataReader := SQLCommand.ExecuteReader();
        WHILE SQLDataReader.Read() DO
            if ServiceHeader.Get(ServiceHeader."Document Type"::Order, Format(SQLDataReader.Item('ExternalId'))) then begin
                if not PrecomOrderNumberLink.Get(ServiceHeader."No.") then begin
                    PrecomOrderNumberLink.Reset();
                    PrecomOrderNumberLink.Init();
                    PrecomOrderNumberLink.Validate("No.", ServiceHeader."No.");
                    PrecomOrderNumberLink.Validate("Precom No.", SQLDataReader.Item('OrderNumber'));
                    PrecomOrderNumberLink.Insert(true);
                end;

                if IsNull(SQLConnection2) then
                    SQLConnection2 := SQLConnection2.SqlConnection();
                SQLConnection2.ConnectionString(ReturnConnString());
                SQLConnection2.Open();
                if IsNull(SQLCommand2) then
                    SQLCommand2 := SQLCommand2.SqlCommand();
                SQLCommand2.Connection(SQLConnection2);
                SQLCommand2.CommandText('UPDATE INT_Order_Number_OUT SET IntegrationHandleDate = ''' + FORMAT(CURRENTDATETIME, 0, '<Year4>-<Month,2>-<Day,2> <Hours24>:<Minutes>') + ''' WHERE Id = ''' + FORMAT(SQLDataReader.Item('Id')) + '''');
                SQLCommand2.ExecuteNonQuery();
                SQLConnection2.Close();
                Clear(SQLConnection2);
            end;

        SQLDataReader.Dispose();
        SQLDataReader.Close();
        Clear(SQLDataReader);
        SQLConnection.Close();
        Clear(SQLConnection);
    end;

    procedure ConvertDate(locDateText: Text[30]): Date
    var
        Year: Integer;
        Month: Integer;
        Day: Integer;
    begin
        if locDateText = '' then
            Exit(0D);
        Evaluate(Day, CopyStr(locDateText, 9, 2));
        Evaluate(Month, CopyStr(locDateText, 6, 2));
        Evaluate(Year, CopyStr(locDateText, 1, 4));
        //Year := Year + 2000;
        //1900-01-01
        // Evaluate(Day, CopyStr(locDateText, 9, 2));
        // Evaluate(Month, CopyStr(locDateText, 6, 2));
        // Evaluate(Year, CopyStr(locDateText, 1, 4));
        Exit(DMY2Date(Day, Month, Year));
    end;

    procedure SplitStringToLines(Streng: Text)
    var
        i: Integer;
        SubStreng: Text[50];
        WordArray: array[150] of Text[50];
        x: Integer;
        Pos: Integer;
        NewStreng: Text;
    begin
        Clear(DescriptionLines);
        Streng := COPYSTR(Streng, 1, 2500);

        NewStreng := '';
        for i := 1 to StrLen(Streng) do
            if (Streng[i] = 10) then // AND (Streng[i][2] = 13)
                NewStreng += ' § '
            else
                NewStreng += FORMAT(Streng[i]);

        Streng := NewStreng;

        i := 1;
        while StrLen(Streng) > 0 do begin
            Pos := STRPOS(Streng, ' ');
            //  if (STRPOS(Streng,'.') < Pos) AND (STRPOS(Streng,'.') > 0) then
            //    Pos := STRPOS(Streng,'.');
            if Pos > 0 then
                SubStreng := COPYSTR(Streng, 1, Pos - 1)
            else
                SubStreng := CopyStr(Streng, 1, 50);
            if CopyStr(SubStreng, 1, 1) = ' ' then
                SubStreng := COPYSTR(SubStreng, 2, 49);
            WordArray[i] := SubStreng;
            if Pos > 0 then
                Streng := COPYSTR(Streng, Pos + 1)
            else
                Streng := '';
            i += 1;
        end;

        COMPRESSARRAY(WordArray);

        x := 1;
        for i := 1 to 150 do
            if WordArray[i] <> '' then
                if WordArray[i] = '§' then
                    x += 1
                else begin
                    if StrLen(DescriptionLines[x]) + StrLen(WordArray[i]) + 1 <= 50 then begin
                        if DescriptionLines[x] <> '' then
                            DescriptionLines[x] += ' ' + WordArray[i]
                        else
                            DescriptionLines[x] += WordArray[i];
                    end else begin
                        x += 1;
                        DescriptionLines[x] := WordArray[i];
                    end;
                    if CopyStr(DescriptionLines[x], StrLen(DescriptionLines[x])) = '.' then
                        x += 1;
                end
            else
                i := 150;

        CompressArray(DescriptionLines);
    end;

    procedure SetGlobalServiceLine(locServiceLine: Record "Service Line")
    begin
        GlobalServiceLine := locServiceLine;
    end;

    local procedure RemoveWorkToDoLines(ServiceHeader: Record "Service Header"): Integer
    var
        ServiceLine: Record "Service Line";
        AllLinesFound: Boolean;
        i: Integer;
        NewWorkToDoLines: Integer;
    begin
        AllLinesFound := False;
        ServiceLine.Reset();
        ServiceLine.SetRange("Document Type", ServiceHeader."Document Type");
        ServiceLine.SetRange("Document No.", ServiceHeader."No.");
        if ServiceLine.FindSet() then
            Repeat
                if ServiceLine.Type = ServiceLine.Type::" " then
                    ServiceLine.Delete(TRUE)
                else
                    AllLinesFound := TRUE;
            Until ((ServiceLine.Next() <= 0) OR (AllLinesFound));

        NewWorkToDoLines := 0;
        if not ServiceLine.FindFirst() then
            Exit(10000)
        else begin
            FOR i := 1 TO 50 DO
                if DescriptionLines[i] <> '' then
                    NewWorkToDoLines += 1;

            //IF NewWorkToDoLines < 2 THEN
            NewWorkToDoLines += 1;

            if ROUND(ServiceLine."Line No." / NewWorkToDoLines, 100, '<') > 0 then
                Exit(ROUND(ServiceLine."Line No." / NewWorkToDoLines, 100, '<'))
            else
                Exit(1);
        end;
    end;

    procedure SetGlobalWorkToDo(locWorkToDo: Text)
    begin
        GlobalWorkToDo := locWorkToDo;
    end;

    local procedure ServItemLineStartingDateReset(ServiceHeader: Record "Service Header"; UseDate: Boolean)
    var
        ServiceItemLine: Record "Service Item Line";
    begin
        ServiceItemLine.Reset();
        ServiceItemLine.SetRange("Document Type", ServiceHeader."Document Type");
        ServiceItemLine.SetRange("Document No.", ServiceHeader."No.");
        if UseDate then
            ServiceItemLine.MODifYALL(ServiceItemLine."Starting Date", ServiceHeader."Starting Date", False)
        else
            ServiceItemLine.MODifYALL(ServiceItemLine."Starting Date", 0D, False);
    end;

    local procedure ReturnLocalDateTime(LocalDate: Date; LocalTime: Time): DateTime
    var
        lLocalTime: Time;
        lDateTimeTxt: Text[30];
        lTimeTxt: Text[30];
        lUTCTime: Time;
        lTimeDiffTxt: Text[30];
        lTimeDiff: Integer;
        lSign: Text[30];
    begin
        Evaluate(lLocalTime, Format(LocalTime));
        lDateTimeTxt := Format(CreateDateTime(LocalDate, lLocalTime), 0, 9);
        lTimeTxt := CopyStr(lDateTimeTxt, StrPos(lDateTimeTxt, 'T') + 1);
        lTimeTxt := CopyStr(lTimeTxt, 1, StrLen(lTimeTxt) - 1);
        Evaluate(lUTCTime, lTimeTxt);
        //lTimeDiffTxt := FORMAT((lLocalTime - lUTCTime) / 3600);
        lTimeDiff := lLocalTime - lUTCTime;
        lSign := '+';
        if lTimeDiffTxt[1] = '-' then begin
            lSign := '-';
            lTimeDiffTxt := DelChr(lTimeDiffTxt, '=', '-');
        end;

        if lSign = '+' then
            LocalTime += lTimeDiff
        else
            LocalTime -= lTimeDiff;

        Exit(CreateDateTime(LocalDate, LocalTime));
    end;

    [TryFunction]
    procedure TestConnection()
    var
        SQLConnection: DotNet NewSqlConnection;
    begin
        if IsNull(SQLConnection) then
            SQLConnection := SQLConnection.SqlConnection();
        SQLConnection.ConnectionString(ReturnConnString());
        SQLConnection.Open();
        SQLConnection.Close();
    end;

    procedure SortServiceLines(pServiceHeader: Record "Service Header")
    var
        ServiceLine: Record "Service Line";
        TempSaveServiceLine: Record "Service Line" temporary;
        ServiceItemLine: Record "Service Item Line";
        LineNo: Integer;
    begin
        ServiceItemLine.Reset();
        ServiceItemLine.SetRange("Document Type", pServiceHeader."Document Type");
        ServiceItemLine.SetRange("Document No.", pServiceHeader."No.");
        if not ServiceItemLine.FindFirst() then
            Clear(ServiceItemLine);

        ServiceLine.Reset();
        ServiceLine.SetRange("Document Type", pServiceHeader."Document Type");
        ServiceLine.SetRange("Document No.", pServiceHeader."No.");
        if ServiceLine.FindSet(False, False) then
            repeat
                TempSaveServiceLine := ServiceLine;
                TempSaveServiceLine.Insert(False);
            until (ServiceLine.Next() = 0);

        ServiceLine.DeleteAll(False);
        ServiceLine.Reset();
        LineNo := 10000;

        TempSaveServiceLine.Reset();
        TempSaveServiceLine.SetRange("Precom Line Type", TempSaveServiceLine."Precom Line Type"::"Work Description");
        if TempSaveServiceLine.FINDSET(False, False) then
            repeat
                ServiceLine := TempSaveServiceLine;
                ServiceLine."Line No." := LineNo;
                LineNo += 10000;
                ServiceLine.Insert(true);
            until (TempSaveServiceLine.Next() = 0);

        ServiceLine.Init();
        ServiceLine.Validate("Document Type", pServiceHeader."Document Type");
        ServiceLine.Validate("Document No.", pServiceHeader."No.");
        ServiceLine.Validate("Line No.", LineNo);
        LineNo += 10000;
        ServiceLine.Type := ServiceLine.Type::" ";
        ServiceLine."Service Item Line No." := ServiceItemLine."Line No.";
        ServiceLine.Description := '..';
        ServiceLine.Insert(true);

        TempSaveServiceLine.Reset();
        TempSaveServiceLine.SetRange("Precom Line Type", TempSaveServiceLine."Precom Line Type"::"Work Done");
        if TempSaveServiceLine.FindSet(False, False) then
            repeat
                ServiceLine := TempSaveServiceLine;
                ServiceLine."Line No." := LineNo;
                LineNo += 10000;
                ServiceLine.Insert(TRUE);
            until (TempSaveServiceLine.Next() = 0);

        ServiceLine.Init();
        ServiceLine.Validate("Document Type", pServiceHeader."Document Type");
        ServiceLine.Validate("Document No.", pServiceHeader."No.");
        ServiceLine.Validate("Line No.", LineNo);
        LineNo += 10000;
        ServiceLine.Type := ServiceLine.Type::" ";
        ServiceLine."Service Item Line No." := ServiceItemLine."Line No.";
        ServiceLine.Description := '..';
        ServiceLine.Insert(true);

        TempSaveServiceLine.Reset();
        TempSaveServiceLine.SetRange("Precom Line Type", TempSaveServiceLine."Precom Line Type"::Hours);
        if TempSaveServiceLine.FindSet(False, False) then
            repeat
                ServiceLine := TempSaveServiceLine;
                ServiceLine."Line No." := LineNo;
                LineNo += 10000;
                ServiceLine.Insert(true);
            until (TempSaveServiceLine.Next() = 0);

        ServiceLine.Init();
        ServiceLine.Validate("Document Type", pServiceHeader."Document Type");
        ServiceLine.Validate("Document No.", pServiceHeader."No.");
        ServiceLine.Validate("Line No.", LineNo);
        LineNo += 10000;
        ServiceLine.Type := ServiceLine.Type::" ";
        ServiceLine."Service Item Line No." := ServiceItemLine."Line No.";
        ServiceLine.Description := '..';
        ServiceLine.Insert(true);

        TempSaveServiceLine.Reset();
        TempSaveServiceLine.SetRange("Precom Line Type", TempSaveServiceLine."Precom Line Type"::Material);
        if TempSaveServiceLine.FindSet(False, False) then
            repeat
                ServiceLine := TempSaveServiceLine;
                ServiceLine."Line No." := LineNo;
                LineNo += 10000;
                ServiceLine.Insert(true);
            until (TempSaveServiceLine.Next() = 0);

        TempSaveServiceLine.Reset();
        TempSaveServiceLine.SetRange("Precom Line Type", TempSaveServiceLine."Precom Line Type"::Cost);
        if TempSaveServiceLine.FindSet(False, False) then
            repeat
                ServiceLine := TempSaveServiceLine;
                ServiceLine."Line No." := LineNo;
                LineNo += 10000;
                ServiceLine.Insert(true);
            until (TempSaveServiceLine.Next() = 0);

        ServiceLine.Init();
        ServiceLine.Validate("Document Type", pServiceHeader."Document Type");
        ServiceLine.Validate("Document No.", pServiceHeader."No.");
        ServiceLine.Validate("Line No.", LineNo);
        LineNo += 10000;
        ServiceLine.Type := ServiceLine.Type::" ";
        ServiceLine."Service Item Line No." := ServiceItemLine."Line No.";
        ServiceLine.Description := '..';
        ServiceLine.Insert(true);

        TempSaveServiceLine.Reset();
        TempSaveServiceLine.SetRange("Precom Line Type", TempSaveServiceLine."Precom Line Type"::" ");
        if TempSaveServiceLine.FindSet(False, False) then
            repeat
                ServiceLine := TempSaveServiceLine;
                ServiceLine."Line No." := LineNo;
                LineNo += 10000;
                ServiceLine.Insert(TRUE);
            until (TempSaveServiceLine.Next() = 0);
    end;
}

