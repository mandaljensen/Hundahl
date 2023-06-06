/// <summary>
/// 14-01-2022 HMJ  Codeunit is used in Job Queue and handles integration.
/// </summary>
codeunit 50502 "PreCom Update Dispatcher"
{
    var
        PrecomUpdateSetup: Record "PreCom Update Setup";

    trigger OnRun()
    begin
        PrecomUpdateSetup.Get();
        if not PrecomUpdateSetup."Use Precom" then
            exit;

        ProcessQueue();
        ImportServiceOrders();
        ImportPreComOrderNos();
        ImportInvoiceInfo();
        ImportItemInfo();
        ImportTimeInfo();
        ImportCostInfo();
        ImportWorkTimeInfo();
        ExportInventory();
    end;

    procedure ProcessQueue()
    var
        PreComUpdateQueue: Record "PreCom Update Queue";
        PreComUpdateQueue2: Record "PreCom Update Queue";
        PreComUpdateManagement: Codeunit "PreCom Update Management";
    begin
        SelectLatestVersion();
        Commit();

        PreComUpdateQueue.Reset();
        PreComUpdateQueue.SetRange("Process Error", false);
        if PreComUpdateQueue.FindSet(false, false) then
            repeat
                if PreComUpdateManagement.Run(PreComUpdateQueue) then begin
                    PreComUpdateQueue2.Get(PreComUpdateQueue."Update Message ID");
                    PreComUpdateQueue2.Delete(true);
                end else begin
                    PreComUpdateQueue2.Get(PreComUpdateQueue."Update Message ID");
                    PreComUpdateQueue2."Process Error" := true;
                    PreComUpdateQueue2."Error Text" := CopyStr(GetLastErrorText, 1, 250);
                    PreComUpdateQueue2.Modify(true);
                end;

                Commit();
            until (PreComUpdateQueue.Next() <= 0);
    end;

    procedure ImportPreComOrderNos()
    var
        PreComUpdateQueue: Record "PreCom Update Queue";
        PreComUpdateManagement: Codeunit "PreCom Update Management";
    begin
        SelectLatestVersion();

        PreComUpdateQueue.Reset();
        PreComUpdateQueue.LockTable(true);
        if PreComUpdateQueue.Find('+') then;

        PreComUpdateQueue.Init();
        PreComUpdateQueue."Update Message ID" := PreComUpdateQueue."Update Message ID" + 1;
        PreComUpdateQueue."Table ID" := -30;
        PreComUpdateManagement.Run(PreComUpdateQueue);
        Commit();
    end;

    procedure ExportInventory()
    var
        PreComUpdateSetup: Record "PreCom Update Setup";
        PreComUpdateManagement: Codeunit "PreCom Update Management";
    begin
        SelectLatestVersion();

        PreComUpdateSetup.Get();
        if CurrentDateTime > PreComUpdateSetup."Next Inventory Update" then begin
            if PreComUpdateSetup."Update Inventory - All Items" then
                PreComUpdateManagement.TransferItemInventoryTotal()
            else
                PreComUpdateManagement.TransferItemInventory();

            PreComUpdateSetup."Next Inventory Update" :=
              ForwardDateTime(PreComUpdateSetup."Next Inventory Update", PreComUpdateSetup."Upd. Inventory Interval (min.)");
            PreComUpdateSetup."Update Inventory - All Items" := false;
            PreComUpdateSetup.Modify(true);
            Commit();
        end;
    end;

    local procedure ForwardDateTime(locDateTime: DateTime; Interval: Integer): DateTime
    var
        NextDate: Date;
        NextTime: Time;
        LastTime: Time;
    begin
        NextDate := DT2Date(locDateTime);
        NextTime := DT2Time(locDateTime);
        LastTime := NextTime;
        NextTime := NextTime + Interval * 60 * 1000;
        if (NextTime < LastTime) or (Interval > 1439) then
            NextDate := NextDate + 1;

        exit(CreateDateTime(NextDate, NextTime));
    end;

    procedure ReturnConnString(): Text[1024]
    var
        PreComUpdateSetup: Record "PreCom Update Setup";
        ConnString: Text[1024];
    begin
        PreComUpdateSetup.Get();
        ConnString :=
            'SERVER=' + PreComUpdateSetup."PreCom SQL Server" + ';'
            + 'DATABASE=' + PreComUpdateSetup."PreCom SQL Database" + ';'
            + 'UID=' + PreComUpdateSetup."PreCom SQL User" + ';'
            + 'PWD=' + PreComUpdateSetup."PreCom SQL Password" + ';'
            + 'Connection Timeout=30';
        exit(ConnString);
    end;

    procedure ImportServiceOrders()
    var
        ServiceHeader: Record "Service Header";
        PreComUpdateQueue: Record "PreCom Update Queue";
        PreComUpdateManagement: Codeunit "PreCom Update Management";
        SQLConnection: DotNet NewSqlConnection;
        SQLConnection2: DotNet NewSqlConnection;
        SQLCommand: DotNet NewSqlCommand;
        SQLCommand2: DotNet NewSqlCommand;
        SQLDataReader: DotNet NewSqlDataReader;
        WorkdoneExternal: Text;
        IdText: Text[40];
        GlobalWorkToDo: Text;
        DT: DateTime;
    begin
        if IsNull(SQLConnection) then
            SQLConnection := SQLConnection.SqlConnection();
        SQLConnection.ConnectionString(ReturnConnString());
        SQLConnection.Open();

        if IsNull(SQLCommand) then
            SQLCommand := SQLCommand.SqlCommand();
        SQLCommand.Connection(SQLConnection);
        SQLCommand.CommandText('SELECT * FROM INT_Order_OUT WHERE (IntegrationHandleDate is null) AND (IntegrationText = '''') Order By IntegrationCreateDate;');
        SQLDataReader := SQLCommand.ExecuteReader();
        while SQLDataReader.Read() do begin
            IdText := Format(SQLDataReader.Item('Id'));
            IdText := DelChr(IdText, '<>', '{}');
            PreComUpdateQueue.Reset();
            PreComUpdateQueue.LockTable(true);
            if PreComUpdateQueue.FindLast() then;

            PreComUpdateQueue.Init();
            PreComUpdateQueue."Update Message ID" := PreComUpdateQueue."Update Message ID" + 1;
            PreComUpdateQueue."Table ID" := -90;
            PreComUpdateQueue.Deleted := SQLDataReader.Item('Deleted');
            PreComUpdateQueue.CustomerNumber := Format(SQLDataReader.Item('CustomerNumber'));
            PreComUpdateQueue.BillingNumber := Format(SQLDataReader.Item('BillingNumber'));
            PreComUpdateQueue.ERPReference := Format(SQLDataReader.Item('ExternalId'));
            DT := SQLDataReader.Item('StartDateTime');
            PreComUpdateQueue.PlannedStartDate := Format(DT2Date(DT), 0, '<Standard Format,9>');
            PreComUpdateQueue.OrderNumber := SQLDataReader.Item('OrderNumber');
            GlobalWorkToDo := SQLDataReader.Item('WorkToDo');
            PreComUpdateQueue.EquipmentNumber := Format(SQLDataReader.Item('EquipmentNumber'));
            WorkdoneExternal := SQLDataReader.Item('WorkDoneExternal');
            if StrLen(WorkdoneExternal) > 250 then begin
                PreComUpdateQueue.WorkDoneExternal1 := CopyStr(WorkdoneExternal, 1, 250);
                WorkdoneExternal := CopyStr(WorkdoneExternal, 251);
            end else begin
                PreComUpdateQueue.WorkDoneExternal1 := CopyStr(WorkdoneExternal, 1, 250);
                WorkdoneExternal := '';
            end;
            if StrLen(WorkdoneExternal) > 250 then begin
                PreComUpdateQueue.WorkDoneExternal2 := CopyStr(WorkdoneExternal, 1, 250);
                WorkdoneExternal := CopyStr(WorkdoneExternal, 251);
            end else begin
                PreComUpdateQueue.WorkDoneExternal2 := CopyStr(WorkdoneExternal, 1, 250);
                WorkdoneExternal := '';
            end;
            if StrLen(WorkdoneExternal) > 250 then begin
                PreComUpdateQueue.WorkDoneExternal3 := CopyStr(WorkdoneExternal, 1, 250);
                WorkdoneExternal := CopyStr(WorkdoneExternal, 251);
            end else begin
                PreComUpdateQueue.WorkDoneExternal3 := CopyStr(WorkdoneExternal, 1, 250);
                WorkdoneExternal := '';
            end;
            PreComUpdateQueue.CauseCode := Format(SQLDataReader.Item('CauseCodeId'));
            PreComUpdateQueue.Type := Format(SQLDataReader.Item('OrderTypeId'));
            PreComUpdateQueue.PrimaryResource := Format(SQLDataReader.Item('DispatcherCode'));
            PreComUpdateQueue.Reference := Format(SQLDataReader.Item('Reference'));
            PreComUpdateQueue.RepairStatus := Format(SQLDataReader.Item('RepairStatus'));
            PreComUpdateQueue.ResponsibilityCenter := Format(SQLDataReader.Item('LocationCode'));
            PreComUpdateQueue.Description := SQLDataReader.Item('K2M_OrderTitle');
            DT := SQLDataReader.Item('AccountingDate');
            PreComUpdateQueue.ActualStartDate := Format(DT2Date(DT), 0, '<Standard Format,9>');
            PreComUpdateQueue.Insert(false);
            Commit();

            //CLEAR(PreComUpdateManagement);
            PreComUpdateManagement.SetGlobalWorkToDo(GlobalWorkToDo);
            if PreComUpdateManagement.Run(PreComUpdateQueue) then begin
                if IsNull(SQLConnection2) then
                    SQLConnection2 := SQLConnection2.SqlConnection();
                SQLConnection2.ConnectionString(ReturnConnString());
                SQLConnection2.Open();
                if IsNull(SQLCommand2) then
                    SQLCommand2 := SQLCommand2.SqlCommand();
                SQLCommand2.Connection(SQLConnection2);
                SQLCommand2.CommandText('UPDATE INT_Order_OUT SET IntegrationHandleDate = ''' + Format(CurrentDateTime, 0, '<Year4>-<Month,2>-<Day,2> <Hours24>:<Minutes>') + ''' WHERE Id = ''' + IdText + '''');
                SQLCommand2.ExecuteNonQuery();
                Clear(SQLCommand2);
                SQLConnection2.Close();
                Clear(SQLConnection2);
            end else begin
                ServiceHeader.SetCurrentKey("Document Type", "No.");
                ServiceHeader.SetRange("Document Type", ServiceHeader."Document Type"::Order);
                if ServiceHeader.FindLast() then
                    if ServiceHeader."Customer No." = '' then
                        ServiceHeader.Delete(true);

                if IsNull(SQLConnection2) then
                    SQLConnection2 := SQLConnection2.SqlConnection();
                SQLConnection2.ConnectionString(ReturnConnString());
                SQLConnection2.Open();
                if IsNull(SQLCommand2) then
                    SQLCommand2 := SQLCommand2.SqlCommand();
                SQLCommand2.Connection(SQLConnection2);
                SQLCommand2.CommandText('UPDATE INT_Order_OUT SET IntegrationText = ''' + DelChr(GetLastErrorText, '=', '''') + ''' WHERE Id = ''' + IdText + '''');
                SQLCommand2.ExecuteNonQuery();
                Clear(SQLCommand2);
                SQLConnection2.Close();
                Clear(SQLConnection2);
            end;
            PreComUpdateQueue.Delete(false);
            Commit();

        end;

        SQLDataReader.Dispose();
        SQLDataReader.Close();
        Clear(SQLDataReader);
        SQLConnection.Close();
        Clear(SQLConnection);
    end;

    procedure ImportInvoiceInfo()
    var
        PreComUpdateQueue: Record "PreCom Update Queue";
        //ServiceHeader: Record "Service Header";
        PreComUpdateManagement: Codeunit "PreCom Update Management";
        SQLConnection: DotNet NewSqlConnection;
        SQLConnection2: DotNet NewSqlConnection;
        SQLCommand: DotNet NewSqlCommand;
        SQLCommand2: DotNet NewSqlCommand;
        SQLDataReader: DotNet NewSqlDataReader;
        //DescriptionText: Text[1024];
        DeleteServiceOrder: Boolean;
        IdText: Text[40];
    begin
        if IsNull(SQLConnection) then
            SQLConnection := SQLConnection.SqlConnection();
        SQLConnection.ConnectionString(ReturnConnString());
        SQLConnection.Open();

        if IsNull(SQLCommand) then
            SQLCommand := SQLCommand.SqlCommand();
        SQLCommand.Connection(SQLConnection);
        SQLCommand.CommandText('SELECT * FROM INT_Invoice_OUT WHERE (IntegrationHandleDate is null) AND (ExternalId <> '''') AND (IntegrationText = '''')');
        SQLDataReader := SQLCommand.ExecuteReader();
        while SQLDataReader.Read() do begin
            IdText := Format(SQLDataReader.Item('Id'));
            IdText := DelChr(IdText, '<>', '{}');
            DeleteServiceOrder := SQLDataReader.Item('Deleted');
            if not DeleteServiceOrder then begin
                PreComUpdateQueue.Reset();
                PreComUpdateQueue.LockTable(true);
                if PreComUpdateQueue.FindLast() then;

                PreComUpdateQueue.Init();
                PreComUpdateQueue."Update Message ID" := PreComUpdateQueue."Update Message ID" + 1;
                PreComUpdateQueue."Table ID" := -100;
                PreComUpdateQueue.ERPReference := Format(SQLDataReader.Item('ExternalId'));
                PreComUpdateQueue.StatusID := Format(SQLDataReader.Item('StatusTypeId'));
                PreComUpdateQueue.EquipmentReading := Format(SQLDataReader.Item('EndReading'));
                PreComUpdateQueue.EquipmentUsageReading := Format(SQLDataReader.Item('EndUsageReading'));
                PreComUpdateQueue.EquipmentMileageReading := Format(SQLDataReader.Item('EndMilageReading'));
                PreComUpdateQueue.ReadingDate := Format(SQLDataReader.Item('EndReadingDate'));
                PreComUpdateQueue.Insert(false);
                Commit();

                if PreComUpdateManagement.Run(PreComUpdateQueue) then begin
                    if IsNull(SQLConnection2) then
                        SQLConnection2 := SQLConnection2.SqlConnection();
                    SQLConnection2.ConnectionString(ReturnConnString());
                    SQLConnection2.Open();
                    if IsNull(SQLCommand2) then
                        SQLCommand2 := SQLCommand2.SqlCommand();
                    SQLCommand2.Connection(SQLConnection2);
                    SQLCommand2.CommandText('UPDATE INT_Invoice_OUT SET IntegrationHandleDate = ''' + Format(CurrentDateTime, 0, '<Year4>-<Month,2>-<Day,2> <Hours24>:<Minutes>') + ''' WHERE Id = ''' + IdText + '''');
                    SQLCommand2.ExecuteNonQuery();
                    Clear(SQLCommand2);
                    SQLConnection2.Close();
                    Clear(SQLConnection2);
                end else begin
                    if IsNull(SQLConnection2) then
                        SQLConnection2 := SQLConnection2.SqlConnection();
                    SQLConnection2.ConnectionString(ReturnConnString());
                    SQLConnection2.Open();
                    if IsNull(SQLCommand2) then
                        SQLCommand2 := SQLCommand2.SqlCommand();
                    SQLCommand2.Connection(SQLConnection2);
                    SQLCommand2.CommandText('UPDATE INT_Invoice_OUT SET IntegrationText = ''' + DelChr(GetLastErrorText, '=', '''') + ''' WHERE Id = ''' + IdText + '''');
                    SQLCommand2.ExecuteNonQuery();
                    Clear(SQLCommand2);
                    SQLConnection2.Close();
                    Clear(SQLConnection2);
                end;

                PreComUpdateQueue.Delete(false);
                Commit();
            end;
        end;

        SQLDataReader.Dispose();
        SQLDataReader.Close();
        Clear(SQLDataReader);
        SQLConnection.Close();
        Clear(SQLConnection);
    end;

    procedure ImportItemInfo()
    var
        PreComUpdateQueue: Record "PreCom Update Queue";
        //ServiceHeader: Record "Service Header";
        PreComUpdateManagement: Codeunit "PreCom Update Management";
        SQLConnection: DotNet NewSqlConnection;
        SQLConnection2: DotNet NewSqlConnection;
        SQLCommand: DotNet NewSqlCommand;
        SQLCommand2: DotNet NewSqlCommand;
        SQLDataReader: DotNet NewSqlDataReader;
        IdText: Text[40];
    begin
        if IsNull(SQLConnection) then
            SQLConnection := SQLConnection.SqlConnection();
        SQLConnection.ConnectionString(ReturnConnString());
        SQLConnection.Open();

        if IsNull(SQLCommand) then
            SQLCommand := SQLCommand.SqlCommand();
        SQLCommand.Connection(SQLConnection);
        SQLCommand.CommandText('SELECT * FROM INT_Order_Invoice_Material_OUT WHERE (IntegrationHandleDate is null) AND (ExternalId <> '''') AND (IntegrationText = '''') Order By IntegrationCreateDate;');
        SQLDataReader := SQLCommand.ExecuteReader();
        while SQLDataReader.Read() do begin
            IdText := Format(SQLDataReader.Item('Id'));
            IdText := DelChr(IdText, '<>', '{}');
            PreComUpdateQueue.Reset();
            PreComUpdateQueue.LockTable(true);
            if PreComUpdateQueue.FindLast() then;

            PreComUpdateQueue.Init();
            PreComUpdateQueue."Update Message ID" := PreComUpdateQueue."Update Message ID" + 1;
            PreComUpdateQueue."Table ID" := -110;
            PreComUpdateQueue.ERPReference := Format(SQLDataReader.Item('ExternalId'));
            PreComUpdateQueue.ArticleNumber := Format(SQLDataReader.Item('ArticleNumber'));
            PreComUpdateQueue.Quantity := ConvertToDecimal(Format(SQLDataReader.Item('Quantity')));
            PreComUpdateQueue.StorePlace := SQLDataReader.Item('ResourceExternalId');
            PreComUpdateQueue."Record ID ERP" := ConvertToInteger(Format(SQLDataReader.Item('RecordIdErp')));
            PreComUpdateQueue."PreCom Record ID" := ConvertToInteger(Format(SQLDataReader.Item('RecordId')));
            PreComUpdateQueue.Deleted := SQLDataReader.Item('Deleted');
            PreComUpdateQueue.Insert(false);
            Commit();

            if PreComUpdateManagement.Run(PreComUpdateQueue) then begin
                if IsNull(SQLConnection2) then
                    SQLConnection2 := SQLConnection2.SqlConnection();
                SQLConnection2.ConnectionString(ReturnConnString());
                SQLConnection2.Open();
                if IsNull(SQLCommand2) then
                    SQLCommand2 := SQLCommand2.SqlCommand();
                SQLCommand2.Connection(SQLConnection2);
                SQLCommand2.CommandText('UPDATE INT_Order_Invoice_Material_OUT SET IntegrationHandleDate = ''' + Format(CurrentDateTime, 0, '<Year4>-<Month,2>-<Day,2> <Hours24>:<Minutes>') + ''' WHERE Id = ''' + IdText + '''');
                SQLCommand2.ExecuteNonQuery();
                Clear(SQLCommand2);
                SQLConnection2.Close();
                Clear(SQLConnection2);
            end else begin
                if IsNull(SQLConnection2) then
                    SQLConnection2 := SQLConnection2.SqlConnection();
                SQLConnection2.ConnectionString(ReturnConnString());
                SQLConnection2.Open();
                if IsNull(SQLCommand2) then
                    SQLCommand2 := SQLCommand2.SqlCommand();
                SQLCommand2.Connection(SQLConnection2);
                SQLCommand2.CommandText('UPDATE INT_Order_Invoice_Material_OUT SET IntegrationText = ''' + DelChr(GetLastErrorText, '=', '''') + ''' WHERE Id = ''' + IdText + '''');
                SQLCommand2.ExecuteNonQuery();
                Clear(SQLCommand2);
                SQLConnection2.Close();
                Clear(SQLConnection2);
            end;

            PreComUpdateQueue.Delete(false);
            Commit();
        end;

        SQLDataReader.Dispose();
        SQLDataReader.Close();
        Clear(SQLDataReader);
        SQLConnection.Close();
        Clear(SQLConnection);
    end;

    procedure ImportTimeInfo()
    var
        PreComUpdateQueue: Record "PreCom Update Queue";
        PreComUpdateManagement: Codeunit "PreCom Update Management";
        SQLConnection: DotNet NewSqlConnection;
        SQLConnection2: DotNet NewSqlConnection;
        SQLCommand: DotNet NewSqlCommand;
        SQLCommand2: DotNet NewSqlCommand;
        SQLDataReader: DotNet NewSqlDataReader;
        IdText: Text[40];
    begin
        if IsNull(SQLConnection) then
            SQLConnection := SQLConnection.SqlConnection();
        SQLConnection.ConnectionString(ReturnConnString());
        SQLConnection.Open();

        if IsNull(SQLCommand) then
            SQLCommand := SQLCommand.SqlCommand();
        SQLCommand.Connection(SQLConnection);
        SQLCommand.CommandText('SELECT * FROM INT_Order_Invoice_Time_OUT WHERE (IntegrationHandleDate is null) AND (ExternalId <> '''') AND (IntegrationText = '''') Order By IntegrationCreateDate;');
        SQLDataReader := SQLCommand.ExecuteReader();
        while SQLDataReader.Read() do begin
            IdText := Format(SQLDataReader.Item('Id'));
            IdText := DelChr(IdText, '<>', '{}');
            PreComUpdateQueue.Reset();
            PreComUpdateQueue.LockTable(true);
            if PreComUpdateQueue.FindLast() then;

            PreComUpdateQueue.Init();
            PreComUpdateQueue."Update Message ID" := PreComUpdateQueue."Update Message ID" + 1;
            PreComUpdateQueue."Table ID" := -120;
            PreComUpdateQueue.ERPReference := Format(SQLDataReader.Item('ExternalId'));
            PreComUpdateQueue.Quantity := ConvertToDecimal(Format(SQLDataReader.Item('Quantity')));
            PreComUpdateQueue.Type := Format(SQLDataReader.Item('TimeType'));
            PreComUpdateQueue.Price := ConvertToDecimal(Format(SQLDataReader.Item('Price')));
            PreComUpdateQueue.PrimaryResource := SQLDataReader.Item('ResourceExternalID');
            PreComUpdateQueue."PreCom Record ID" := ConvertToInteger(SQLDataReader.Item('RecordId'));
            PreComUpdateQueue.Insert(false);
            Commit();

            if PreComUpdateManagement.Run(PreComUpdateQueue) then begin
                if IsNull(SQLConnection2) then
                    SQLConnection2 := SQLConnection2.SqlConnection();
                SQLConnection2.ConnectionString(ReturnConnString());
                SQLConnection2.Open();
                if IsNull(SQLCommand2) then
                    SQLCommand2 := SQLCommand2.SqlCommand();
                SQLCommand2.Connection(SQLConnection2);
                SQLCommand2.CommandText('UPDATE INT_Order_Invoice_Time_OUT SET IntegrationHandleDate = ''' + Format(CurrentDateTime, 0, '<Year4>-<Month,2>-<Day,2> <Hours24>:<Minutes>') + ''' WHERE Id = ''' + IdText + '''');
                SQLCommand2.ExecuteNonQuery();
                Clear(SQLCommand2);
                SQLConnection2.Close();
                Clear(SQLConnection2);
            end else begin
                if IsNull(SQLConnection2) then
                    SQLConnection2 := SQLConnection2.SqlConnection();
                SQLConnection2.ConnectionString(ReturnConnString());
                SQLConnection2.Open();
                if IsNull(SQLCommand2) then
                    SQLCommand2 := SQLCommand2.SqlCommand();
                SQLCommand2.Connection(SQLConnection2);
                SQLCommand2.CommandText('UPDATE INT_Order_Invoice_Time_OUT SET IntegrationText = ''' + DelChr(GetLastErrorText, '=', '''') + ''' WHERE Id = ''' + IdText + '''');
                SQLCommand2.ExecuteNonQuery();
                Clear(SQLCommand2);
                SQLConnection2.Close();
                Clear(SQLConnection2);
            end;

            PreComUpdateQueue.Delete(false);
            Commit();
        end;

        SQLDataReader.Dispose();
        SQLDataReader.Close();
        Clear(SQLDataReader);
        SQLConnection.Close();
        Clear(SQLConnection);
    end;

    procedure ImportCostInfo()
    var
        PreComUpdateQueue: Record "PreCom Update Queue";
        //ServiceHeader: Record "Service Header";
        PreComUpdateManagement: Codeunit "PreCom Update Management";
        SQLConnection: DotNet NewSqlConnection;
        SQLConnection2: DotNet NewSqlConnection;
        SQLCommand: DotNet NewSqlCommand;
        SQLCommand2: DotNet NewSqlCommand;
        SQLDataReader: DotNet NewSqlDataReader;
        DeleteServiceOrder: Boolean;
        IdText: Text[40];
    begin
        if IsNull(SQLConnection) then
            SQLConnection := SQLConnection.SqlConnection();
        SQLConnection.ConnectionString(ReturnConnString());
        SQLConnection.Open();

        if IsNull(SQLCommand) then
            SQLCommand := SQLCommand.SqlCommand();
        SQLCommand.Connection(SQLConnection);
        SQLCommand.CommandText('SELECT * FROM INT_Order_Invoice_Cost_OUT WHERE (IntegrationHandleDate is null) AND (ExternalId <> '''') AND (IntegrationText = '''') Order By IntegrationCreateDate;');
        SQLDataReader := SQLCommand.ExecuteReader();
        while SQLDataReader.Read() do begin
            IdText := Format(SQLDataReader.Item('Id'));
            IdText := DelChr(IdText, '<>', '{}');
            DeleteServiceOrder := SQLDataReader.Item('Deleted');
            if not DeleteServiceOrder then begin
                PreComUpdateQueue.Reset();
                PreComUpdateQueue.LockTable(true);
                if PreComUpdateQueue.FindLast() then;

                PreComUpdateQueue.Init();
                PreComUpdateQueue."Update Message ID" := PreComUpdateQueue."Update Message ID" + 1;
                PreComUpdateQueue."Table ID" := -130;
                PreComUpdateQueue.ERPReference := Format(SQLDataReader.Item('ExternalId'));
                PreComUpdateQueue.ArticleNumber := Format(SQLDataReader.Item('CostItem'));
                PreComUpdateQueue.Description := Format(SQLDataReader.Item('CostDescription'));
                PreComUpdateQueue.Quantity := ConvertToDecimal(Format(SQLDataReader.Item('Quantity')));
                PreComUpdateQueue.Price := ConvertToDecimal(Format(SQLDataReader.Item('Price')));
                PreComUpdateQueue.Insert(false);
                Commit();

                if PreComUpdateManagement.Run(PreComUpdateQueue) then begin
                    if IsNull(SQLConnection2) then
                        SQLConnection2 := SQLConnection2.SqlConnection();
                    SQLConnection2.ConnectionString(ReturnConnString());
                    SQLConnection2.Open();
                    if IsNull(SQLCommand2) then
                        SQLCommand2 := SQLCommand2.SqlCommand();
                    SQLCommand2.Connection(SQLConnection2);
                    SQLCommand2.CommandText('UPDATE INT_Order_Invoice_Cost_OUT SET IntegrationHandleDate = ''' + Format(CurrentDateTime, 0, '<Year4>-<Month,2>-<Day,2> <Hours24>:<Minutes>') + ''' WHERE Id = ''' + IdText + '''');
                    SQLCommand2.ExecuteNonQuery();
                    Clear(SQLCommand2);
                    SQLConnection2.Close();
                    Clear(SQLConnection2);
                end else begin
                    if IsNull(SQLConnection2) then
                        SQLConnection2 := SQLConnection2.SqlConnection();
                    SQLConnection2.ConnectionString(ReturnConnString());
                    SQLConnection2.Open();
                    if IsNull(SQLCommand2) then
                        SQLCommand2 := SQLCommand2.SqlCommand();
                    SQLCommand2.Connection(SQLConnection2);
                    SQLCommand2.CommandText('UPDATE INT_Order_Invoice_Cost_OUT SET IntegrationText = ''' + DelChr(GetLastErrorText, '=', '''') + ''' WHERE Id = ''' + IdText + '''');
                    SQLCommand2.ExecuteNonQuery();
                    Clear(SQLCommand2);
                    SQLConnection2.Close();
                    Clear(SQLConnection2);
                end;

                PreComUpdateQueue.Delete(false);
                Commit();
            end;
        end;

        SQLDataReader.Dispose();
        SQLDataReader.Close();
        Clear(SQLDataReader);
        SQLConnection.Close();
        Clear(SQLConnection);
    end;

    procedure ImportWorkTimeInfo()
    var
        PreComUpdateQueue: Record "PreCom Update Queue";
        PreComUpdateManagement: Codeunit "PreCom Update Management";
        SQLConnection: DotNet NewSqlConnection;
        SQLConnection2: DotNet NewSqlConnection;
        SQLCommand: DotNet NewSqlCommand;
        SQLCommand2: DotNet NewSqlCommand;
        SQLDataReader: DotNet NewSqlDataReader;
        IdText: Text[40];
        DT: DateTime;
    begin
        if IsNull(SQLConnection) then
            SQLConnection := SQLConnection.SqlConnection();
        SQLConnection.ConnectionString(ReturnConnString());
        SQLConnection.Open();

        if IsNull(SQLCommand) then
            SQLCommand := SQLCommand.SqlCommand();
        SQLCommand.Connection(SQLConnection);
        SQLCommand.CommandText('SELECT * FROM INT_Work_Time_OUT WHERE (IntegrationHandleDate is null) AND (IntegrationText = '''') AND (RegistrationType = ''-1'') Order By IntegrationCreateDate;');
        SQLDataReader := SQLCommand.ExecuteReader();
        while SQLDataReader.Read() do begin
            IdText := Format(SQLDataReader.Item('Id'));
            IdText := DelChr(IdText, '<>', '{}');
            PreComUpdateQueue.Reset();
            PreComUpdateQueue.LockTable(true);
            if PreComUpdateQueue.FindLast() then;

            PreComUpdateQueue.Init();
            PreComUpdateQueue."Update Message ID" := PreComUpdateQueue."Update Message ID" + 1;
            PreComUpdateQueue."Table ID" := -140;
            PreComUpdateQueue.ERPReference := SQLDataReader.Item('ExternalId');
            DT := SQLDataReader.Item('StartDate');
            PreComUpdateQueue.PlannedStartDate := Format(DT2Date(DT), 0, '<Standard Format,9>');
            DT := SQLDataReader.Item('EndDate');
            PreComUpdateQueue.PlannedEndDate := Format(DT2Date(DT), 0, '<Standard Format,9>');
            PreComUpdateQueue.PrimaryResource := SQLDataReader.Item('ResourceExternalID');
            PreComUpdateQueue.Type := Format(SQLDataReader.Item('TimeTypeConverted'));
            PreComUpdateQueue.Insert(false);
            Commit();

            if PreComUpdateManagement.Run(PreComUpdateQueue) then begin
                if IsNull(SQLConnection2) then
                    SQLConnection2 := SQLConnection2.SqlConnection();
                SQLConnection2.ConnectionString(ReturnConnString());
                SQLConnection2.Open();
                if IsNull(SQLCommand2) then
                    SQLCommand2 := SQLCommand2.SqlCommand();
                SQLCommand2.Connection(SQLConnection2);
                SQLCommand2.CommandText('UPDATE INT_Work_Time_OUT SET IntegrationHandleDate = ''' + Format(CurrentDateTime, 0, '<Year4>-<Month,2>-<Day,2> <Hours24>:<Minutes>') + ''' WHERE Id = ''' + IdText + '''');
                SQLCommand2.ExecuteNonQuery();
                Clear(SQLCommand2);
                SQLConnection2.Close();
                Clear(SQLConnection2);
            end else begin
                if IsNull(SQLConnection2) then
                    SQLConnection2 := SQLConnection2.SqlConnection();
                SQLConnection2.ConnectionString(ReturnConnString());
                SQLConnection2.Open();
                if IsNull(SQLCommand2) then
                    SQLCommand2 := SQLCommand2.SqlCommand();
                SQLCommand2.Connection(SQLConnection2);
                SQLCommand2.CommandText('UPDATE INT_Work_Time_OUT SET IntegrationText = ''' + DelChr(GetLastErrorText, '=', '''') + ''' WHERE Id = ''' + IdText + '''');
                SQLCommand2.ExecuteNonQuery();
                Clear(SQLCommand2);
                SQLConnection2.Close();
                Clear(SQLConnection2);
            end;

            PreComUpdateQueue.Delete(false);
            Commit();
        end;

        SQLDataReader.Dispose();
        SQLDataReader.Close();
        Clear(SQLDataReader);
        SQLConnection.Close();
        Clear(SQLConnection);
    end;

    local procedure ConvertToDecimal(String: Text): Decimal
    var
        DecimalValue: Decimal;
    begin
        if Evaluate(DecimalValue, String) then
            exit(DecimalValue)
        else
            exit(0);
    end;

    local procedure ConvertToInteger(String: Text): Integer
    var
        IntegerValue: Integer;
    begin
        if Evaluate(IntegerValue, String) then
            exit(IntegerValue)
        else
            exit(0);
    end;
}

