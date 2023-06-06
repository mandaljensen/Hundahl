/// <summary>
/// 14-01-2022 HMJ  Page created.
/// </summary>
page 50500 "PreCom Update Setup"
{
    Caption = 'PreCom Update Setup';
    PageType = Card;
    SourceTable = "PreCom Update Setup";
    ApplicationArea = All;
    UsageCategory = Administration;

    layout
    {
        area(content)
        {
            group(Generelt)
            {
                field("Use Precom"; "Use Precom")
                {
                    ApplicationArea = All;
                    ToolTip = 'Specifies the value of the Use Precom field.';
                }
                field("PreCom SQL Server"; "PreCom SQL Server")
                {
                    ToolTip = 'Specifies the value of the PreCom SQL Server field.';
                    ApplicationArea = All;
                }
                field("PreCom SQL Database"; "PreCom SQL Database")
                {
                    ToolTip = 'Specifies the value of the PreCom SQL Database field.';
                    ApplicationArea = All;
                }
                field("PreCom SQL User"; "PreCom SQL User")
                {
                    ToolTip = 'Specifies the value of the PreCom SQL User field.';
                    ApplicationArea = All;
                }
                field("PreCom SQL Password"; "PreCom SQL Password")
                {
                    ToolTip = 'Specifies the value of the PreCom SQL Password field.';
                    ApplicationArea = All;
                }
                field("Upd. Inventory Interval (min.)"; "Upd. Inventory Interval (min.)")
                {
                    ToolTip = 'Specifies the value of the Upd. Inventory Interval (min.) field.';
                    ApplicationArea = All;
                }
                field("Next Inventory Update"; "Next Inventory Update")
                {
                    ToolTip = 'Specifies the value of the Next Inventory Update field.';
                    ApplicationArea = All;
                }
                field("Default Service Order Type"; "Default Service Order Type")
                {
                    ToolTip = 'Specifies the value of the Default Service Order Type field.';
                    ApplicationArea = All;
                }
                field("Default Dimension 1 Code"; "Default Dimension 1 Code")
                {
                    ToolTip = 'Specifies the value of the Default Dimension 1 Code field.';
                    ApplicationArea = All;
                }
                field("Document Path"; "Document Path")
                {
                    ToolTip = 'Specifies the value of the Document Path field.';
                    ApplicationArea = All;
                }
                field("Time Reg. Journal"; "Time Reg. Journal")
                {
                    ToolTip = 'Specifies the value of the Time Reg. Journal field.';
                    ApplicationArea = All;
                }
            }
        }
    }

    actions
    {
        area(Processing)
        {
            action(Process)
            {
                ApplicationArea = All;
                Caption = 'Process';
                Image = Process;
                RunObject = codeunit "PreCom Update Dispatcher";
                ToolTip = 'Executes the Process action.';
                Visible = true;
            }
            action(TestDatabaseConnection)
            {
                Caption = 'Test Database Connection';
                ApplicationArea = All;
                ToolTip = 'Executes the Test Database Connection action.';
                Image = TestDatabase;
                Promoted = true;
                PromotedCategory = Process;
                PromotedIsBig = true;

                trigger OnAction()
                var
                    PrecomUpdateManagement: Codeunit "PreCom Update Management";
                    ConnectionSuccesLbl: label 'Connection OK';
                    ConnectionFailLbl: Label 'Connection failed with the following error:\%1', Comment = '%1 = Latest error text';
                begin
                    if PrecomUpdateManagement.TestConnection() then
                        Message(ConnectionSuccesLbl)
                    else
                        Message(StrSubstNo(ConnectionFailLbl, GetLastErrorText()));
                end;
            }
            action(TestConvertDate)
            {
                Caption = 'Test Convert Date';
                ApplicationArea = All;
                Image = TestReport;
                Visible = false;

                trigger OnAction()
                var
                    PrecomUpdateManagement: Codeunit "PreCom Update Management";
                    DT: DateTime;
                    ImportDate: Text[25];
                begin
                    DT := CreateDateTime(19000101D, 000000T);
                    ImportDate := Format(DT2Date(DT), 0, '<Standard Format,9>');
                    PrecomUpdateManagement.ConvertDate(ImportDate)
                end;
            }
        }
    }
}

