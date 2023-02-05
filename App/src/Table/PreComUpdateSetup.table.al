/// <summary>
/// 14-01-2022 HMJ  Table contains setup values for Precom Integration.
/// </summary>
table 50501 "PreCom Update Setup"
{
    fields
    {
        field(1; "Primary Key"; Integer)
        {
            Caption = 'Primary Key';
        }
        field(2; "PreCom SQL Server"; Text[200])
        {
            Caption = 'PreCom SQL Server';
        }
        field(3; "PreCom SQL User"; Text[50])
        {
            Caption = 'PreCom SQL User';
        }
        field(4; "PreCom SQL Password"; Text[30])
        {
            Caption = 'PreCom SQL Password';
        }
        field(5; "PreCom SQL Database"; Text[200])
        {
            Caption = 'PreCom SQL Database';
        }
        field(6; "Upd. Inventory Interval (min.)"; Integer)
        {
            Caption = 'Upd. Inventory Interval (min.)';
        }
        field(7; "Next Inventory Update"; DateTime)
        {
            Caption = 'Next Inventory Update';
        }
        field(8; "Default Service Order Type"; Code[10])
        {
            Caption = 'Default Service Order Type';
            TableRelation = "Service Order Type";
        }
        field(9; "Default Dimension 1 Code"; Code[20])
        {
            CaptionClass = '1,2,1';
            Caption = 'Default Dimension 1 Code';
            TableRelation = "Dimension Value".Code WHERE("Global Dimension No." = CONST(1));
        }
        field(10; "Update Inventory - All Items"; Boolean)
        {
            Caption = 'Update Inventory - All Items';
        }
        field(11; "Help Materials Item No."; Code[20])
        {
            Caption = 'Help Materials Item No.';
            TableRelation = Item;
        }
        field(12; "Help Material Percent"; Decimal)
        {
            Caption = 'Help Material Percent';
            MinValue = 0;
        }
        field(13; "Document Path"; Text[250])
        {
            Caption = 'Document Path';

            trigger OnLookup()
            begin
                FileManagement.SelectFolderDialog(DocumentPathFolderSelectionCaptionLbl, "Document Path");
            end;
        }
        field(14; "Time Reg. Journal"; Code[20])
        {
            Caption = 'Time Reg. Journal';
            TableRelation = "Time Reg. Journal Batch";
        }
        field(15; "Use Precom"; Boolean)
        {
            Caption = 'Use Precom';
            DataClassification = CustomerContent;
        }
    }

    keys
    {
        key(Key1; "Primary Key")
        {
            Clustered = true;
        }
    }

    fieldgroups
    {
    }

    var
        FileManagement: Codeunit "File Management";
        DocumentPathFolderSelectionCaptionLbl: Label 'Choose folder';
}

