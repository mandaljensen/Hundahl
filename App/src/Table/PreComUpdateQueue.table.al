table 50500 "PreCom Update Queue"
{
    fields
    {
        field(1; "Update Message ID"; Integer)
        {
        }
        field(2; RecordID; RecordID)
        {
        }
        field(3; "Command Type"; Option)
        {
            OptionMembers = Insert,Update,Delete,Rename;
        }
        field(4; "Table ID"; Integer)
        {
        }
        field(5; "Process Error"; Boolean)
        {
        }
        field(6; "Error Text"; Text[250])
        {
        }
        field(100; ERPReference; Text[50])
        {
        }
        field(101; ActualEndDate; Text[25])
        {
        }
        field(102; StatusID; Text[250])
        {
        }
        field(103; EquipmentNumber; Text[50])
        {
        }
        field(104; WorkDoneExternal1; Text[250])
        {
        }
        field(105; EquipmentReading; Text[250])
        {
        }
        field(106; ReadingDate; Text[25])
        {
        }
        field(107; ArticleNumber; Text[50])
        {
        }
        field(108; StorePlace; Text[120])
        {
        }
        field(109; Quantity; Decimal)
        {
        }
        field(110; UserID; Text[50])
        {
        }
        field(111; CauseCode; Text[50])
        {
        }
        field(112; Value; Decimal)
        {
        }
        field(113; ReferenceType; Text[250])
        {
        }
        field(114; Description; Text[250])
        {
        }
        field(115; Price; Decimal)
        {
        }
        field(116; CustomerNumber; Text[50])
        {
        }
        field(117; PrimaryResource; Text[50])
        {
        }
        field(118; PlannedStartDate; Text[25])
        {
        }
        field(119; PlannedEndDate; Text[25])
        {
        }
        field(120; OrderNumber; Integer)
        {
        }
        field(121; WorkToDo; Text[250])
        {
        }
        field(122; ContractNumber; Text[50])
        {
        }
        field(123; Type; Text[50])
        {
        }
        field(124; Deleted; Boolean)
        {
        }
        field(125; WorkDoneExternal2; Text[250])
        {
        }
        field(126; WorkDoneExternal3; Text[250])
        {
        }
        field(127; ActualStartDate; Text[25])
        {
        }
        field(128; BillingNumber; Text[50])
        {
        }
        field(129; Reference; Text[250])
        {
        }
        field(130; "PreCom Record ID"; Integer)
        {
        }
        field(131; ContractPeriodStart; Text[25])
        {
        }
        field(132; ContractMainPeriodStart; Text[25])
        {
        }
        field(133; ContractLastPeriod; Text[25])
        {
        }
        field(134; ContractLastMainPeriod; Text[25])
        {
        }
        field(135; EquipmentUsageReading; Text[30])
        {
        }
        field(136; EquipmentMileageReading; Text[30])
        {
        }
        field(137; RepairStatus; Text[10])
        {
        }
        field(138; ResponsibilityCenter; Text[30])
        {
            DataClassification = ToBeClassified;
        }
        field(139; "Record ID ERP"; Integer)
        {

        }
    }

    keys
    {
        key(Key1; "Update Message ID")
        {
            Clustered = true;
        }
    }

    fieldgroups
    {
    }
}

