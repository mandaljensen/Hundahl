tableextension 50500 "Service Line Ext." extends "Service Line"
{
    fields
    {
        field(50005; "Precom Line Type"; Option)
        {
            Caption = 'Precom Line Type';
            DataClassification = ToBeClassified;
            OptionMembers = " ","Work Description","Work Done","Hours","Material","Cost";
        }
        field(50006; "Precom Record ID"; Integer)
        {
            Caption = 'Precom Record ID';
            DataClassification = ToBeClassified;
        }
    }

    trigger OnAfterInsert()
    begin
        if "Precom Line Type" = "Precom Line Type"::" " then
            case Type of
                Type::Cost:
                    "Precom Line Type" := "Precom Line Type"::Cost;
                Type::Item:
                    "Precom Line Type" := "Precom Line Type"::Material;
                Type::Resource:
                    "Precom Line Type" := "Precom Line Type"::Hours;
            end;
    end;
}
