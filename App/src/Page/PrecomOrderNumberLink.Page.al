page 50502 "Precom Order Number Link"
{
    ApplicationArea = All;
    Caption = 'Precom Order Number Link';
    PageType = List;
    SourceTable = "Precom Order Number Link";
    UsageCategory = Lists;

    layout
    {
        area(content)
        {
            repeater(General)
            {
                field("No."; Rec."No.")
                {
                    ToolTip = 'Specifies the value of the No. field.';
                    ApplicationArea = All;
                }
                field("Precom No."; Rec."Precom No.")
                {
                    ToolTip = 'Specifies the value of the Precom No. field.';
                    ApplicationArea = All;
                }
            }
        }
    }
}
