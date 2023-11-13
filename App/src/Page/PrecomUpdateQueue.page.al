/// <summary>
/// K2M/14012022/HMJ  Page shows entries in Precom Queue.
/// </summary>
page 50501 "Precom Update Queue"
{
    PageType = List;
    SourceTable = "PreCom Update Queue";
    Caption = 'Precom Update Queue';
    ApplicationArea = All;
    UsageCategory = Lists;

    layout
    {
        area(content)
        {
            repeater(Group)
            {
                field(RecordID; RecordID)
                {
                    ToolTip = 'Specifies the value of the RecordID field.';
                }
                field("Table ID"; "Table ID")
                {
                    ToolTip = 'Specifies the value of the Table ID field.';
                }
                field("Command Type"; "Command Type")
                {
                    ToolTip = 'Specifies the value of the Command Type field.';
                }
                field("Process Error"; "Process Error")
                {
                    ToolTip = 'Specifies the value of the Process Error field.';
                }
                field("Error Text"; "Error Text")
                {
                    ToolTip = 'Specifies the value of the Error Text field.';
                }
            }
        }
    }

    actions
    {
        /*         area(Processing)
                {
                    action(ProcessQueue)
                    {
                        Caption = 'Process Queue';
                        ApplicationArea = All;
                        RunObject = codeunit "PreCom Update Dispatcher";
                    }
                } */
    }
}

