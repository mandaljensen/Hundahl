pageextension 50500 "Service Order Quick Extension" extends "Service Order (Quick)"
{
    actions
    {
        addlast("F&unctions")
        {
            action(SendToPrecom)
            {
                Caption = 'Send to Precom';
                ApplicationArea = All;
                ToolTip = 'Send Service Order to Precom';
                Image = TransferOrder;
                Promoted = true;
                PromotedCategory = Process;
                PromotedIsBig = true;

                trigger OnAction()
                var
                    PreComUpdateManagement: Codeunit "PreCom Update Management";
                begin
                    PreComUpdateManagement.WriteServiceOrder(Rec, false);
                end;
            }
        }
    }
}
