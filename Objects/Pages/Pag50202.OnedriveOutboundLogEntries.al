page 50202 "Onedrive Outbound Log Entries"
{
    ApplicationArea = All;
    Caption = 'Onedrive Outbound Log Entries';
    PageType = List;
    SourceTable = "OneDrive Outbound Log Entries";
    UsageCategory = Lists;
    Editable = false;
    InsertAllowed = false;
    ModifyAllowed = false;
    DeleteAllowed = False;
    layout
    {
        area(content)
        {
            repeater(General)
            {
                field("Entry No."; Rec."Entry No.")
                {
                    ApplicationArea = All;
                }
                field("Description"; Rec.Description)
                {
                    ApplicationArea = All;
                }
                field("Has Error"; Rec."Has Error")
                {
                    ApplicationArea = All;
                }
                field("Response Status"; Rec."Response Status")
                {
                    ApplicationArea = All;
                }

                field("Error Message"; Rec."Error Message")
                {
                    ApplicationArea = All;
                }
                field("Created DateTime"; Rec."Created DateTime")
                {
                    ApplicationArea = All;
                }
                field("Processed DateTime"; Rec."Processed DateTime")
                {
                    ApplicationArea = All;
                }
                field("Processed File"; Rec."Processed File")
                {
                    ApplicationArea = All;

                }
                field("Excel Row Number"; Rec."Excel Row Number")
                {
                    ApplicationArea = All;

                }

            }
        }
    }
    actions
    {
        area(processing)
        {
            action(Process)
            {
                ApplicationArea = All;
                Caption = 'Process';
                Promoted = true;
                PromotedIsBig = true;
                Image = Process;
                PromotedCategory = Process;
                trigger OnAction()
                Var
                    _Text000: TextConst ENU = 'Process all, Process only selected';
                    _txtProcessPaylodEntry: TextConst ENU = 'PROCESSPAYLOADENTRY';
                    _OneDriveMgt: Codeunit "One Drive Mgt.";
                    _Selection: Integer;
                begin
                    // _Selection := STRMENU(_Text000, 2);
                    // CASE _Selection OF
                    //     0:
                    //         EXIT;
                    //     1:
                    //         _TCOutboundCommEntries.COPYFILTERS(Rec);
                    //     2:
                    //         BEGIN
                    //             _TCOutboundCommEntries.COPY(Rec);
                    //             CurrPage.SETSELECTIONFILTER(_TCOutboundCommEntries);
                    //         END;
                    // END;
                    _OneDriveMgt.ManualProcess();
                    CurrPage.UPDATE(TRUE);
                end;
            }

            Group("Delete Log Entries")
            {
                action(Delete7Days)
                {
                    ApplicationArea = Basic, Suite;
                    Caption = 'Delete Entries Older Than 7 Days';
                    Ellipsis = true;
                    Image = ClearLog;
                    Promoted = true;
                    PromotedCategory = Process;
                    PromotedOnly = true;
                    ToolTip = 'Clear the list of log enties that are older than 7 days.';
                    trigger OnAction()
                    begin
                        Rec.DeleteEntries(7);
                    end;
                }
                action(Delete60Days)
                {
                    ApplicationArea = Basic, Suite;
                    Caption = 'Delete Entries Older Than 60 Days';
                    Ellipsis = true;
                    Image = ClearLog;
                    Promoted = true;
                    PromotedCategory = Process;
                    PromotedOnly = true;
                    ToolTip = 'Clear the list of log enties that are older than 60 days.';
                    trigger OnAction()
                    begin
                        Rec.DeleteEntries(60);
                    end;
                }
                action(Delete0Days)
                {
                    ApplicationArea = Basic, Suite;
                    Caption = 'Delete All Entries';
                    Ellipsis = true;
                    Image = ClearLog;
                    Promoted = true;
                    PromotedCategory = Process;
                    PromotedOnly = true;
                    ToolTip = 'Clear the list of all log enties.';
                    trigger OnAction()
                    begin
                        Rec.DeleteEntries(0);
                    end;
                }
            }
        }
    }
}
