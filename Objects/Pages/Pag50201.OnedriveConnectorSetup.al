page 50201 "Onedrive Connector Setup"
{
    Caption = 'Onedrive Connector Setup';
    PageType = Card;
    SourceTable = "Onedrive Connector Setup";
    UsageCategory = Administration;
    ApplicationArea = ALL;

    layout
    {
        area(Content)
        {
            group(GroupName)
            {
                field(ClientID; Rec."Client ID")
                {
                    ApplicationArea = All;
                }
                field(ClientSecret; Rec."Client Secret")
                {
                    ApplicationArea = All;
                }
                field("Redirect URL"; Rec."Redirect URL")
                {
                    ApplicationArea = All;
                }
                field("Authorization URL Path"; Rec."Authorization URL")
                {
                    ApplicationArea = All;
                }
                field("Access Token URL Path"; Rec."Access Token URL")
                {
                    ApplicationArea = All;
                }
                field("Drive ID"; Rec."Drive ID")
                {
                    ApplicationArea = All;
                }
                field("Folder Name"; Rec."Folder Name")
                {
                    ApplicationArea = All;
                }
                field("Import Folder"; Rec."Import Folder")
                {
                    ApplicationArea = All;
                }
                field(Scope; Rec.Scope)
                {
                    ApplicationArea = All;
                }
                field("File Name"; Rec."File Name")
                {
                    ApplicationArea = All;
                }
                field("Move Folder Id"; Rec."Move Folder Id")
                {
                    ApplicationArea = All;
                }
                group("New Access Token")
                {
                    Caption = 'Access Token';

                    field(NewAccessToken; AccessToken)
                    {
                        ApplicationArea = Basic, Suite;
                        Importance = Additional;
                        MultiLine = true;
                        ShowCaption = false;
                        Editable = false;

                        ToolTip = 'Specifies the value of the Access Token field.';

                        trigger OnValidate()
                        begin
                            Rec.SetAccessToken(AccessToken);
                        end;
                    }
                }
            }
        }
    }

    actions
    {
        area(processing)
        {
            action(AccessToken)
            {
                ApplicationArea = Basic, Suite;
                Caption = 'Refresh and Save Access Token';
                Image = Refresh;
                Promoted = true;
                PromotedCategory = Process;
                PromotedOnly = true;
                ToolTip = 'Refresh and Save Access Token.';

                trigger OnAction()
                var
                    MessageText: Text;
                begin
                    if not Rec.SaveAccessToken(MessageText) then begin
                        Commit(); // save new "Status" value
                        Error(MessageText);
                    end;

                    Message(MessageText);
                end;
            }
        }
    }

    trigger OnOpenPage()
    begin
        Rec.reset;
        IF Rec.IsEmpty THEN begin
            Rec.Init();
            rec.insert();
        end;
    end;

    trigger OnAfterGetRecord()
    begin
        AccessToken := Rec.GetAccessToken();
    end;

    var
        AccessToken: Text;
}
