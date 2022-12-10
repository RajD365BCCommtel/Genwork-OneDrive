//Description: This ia a new table "OneDrive Outbound Log Entries"
table 50201 "OneDrive Outbound Log Entries"
{
    Caption = 'OneDrive Outbound Log Entries';
    LookupPageId = "Onedrive Outbound Log Entries";
    DrillDownPageId = "Onedrive Outbound Log Entries";
    DataClassification = ToBeClassified;
    fields
    {
        field(1; "Entry No."; Integer)
        {
            Caption = 'Entry No.';
            DataClassification = ToBeClassified;
        }
        field(2; "Created DateTime"; DateTime)
        {
            DataClassification = ToBeClassified;
        }
        field(3; "Error Message"; text[250])
        {
            DataClassification = ToBeClassified;
        }
        field(4; "Response Status"; Option)
        {
            OptionMembers = ,Successful,Unsuccessful;
            DataClassification = ToBeClassified;
        }
        field(5; "Description"; Text[100])
        {
            DataClassification = ToBeClassified;
        }
        field(6; "Processed DateTime"; DateTime)
        {
            DataClassification = ToBeClassified;
        }
        field(7; active; Boolean)
        {
            DataClassification = ToBeClassified;
        }
        field(8; "Has Error"; Boolean)
        {
            DataClassification = ToBeClassified;
        }
        field(9; "Processed File"; text[100])
        {
            Caption = 'Preocessed File Name';
            DataClassification = ToBeClassified;
        }
        field(10; "Excel Row Number"; integer)
        {
            DataClassification = ToBeClassified;
        }
    }
    keys
    {
        key(PK; "Entry No.")
        {
            Clustered = true;
        }
    }
    trigger OnInsert()
    begin
        IF "Entry No." = 0 THEN
            "Entry No." := GetNextEntryNo;
    end;

    PROCEDURE GetNextEntryNo() _NextEntryNo: Integer;
    VAR
        _OneDriveOutboundLogEntries: Record "OneDrive Outbound Log Entries";
    BEGIN
        _OneDriveOutboundLogEntries.LOCKTABLE;
        IF _OneDriveOutboundLogEntries.FINDLAST THEN
            _NextEntryNo := _OneDriveOutboundLogEntries."Entry No." + 1
        ELSE
            _NextEntryNo := 1;
    END;

    procedure DeleteEntries(DaysOld: Integer)
    begin
        If not Confirm(Text001) then
            exit;
        Window.open(DeletingMsg);
        IF DaysOld > 0 then
            SETFILTER("Created DateTime", '<=%1', CreateDateTime((TODAY - DaysOld), 0T));
        DeleteAll();
        Window.Close();
        SetRange("Created DateTime");
        Message(DeletedMsg);
    end;

    var
        Text001: TextConst ENU = 'you sure that you want to delete Onedrive Outbound Log entries?';
        DeletingMsg: TextConst ENU = 'Deleting Entries...';
        DeletedMsg: TextConst ENU = 'Entries have been deleted.';
        Window: Dialog;
}
