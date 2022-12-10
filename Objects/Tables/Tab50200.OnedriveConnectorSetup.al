//Created: 04/07/22 by Rajesh(rajesh.gan@tyreconnect.com.au)
//Description: This ia a new Table "TC Saleforce Integration Setup"
table 50200 "Onedrive Connector Setup"
{
    Caption = 'Onedrive Connector Setup';
    DataClassification = ToBeClassified;
    DrillDownPageId = "Onedrive Connector Setup";
    LookupPageId = "Onedrive Connector Setup";

    fields
    {
        field(1; "Primary Key"; Code[10])
        {
            Caption = 'Primary Key';
        }
        field(2; "Client ID"; Text[250])
        {
            Caption = 'Client ID';
            DataClassification = EndUserIdentifiableInformation;
        }
        field(3; "Client Secret"; Text[250])
        {
            Caption = 'Client Secret';
            DataClassification = EndUserIdentifiableInformation;
        }
        field(4; "Redirect URL"; Text[250])
        {
            Caption = 'Redirect URL';
        }
        field(5; Scope; Text[250])
        {
            Caption = 'Scope';
        }
        field(6; "Authorization URL"; Text[250])
        {
            Caption = 'Authorization URL';

            trigger OnValidate()
            var
                WebRequestHelper: Codeunit "Web Request Helper";
            begin
                if "Authorization URL" <> '' then
                    WebRequestHelper.IsSecureHttpUrl("Authorization URL");
            end;
        }
        field(7; "Access Token URL"; Text[250])
        {
            Caption = 'Access Token URL';

            trigger OnValidate()
            var
                WebRequestHelper: Codeunit "Web Request Helper";
            begin
                if "Access Token URL" <> '' then
                    WebRequestHelper.IsSecureHttpUrl("Access Token URL");
            end;
        }
        field(9; "Access Token"; Blob)
        {
            Caption = 'Access Token';
            DataClassification = EndUserIdentifiableInformation;
        }
        field(10; "Refresh Token"; Blob)
        {
            Caption = 'Refresh Token';
            DataClassification = EndUserIdentifiableInformation;
        }
        field(13; "Authorization Time"; DateTime)
        {
            Caption = 'Authorization Time';
            Editable = false;
            DataClassification = EndUserIdentifiableInformation;
        }
        field(14; "Expires In"; Integer)
        {
            Caption = 'Expires In';
            Editable = false;
            DataClassification = EndUserIdentifiableInformation;
        }
        field(15; "Ext. Expires In"; Integer)
        {
            Caption = 'Ext. Expires In';
            Editable = false;
            DataClassification = EndUserIdentifiableInformation;
        }
        field(16; "Drive ID"; Text[50])
        {
            Caption = 'Drive ID';
        }
        field(17; "Folder Name"; Text[50])
        {
            Caption = 'Folder Name';
        }
        field(18; Status; Option)
        {
            Caption = 'Status';
            OptionCaption = ' ,Enabled,Disabled,Connected,Error';
            OptionMembers = " ",Enabled,Disabled,Connected,Error;
        }
        field(19; "Import Folder"; Text[50])
        {
            Caption = 'Import Folder';
        }
        field(20; "File Name"; Text[50])
        {
            Caption = 'File Name';
        }
        field(21; "Move Folder Id"; Text[50])
        {
            Caption = 'Move Folder Id';
        }
    }

    keys
    {
        key(PK; "Primary Key")
        {
            Clustered = true;
        }
    }


    procedure SaveAccessToken(var MessageText: Text) Result: Boolean
    var
        Processed: Boolean;
    begin
        if not Processed then
            Result := AccessTokenMgmnt.RefreshAndSaveAccessToken(Rec, MessageText);
    end;

    local procedure CheckAndAppendURLPath(var value: Text)
    begin
        if value <> '' then
            if value[1] <> '/' then
                value := '/' + value;
    end;

    procedure SetAccessToken(NewToken: Text)
    var
        OutStream: OutStream;
    begin
        Clear("Access Token");
        "Access Token".CreateOutStream(OutStream, TEXTENCODING::UTF8);
        OutStream.WriteText(NewToken);
        Modify();
    end;

    procedure GetAccessToken() AccessToken: Text
    var
        TypeHelper: Codeunit "Type Helper";
        InStream: InStream;
    begin
        CalcFields("Access Token");
        "Access Token".CreateInStream(InStream, TEXTENCODING::UTF8);
        if not TypeHelper.TryReadAsTextWithSeparator(InStream, TypeHelper.LFSeparator(), AccessToken) then Message(ReadingDataSkippedMsg, FieldCaption("Access Token"));
    end;


    Var
        TokenLbl: Label 'token';
        AccessTokenMgmnt: Codeunit "OneDrive Access Token Mgmnt.";
        ReadingDataSkippedMsg: Label 'Loading field %1 will be skipped because there was an error when reading the data.\To fix the current data, contact your administrator.\Alternatively, you can overwrite the current data by entering data in the field.', Comment = '%1=field caption';

}

