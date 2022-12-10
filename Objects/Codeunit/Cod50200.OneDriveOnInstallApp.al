codeunit 50200 "OneDrive OnInstall App"
{
    Subtype = Install;
    trigger OnRun()
    begin

    end;

    trigger OnInstallAppPerCompany();
    begin
        InsertIntegrationSetup();
        CreateJobQueue();
    end;

    local procedure InsertIntegrationSetup()
    var
        _ConnectorSetup: Record "Onedrive Connector Setup";
        _AccessTokenURLtxt: TextConst ENU = 'https://login.microsoftonline.com/a007e963-f6fd-40bf-87d6-0320fb2e8432/oauth2/v2.0/token';
        _AuthTokenURLtxt: TextConst ENU = 'https://login.microsoftonline.com/a007e963-f6fd-40bf-87d6-0320fb2e8432/oauth2/v2.0/authorize';
        _ClientIDtxt: TextConst ENU = 'ea713f53-2720-41aa-a5b9-010b92637c36';
        _ClientSecrettxt: TextConst ENU = '5KO8Q~kTe36v54qK1I5sc5qil.tFaM5swBlhwcC_';
        _Scopetxt: TextConst ENU = 'https://graph.microsoft.com/.default offline_access';
        _GrantTypetxt: TextConst ENU = 'Authorization Code';
        _RedirectURLtxt: TextConst ENU = 'https://businesscentral.dynamics.com/a007e963-f6fd-40bf-87d6-0320fb2e8432/oauthlanding.htm';
        _DriveId: TextConst ENU = 'f38cac7b-d96b-4ad7-8717-588bacd89bd4';//'09686f24-69c9-4a55-a604-4e7e86183acd';
        _FolderName: TextConst ENU = 'DigitalOrders';//'DigitalOrders';
        _ImportFolder: TextConst ENU = 'ItemInventoryStatus';//ItemInventoryStatus
        _FileName: TextConst ENU = 'ConsolidatedBatchReport';
        _MoveFolderID: TextConst ENU = '01MF7ZTLVWS2QOVP6WWRGISIW33MEONIIO';//MoveFolderID
    begin
        IF NOT _ConnectorSetup.GET THEN BEGIN
            _ConnectorSetup.Init();
            _ConnectorSetup.INSERT(TRUE);
        END;
        _ConnectorSetup.VALIDATE("Access Token URL", _AccessTokenURLtxt);
        _ConnectorSetup."Authorization URL" := _AuthTokenURLtxt;
        _ConnectorSetup."Redirect URL" := _RedirectURLtxt;
        _ConnectorSetup."Client ID" := _ClientIDtxt;
        _ConnectorSetup."Client Secret" := _ClientSecrettxt;
        _ConnectorSetup.Scope := _Scopetxt;
        _ConnectorSetup."Drive ID" := _DriveId;
        _ConnectorSetup."Folder Name" := _FolderName;
        _ConnectorSetup."Import Folder" := _ImportFolder;
        _ConnectorSetup."File Name" := _FileName;
        _ConnectorSetup."Move Folder Id" := _MoveFolderID;
        _ConnectorSetup.Modify(TRUE);
    end;

    local procedure CreateJobQueue();
    var
        _JobQueueEntry: Record "Job Queue Entry";
    begin
        WITH _JobQueueEntry DO BEGIN
            Reset();
            SetRange("Object ID to Run", Report::"OneDrive Sales Order Report");
            SETRANGE("Object Type to Run", "Object Type to Run"::Report);
            IF not FINDFIRST THEN begin
                INIT;
                VALIDATE("Object Type to Run", "Object Type to Run"::Report);
                VALIDATE("Object ID to Run", Report::"OneDrive Sales Order Report");
                Validate(Description, 'Store Sales Order report to One Drive');
                VALIDATE("Run on Mondays", TRUE);
                VALIDATE("Run on Tuesdays", TRUE);
                VALIDATE("Run on Wednesdays", TRUE);
                VALIDATE("Run on Thursdays", TRUE);
                VALIDATE("Run on Fridays", TRUE);
                VALIDATE("Run on Saturdays", TRUE);
                VALIDATE("Run on Sundays", TRUE);
                //VALIDATE("Starting Time", 230000T);
                VALIDATE("No. of Minutes between Runs", 5);
                VALIDATE("Maximum No. of Attempts to Run", 3);
                INSERT(TRUE);
                SetStatus(Status::"On Hold");
            END;
        END;
    end;
}
