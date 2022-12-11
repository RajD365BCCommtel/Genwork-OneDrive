Codeunit 50204 "Onedrive Jobque. Executer"
{
    TableNo = "Closing Stock Details Buffer";
    Trigger OnRun()
    BEGIN
        CASE JQEParameter OF
            ProcessStockDetails():
                InsertDataFromTempStock(Rec);
            SendEmailAndCopyDelete():
                EmailCopyAndDeleteFile();
            ELSE
                ERROR(STRSUBSTNO(GetUnknownParameter, JQEParameter));
        END;
    END;

    local procedure InsertDataFromTempStock(var TempClosingStockDetails: Record "Closing Stock Details Buffer" temporary)
    var
        ClosingStockDetails: Record "Closing Stock Details";
        EntryNo: Integer;
    begin
        CheckMasterAvailable(TempClosingStockDetails);
        EntryNo := 0;
        ClosingStockDetails.Reset();
        if ClosingStockDetails.FindLast() then
            EntryNo := ClosingStockDetails."Entry No.";
        IF NOT ClosingStockDetails.GET(TempClosingStockDetails."Entry Date", TempClosingStockDetails."Item No.", TempClosingStockDetails."Location Code", TempClosingStockDetails."Bin Code") THEN begin
            ClosingStockDetails.Init();
            EntryNo := EntryNo + 1;
            ClosingStockDetails."Entry No." := EntryNo;
            ClosingStockDetails."Item No." := TempClosingStockDetails."Item No.";
            ClosingStockDetails."Location Code" := TempClosingStockDetails."Location Code";
            ClosingStockDetails."Bin Code" := TempClosingStockDetails."Bin Code";
            ClosingStockDetails."Stock Quantity" := TempClosingStockDetails."Stock Quantity";
            ClosingStockDetails."Entry Date" := TempClosingStockDetails."Entry Date";
            ClosingStockDetails.Insert();
        end else begin
            ClosingStockDetails."Stock Quantity" := TempClosingStockDetails."Stock Quantity";
            ClosingStockDetails.Modify();
        end;
    end;

    Local procedure EmailCopyAndDeleteFile()
    var
        OneDriveMgt: Codeunit "One Drive Mgt.";
    begin
        Clear(OneDriveMgt);
        OneDriveMgt.EmailCopyAndDeleteOnedriveFile(FileName, CounterOK, CounterError);
    end;

    local procedure CheckMasterAvailable(var TempClosingStockDetails: Record "Closing Stock Details Buffer" temporary)
    var
        Item: Record Item;
        Location: Record Location;
        Bin: Record Bin;
        Result: Text;
        ItemNotFoundErr: Label 'Item %1 does not exit in Item Master.';
        ItemNoValueEmptyErr: Label 'Item No is empty in current excel sheet.';
        LocationNotFoundErr: Label 'Location %1 does not exit in Location Master.';
        BinNotFoundErr: Label 'Location %1 with Bin %2 does not exit in Bin Master.';
    begin
        Result := '';
        IF TempClosingStockDetails."Item No." = '' then
            Result := ItemNoValueEmptyErr
        else begin
            IF Not Item.GET(TempClosingStockDetails."Item No.") then
                Result := StrSubstNo(ItemNotFoundErr, TempClosingStockDetails."Item No.")
        end;

        IF TempClosingStockDetails."Location Code" <> '' THEN
            IF Not Location.GET(TempClosingStockDetails."Location Code") then begin
                IF Result = '' THEN
                    Result := StrSubstNo(LocationNotFoundErr, TempClosingStockDetails."Location Code")
                else
                    Result := Result + '' + StrSubstNo(LocationNotFoundErr, TempClosingStockDetails."Location Code")

            end;
        IF TempClosingStockDetails."Bin Code" <> '' THEN
            IF TempClosingStockDetails."Location Code" <> '' THEN BEGIN
                IF Not Bin.Get(TempClosingStockDetails."Location Code", TempClosingStockDetails."Bin Code") then begin
                    IF Result = '' THEN
                        Result := StrSubstNo(BinNotFoundErr, TempClosingStockDetails."Location Code", TempClosingStockDetails."Bin Code")
                    else
                        Result := Result + '' + StrSubstNo(BinNotFoundErr, TempClosingStockDetails."Location Code", TempClosingStockDetails."Bin Code");
                end;
            end;
        IF Result <> '' then
            error(Result);
    end;

    PROCEDURE ProcessTempStockEntry(var TempClosingStockDetails: Record "Closing Stock Details Buffer" temporary): Boolean;
    VAR
        _OneDriveOutboundLogEntries2: Record "OneDrive Outbound Log Entries";
        _JobQueueExecuter: Codeunit "Onedrive Jobque. Executer";
    BEGIN
        COMMIT;
        _JobQueueExecuter.SetJQEParameter(JQEParameter);
        _JobQueueExecuter.SetFileName(FileName);
        IF _JobQueueExecuter.RUN(TempClosingStockDetails) THEN BEGIN
            //CreateOutboundEntries();
            COMMIT;
            EXIT(TRUE);
        END ELSE BEGIN
            FailureOutboudEntryForStockDetail(FileName, TempClosingStockDetails."Excel Line No.");
            COMMIT;
            EXIT(FALSE);
        END;
    END;

    PROCEDURE ProcessSendEmailAndCopyDelete(var TempClosingStockDetails: Record "Closing Stock Details Buffer" temporary): Boolean;
    VAR
        _OneDriveOutboundLogEntries2: Record "OneDrive Outbound Log Entries";
        _JobQueueExecuter: Codeunit "Onedrive Jobque. Executer";
    BEGIN
        COMMIT;
        ClearLastError();
        _JobQueueExecuter.SetJQEParameter(JQEParameter);
        _JobQueueExecuter.SetFileName(FileName);
        _JobQueueExecuter.SetCounterParameters(CounterOK, CounterError);
        IF _JobQueueExecuter.RUN(TempClosingStockDetails) THEN
            EXIT(TRUE)
        ELSE
            EXIT(FALSE);
    END;



    local procedure CreateOutboundEntries()
    var
        OneDriveOutboundLogEntries: Record "OneDrive Outbound Log Entries";
    begin
        OneDriveOutboundLogEntries.INIT;
        OneDriveOutboundLogEntries.INSERT(TRUE);
        OneDriveOutboundLogEntries."Created DateTime" := CURRENTDATETIME;
        OneDriveOutboundLogEntries."Response Status" := OneDriveOutboundLogEntries."Response Status"::Successful;
        OneDriveOutboundLogEntries.Modify(true);
        COMMIT;
    end;

    local procedure FailureOutboudEntryForStockDetail(_fileName: Text[100]; _RowNo: Integer)
    var
        OneDriveOutboundLogEntries: Record "OneDrive Outbound Log Entries";
    begin
        OneDriveOutboundLogEntries.INIT;
        OneDriveOutboundLogEntries.INSERT(TRUE);
        OneDriveOutboundLogEntries.Description := 'Closing Stock Detail';
        OneDriveOutboundLogEntries."Created DateTime" := CURRENTDATETIME;
        OneDriveOutboundLogEntries."Response Status" := OneDriveOutboundLogEntries."Response Status"::Unsuccessful;
        OneDriveOutboundLogEntries."Has Error" := TRUE;
        OneDriveOutboundLogEntries."Error Message" := COPYSTR(GETLASTERRORTEXT, 1, MAXSTRLEN(OneDriveOutboundLogEntries."Error Message"));
        OneDriveOutboundLogEntries."Processed File" := _fileName;
        OneDriveOutboundLogEntries."Excel Row Number" := _RowNo;
        OneDriveOutboundLogEntries.Modify(true);
        COMMIT;
    end;

    PROCEDURE SetJQEParameter(_JQEParameter: Text[250]);
    BEGIN
        JQEParameter := _JQEParameter;
    END;

    PROCEDURE SetCounterParameters(_CounterOK: Integer; _CounterError: Integer);
    BEGIN
        CounterOK := _CounterOK;
        CounterError := _CounterError;
    END;

    PROCEDURE CreateOrderPayload(): Text[250];
    BEGIN
        EXIT(txtCreateOrderPayload);
    END;

    PROCEDURE ProcessStockDetails(): Text[250];
    BEGIN
        EXIT(txtProcessStockDetails);
    END;

    PROCEDURE SendEmailAndCopyDelete(): Text[250];
    BEGIN
        EXIT(txtSendEmailAndCopyDelete);
    END;

    procedure SetFileName(_FileName: Text)
    begin
        FileName := _FileName;
    end;


    procedure GetUnknownParameter(): Text[250];
    BEGIN
        EXIT(txtUnknownParameter);
    END;

    VAR
        txtCreateOrderPayload: TextConst ENU = 'CREATEORDERPAYLOAD';
        txtProcessStockDetails: TextConst ENU = 'PROCESSSTOCKDETAILS';
        txtSendEmailAndCopyDelete: TextConst ENU = 'EMAILANDDELETE';
        txtUnknownParameter: TextConst ENU = 'Job Queue parameter %1 is unknown!;';
        JQEParameter: Text[250];
        FileName: Text;
        CounterOK: Integer;
        CounterError: Integer;

}