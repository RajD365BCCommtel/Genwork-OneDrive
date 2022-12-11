Codeunit 50203 "Onedrive JobQueue Handler"
{
    TableNo = "Job Queue Entry";
    trigger OnRun()
    var
        _OneDriveOutboundLogEntries: record "OneDrive Outbound Log Entries";
        _OneDriveMgt: Codeunit "One Drive Mgt.";
        _OneDriveStream: InStream;
        _IsExecuted: Boolean;
        _ConnectorSetup: Record "Onedrive Connector Setup";
        _FileName: Text;
    BEGIN
        ClearLastError();
        CLEAR(JobQueueExectuer);
        _IsExecuted := FALSE;
        JQEParameter := UPPERCASE(Rec."Parameter String");
        CASE JQEParameter OF
            JobQueueExectuer.ProcessStockDetails():
                begin
                    DeleteTempStockRecords();
                    Clear(_OneDriveMgt);
                    _ConnectorSetup.Get();
                    //_FileName := _ConnectorSetup."File Name" + '_' + Format(Today, 0, '<Day,2><Month,2><Year4>') + '.xlsx';
                    _FileName := _ConnectorSetup."File Name" + '.xlsx';
                    FileName := _FileName;
                    _OneDriveMgt.SetFileName(_FileName);
                    _IsExecuted := _OneDriveMgt.RUN;
                    IF _IsExecuted THEN begin
                        _OneDriveMgt.GetTempClosingStockEntries(TempClosingStockDetails);
                        ProcessTempEntries(TempClosingStockDetails);
                    end else
                        FailureStockDetailEntry(_FileName);
                end;
            ELSE
                ERROR(STRSUBSTNO(JobQueueExectuer.GetUnknownParameter, JQEParameter));
        END;
    END;

    local procedure DeleteTempStockRecords()
    begin
        TempClosingStockDetails.RESET;
        TempClosingStockDetails.DeleteAll();
    end;

    procedure ProcessTempEntries(var _TempClosingStockDetails: Record "Closing Stock Details Buffer" temporary)
    var
        _JobQueueExectuer: Codeunit "Onedrive Jobque. Executer";
        _ClosingStockDetails: Record "Closing Stock Details";
        _CounterAll: Integer;
        _Counter: Integer;
        _CounterOK: Integer;
        _CounterError: Integer;
        _Window: Dialog;
        _Text001: TextConst ENU = 'Scanning and Preparing Closing Stock Details. Please wait...\';
        _Text002: TextConst ENU = '@@@@@@@@@@@1@@@@@@@@@@@@';
        _Text003: TextConst ENU = 'Do you really want to process all Documents?';
        _Text004: TextConst ENU = 'Process finished. %1 Successful, %2 with Error.';
    begin
        IF GUIALLOWED THEN BEGIN
            IF NOT CONFIRM(_Text003) THEN
                EXIT;
            _Window.OPEN(_Text001 + _Text002);
        END;
        TempItem.RESET;
        TempItem.DeleteAll();
        _CounterAll := _TempClosingStockDetails.COUNT + 1;
        _Counter := 1;
        IF _TempClosingStockDetails.FINDSET(TRUE, FALSE) THEN begin
            REPEAT
                _Counter += 1;
                IF GUIALLOWED THEN
                    _Window.UPDATE(1, ROUND(10000 / _CounterAll * _Counter, 1));
                CLEAR(_JobQueueExectuer);
                _JobQueueExectuer.SetJQEParameter(JQEParameter);
                IF _JobQueueExectuer.ProcessTempStockEntry(_TempClosingStockDetails) THEN
                    _CounterOK += 1
                ELSE
                    _CounterError += 1;
            UNTIL _TempClosingStockDetails.NEXT = 0;

            _ClosingStockDetails.RESET;
            _ClosingStockDetails.SetRange("Entry Date", Today());
            IF _ClosingStockDetails.FINDSET then
                repeat
                    UpdateItemFromClosingStock(_ClosingStockDetails);
                until _ClosingStockDetails.Next = 0;
        end;

        //CreateAndSendEmail(_CounterOK, _CounterError);
        DeleteTempStockRecords();
        CLEAR(_JobQueueExectuer);
        _JobQueueExectuer.SetJQEParameter(txtSendEmailAndCopyDelete);
        _JobQueueExectuer.SetFileName(FileName);
        _JobQueueExectuer.SetCounterParameters(_CounterOK, _CounterError);
        IF _JobQueueExectuer.ProcessSendEmailAndCopyDelete(_TempClosingStockDetails) THEN
            IF GUIALLOWED THEN
                MESSAGE('Mail sent and File moved to Processed folder')
            else
                FailureStockDetailEntry(FileName);

        IF GUIALLOWED THEN BEGIN
            _Window.CLOSE;
            MESSAGE(_Text004, _CounterOK, _CounterError);
        END;
    END;

    local procedure UpdateItemFromClosingStock(_ClosingStockDetails: Record "Closing Stock Details")
    var
        _ClosingStockDetails1: Record "Closing Stock Details";
        _StockQty: Decimal;
        _Item: Record Item;
        _PriceListLine: Record "Price List Line";
    begin
        _StockQty := 0;
        IF NOT TempItem.GET(_ClosingStockDetails."Item No.") then begin
            TempItem."No." := _ClosingStockDetails."Item No.";
            TempItem.insert;
            _ClosingStockDetails1.RESET;
            _ClosingStockDetails1.SETRANGE("Item No.", _ClosingStockDetails."Item No.");
            _ClosingStockDetails1.SetRange("Entry Date", Today());
            _ClosingStockDetails1.SetFilter("Entry No.", '<>%1', _ClosingStockDetails."Entry No.");
            IF _ClosingStockDetails1.FINDSET then
                repeat
                    _StockQty += _ClosingStockDetails1."Stock Quantity";
                until _ClosingStockDetails1.Next = 0;
            _StockQty += _ClosingStockDetails."Stock Quantity";
            _Item.get(_ClosingStockDetails."Item No.");
            _Item."Stock Remaining Quantity" := _StockQty;
            _Item.Modify(True);
            _PriceListLine.Reset;
            _PriceListLine.SetCurrentKey("Asset Type", "Asset No.", "Source Type", "Source No.", "Starting Date", "Currency Code", "Variant Code", "Unit of Measure Code", "Minimum Quantity");
            _PriceListLine.SETRANGE("Asset Type", _PriceListLine."Asset Type"::Item);
            _PriceListLine.SETRANGE("Asset No.", _ClosingStockDetails."Item No.");
            _PriceListLine.SetRange(Status, _PriceListLine.Status::Active);
            IF _PriceListLine.FINDSET() THEN
                repeat
                    _PriceListLine."Stock Remaining Quantity" := _StockQty;
                    _PriceListLine.Modify(true);
                until _PriceListLine.next = 0;
            commit;
        end;
    end;




    local procedure FailureStockDetailEntry(FileName: Text)
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
        OneDriveOutboundLogEntries."Processed File" := FileName;
        OneDriveOutboundLogEntries.Modify(true);
        COMMIT;
    end;

    PROCEDURE SetJQEParameter(_JQEParameter: Text[250]);
    BEGIN
        JQEParameter := _JQEParameter;
    END;

    procedure SetFileName(_FileName: Text)
    begin
        FileName := _FileName;
    end;

    VAR

        ExcelBuffer: Record "Excel Buffer" temporary;
        TempClosingStockDetails: Record "Closing Stock Details Buffer" temporary;
        TempItem: Record Item temporary;
        JobQueueExectuer: Codeunit "Onedrive Jobque. Executer";
        JQEParameter: Text[250];
        txtSendEmailAndCopyDelete: TextConst ENU = 'EMAILANDDELETE';
        FileName: Text;

}