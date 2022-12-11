codeunit 50201 "One Drive Mgt."
{
    trigger OnRun()

    begin
        LoadData();
    end;

    procedure LoadData()
    var
        OneDriveStream: InStream;
        ConnectorSetup: Record "Onedrive Connector Setup";
        ClosingStockDetailsBuffer: Record "Closing Stock Details Buffer";
    begin
        ClosingStockDetailsBuffer.RESET;
        ClosingStockDetailsBuffer.DeleteAll();
        ConnectorSetup.Get();
        if ConnectorSetup."File Name" = '' then
            Error('Please Enter the FileName in OneDrive ConnectorSetup.');
        //FileName := ConnectorSetup."File Name" + '_' + Format(Today, 0, '<Day,2><Month,2><Year4>') + '.xlsx';
        FileName := ConnectorSetup."File Name" + '.xlsx';
        GetOnedriveFileStream(FileName, OneDriveStream);
        ReadExcelSheet(OneDriveStream);
        ImportExcelDataToTempStock();
    end;



    procedure GetOnedriveFileStream(OneDriveFile: Text; var Stream: InStream)
    var
        Client: HttpClient;
        Headers: HttpHeaders;
        RequestMessage: HttpRequestMessage;
        ResponseMessage: HttpResponseMessage;
        RequestContent: HttpContent;
        ConnectorSetup: Record "Onedrive Connector Setup";
        _AccessTokenMgmnt: codeunit "OneDrive Access Token Mgmnt.";
        _AccessToken: Text;
        ResponseText: Text;
        _MessageTxt: Text;
        SetRequestURL: Text;
        ErrorMessage: Text;
    begin
        ConnectorSetup.Get();
        ConnectorSetup.TestField("Drive ID");
        ConnectorSetup.TestField("Import Folder");
        _AccessToken := ConnectorSetup.GetAccessToken();
        _AccessTokenMgmnt.InvokeAccessToken(ConnectorSetup, _MessageTxt, _AccessToken, true);
        Headers := Client.DefaultRequestHeaders();
        Headers.Add('Authorization', StrSubstNo('Bearer %1', _AccessToken));
        RequestMessage.SetRequestUri(StrSubstNo('https://graph.microsoft.com/v1.0/drives/' + ConnectorSetup."Drive ID" + '/root:/' + ConnectorSetup."Import Folder" + '/%1:/content', OneDriveFile));
        RequestMessage.Method := 'GET';
        if NOT Client.Send(RequestMessage, ResponseMessage) then
            if ResponseMessage.IsBlockedByEnvironment() then
                ErrorMessage := StrSubstNo(EnvironmentBlocksErr, RequestMessage.GetRequestUri())
            else
                ErrorMessage := StrSubstNo(ConnectionErr, RequestMessage.GetRequestUri());

        if ErrorMessage <> '' then
            Error(ErrorMessage);
        if ResponseMessage.IsSuccessStatusCode() then
            ResponseMessage.Content.ReadAs(Stream)
        else
            ErrorMessage := StrSubstNo('HTTP error %1 (%2)', ResponseMessage.HttpStatusCode(), ResponseMessage.ReasonPhrase());
        if ErrorMessage <> '' then
            Error(ErrorMessage);
    end;

    local procedure ReadExcelSheet(_Stream: InStream)
    var
        SheetName: Text;
    begin
        ExcelBuffer.Reset();
        ExcelBuffer.DeleteAll();
        SheetName := ExcelBuffer.SelectSheetsNameStream(_Stream);
        ExcelBuffer.OpenBookStream(_Stream, SheetName);
        ExcelBuffer.ReadSheet();
    end;

    local procedure ImportExcelDataToTempStock()
    var
        RowNo: Integer;
        ColumnNo: Integer;
        MaxRowNo: Integer;
        EntryNo: Integer;
        _ItemNo: Text;
        _LocationNo: Text;
        _BinCode: Text;
        _StockQty: Decimal;
        _StockQtyTxt: text;
        _EntryDate: Date;
    begin
        EntryNo := 0;
        _EntryDate := Today();
        TempClosingStockDetails.RESET;
        TempClosingStockDetails.DeleteAll();
        ExcelBuffer.Reset();
        if ExcelBuffer.FindLast() then
            MaxRowNo := ExcelBuffer."Row No.";
        for RowNo := 2 to MaxRowNo do begin
            _ItemNo := '';
            _LocationNo := '';
            _BinCode := '';
            _StockQtyTxt := '';
            Clear(_StockQty);
            _ItemNo := GetValueAtCell(RowNo, 3);
            _LocationNo := GetValueAtCell(RowNo, 1);
            _BinCode := GetValueAtCell(RowNo, 5);
            _StockQtyTxt := GetValueAtCell(RowNo, 16);
            IF Evaluate(_StockQty, _StockQtyTxt) then;

            IF NOT TempClosingStockDetails.GET(_EntryDate, _ItemNo, _LocationNo, _BinCode) THEN begin
                TempClosingStockDetails.Init();
                EntryNo := EntryNo + 1;
                TempClosingStockDetails."Entry Date" := _EntryDate;
                TempClosingStockDetails."Entry No." := EntryNo;
                TempClosingStockDetails."Item No." := _ItemNo;
                TempClosingStockDetails."Location Code" := _LocationNo;
                TempClosingStockDetails."Bin Code" := _BinCode;
                TempClosingStockDetails."Stock Quantity" := _StockQty;
                TempClosingStockDetails."Excel Line No." := RowNo;
                TempClosingStockDetails.Insert();
            end else begin
                TempClosingStockDetails."Stock Quantity" := TempClosingStockDetails."Stock Quantity" + _StockQty;
                TempClosingStockDetails.Modify();
            end;
        end;
        IF TempClosingStockDetails.IsEmpty then
            Error(StrSubstNo(ExcelFileDataErr, FileName));
    end;

    local procedure GetValueAtCell(RowNo: Integer; ColNo: Integer): Text;
    begin
        if ExcelBuffer.GET(RowNo, ColNo) then
            exit(ExcelBuffer."Cell Value as Text");
    end;

    procedure ManualProcess()
    var
        OnedriveJobQueueHandler: Codeunit "Onedrive JobQueue Handler";
    begin
        LoadData();
        OnedriveJobQueueHandler.SetJQEParameter('PROCESSSTOCKDETAILS');
        OnedriveJobQueueHandler.SetFileName(FileName);
        OnedriveJobQueueHandler.ProcessTempEntries(TempClosingStockDetails);
    end;

    procedure EmailCopyAndDeleteOnedriveFile(OneDriveFile: Text; _CounterOk: Integer; _CounterErr: Integer)
    var
        Client: HttpClient;
        Headers: HttpHeaders;
        RequestMessage: HttpRequestMessage;
        ResponseMessage: HttpResponseMessage;
        RequestContent: HttpContent;
        RequestContentHeader: HttpHeaders;
        ConnectorSetup: Record "Onedrive Connector Setup";
        _AccessTokenMgmnt: codeunit "OneDrive Access Token Mgmnt.";
        RequestJson: JsonObject;
        ReqJson: JsonObject;
        OneDriveURL: Text;
        _JsonText: Text;
        _AccessToken: Text;
        ResponseText: Text;
        _MessageTxt: Text;
        SetRequestURL: Text;
        ErrorMessage: Text;
    begin
        CreateAndSendEmail(_CounterOk, _CounterErr);
        ConnectorSetup.Get();
        ConnectorSetup.TestField("Drive ID");
        ConnectorSetup.TestField("Import Folder");
        ConnectorSetup.TestField("Move Folder Id");
        _AccessToken := ConnectorSetup.GetAccessToken();
        _AccessTokenMgmnt.InvokeAccessToken(ConnectorSetup, _MessageTxt, _AccessToken, true);
        Headers := Client.DefaultRequestHeaders();
        OneDriveURL := StrSubstNo('https://graph.microsoft.com/v1.0/drives/' + ConnectorSetup."Drive ID" + '/root:/' + ConnectorSetup."Import Folder" + '/%1:/Copy', OneDriveFile);
        RequestJson.Add('name', OneDriveFile);
        RequestJson.Add('parentReference', GetParentReferenceJson(ConnectorSetup."Move Folder Id"));
        RequestJson.WriteTo(_JsonText);
        RequestContent.WriteFrom(_JsonText);
        RequestContent.GetHeaders(RequestContentHeader);
        RequestContentHeader.Clear();
        RequestContentHeader.Add('Content-Type', 'application/json');
        Headers.Add('Authorization', StrSubstNo('Bearer %1', _AccessToken));
        if not Client.post(OneDriveURL, RequestContent, ResponseMessage) then
            if ResponseMessage.IsBlockedByEnvironment() then
                ErrorMessage := StrSubstNo(EnvironmentBlocksErr, RequestMessage.GetRequestUri())
            else
                ErrorMessage := StrSubstNo(ConnectionErr, RequestMessage.GetRequestUri());
        if ErrorMessage <> '' then
            Error(ErrorMessage);
        if ResponseMessage.IsSuccessStatusCode() then
            ErrorMessage := ''
        else
            ErrorMessage := StrSubstNo('HTTP error %1 (%2)', ResponseMessage.HttpStatusCode(), ResponseMessage.ReasonPhrase());
        if ErrorMessage <> '' then
            Error(ErrorMessage);
        DeleteOnedriveFile(OneDriveFile);
    end;

    local procedure DeleteOnedriveFile(OneDriveFile: Text)
    var
        Client: HttpClient;
        Headers: HttpHeaders;
        RequestMessage: HttpRequestMessage;
        ResponseMessage: HttpResponseMessage;
        RequestContent: HttpContent;
        RequestContentHeader: HttpHeaders;
        ConnectorSetup: Record "Onedrive Connector Setup";
        _AccessTokenMgmnt: codeunit "OneDrive Access Token Mgmnt.";
        RequestJson: JsonObject;
        ReqJson: JsonObject;
        OneDriveURL: Text;
        _JsonText: Text;
        _AccessToken: Text;
        ResponseText: Text;
        _MessageTxt: Text;
        SetRequestURL: Text;
        ErrorMessage: Text;
    begin
        ConnectorSetup.Get();
        ConnectorSetup.TestField("Drive ID");
        ConnectorSetup.TestField("Import Folder");
        _AccessToken := ConnectorSetup.GetAccessToken();
        _AccessTokenMgmnt.InvokeAccessToken(ConnectorSetup, _MessageTxt, _AccessToken, true);
        Headers := Client.DefaultRequestHeaders();
        OneDriveURL := StrSubstNo('https://graph.microsoft.com/v1.0/drives/' + ConnectorSetup."Drive ID" + '/root:/' + ConnectorSetup."Import Folder" + '/%1:', OneDriveFile);
        Headers.Add('Authorization', StrSubstNo('Bearer %1', _AccessToken));
        if not Client.Delete(OneDriveURL, ResponseMessage) then
            if ResponseMessage.IsBlockedByEnvironment() then
                ErrorMessage := StrSubstNo(EnvironmentBlocksErr, RequestMessage.GetRequestUri())
            else
                ErrorMessage := StrSubstNo(ConnectionErr, RequestMessage.GetRequestUri());
        if ErrorMessage <> '' then
            Error(ErrorMessage);
        if ResponseMessage.IsSuccessStatusCode() then
            ErrorMessage := ''
        else
            ErrorMessage := StrSubstNo('HTTP error %1 (%2)', ResponseMessage.HttpStatusCode(), ResponseMessage.ReasonPhrase());
        if ErrorMessage <> '' then
            Error(ErrorMessage);
    end;



    local procedure GetParentReferenceJson(MoveFolderId: text) ParentRefernceJson: JsonObject
    begin
        ParentRefernceJson.Add('id', MoveFolderId)
    end;



    procedure GetTempClosingStockEntries(var TempClosingStockDetails1: Record "Closing Stock Details Buffer" temporary)
    begin
        IF TempClosingStockDetails.FINDSET then
            repeat
                TempClosingStockDetails1 := TempClosingStockDetails;
                TempClosingStockDetails1.insert;
            until TempClosingStockDetails.next = 0;
    end;

    procedure UploadItem(FileName: Text; Stream: InStream): Boolean
    var
        Client: HttpClient;
        Headers: HttpHeaders;
        RequestMessage: HttpRequestMessage;
        ResponseMessage: HttpResponseMessage;
        RequestContent: HttpContent;
        ConnectorSetup: Record "Onedrive Connector Setup";
        _AccessTokenMgmnt: codeunit "OneDrive Access Token Mgmnt.";
        _AccessToken: Text;
        ResponseText: Text;
        _MessageTxt: Text;
        SetRequestURL: Text;
    begin
        ConnectorSetup.Get();
        _AccessToken := ConnectorSetup.GetAccessToken();
        _AccessTokenMgmnt.InvokeAccessToken(ConnectorSetup, _MessageTxt, _AccessToken, true);
        //UploadIntoStream('Upload a file', '', '', FileName, Stream);
        Headers := Client.DefaultRequestHeaders();
        Headers.Add('Authorization', StrSubstNo('Bearer %1', _AccessToken));

        RequestMessage.SetRequestUri(StrSubstNo('https://graph.microsoft.com/v1.0/drives/' + ConnectorSetup."Drive ID" + '/root:/' + ConnectorSetup."Folder Name" + '/%1:/content', FileName));

        RequestMessage.Method := 'PUT';

        RequestContent.WriteFrom(Stream);
        RequestMessage.Content := RequestContent;

        if Client.Send(RequestMessage, ResponseMessage) then
            if ResponseMessage.IsSuccessStatusCode() then
                exit(true); //success

        exit(false); //fail
    end;


    procedure SetFileName(_FileName: Text)
    begin
        FileName := _FileName;
    end;

    procedure CreateAndSendEmail(CounterOk: Integer; CounterErr: Integer)
    var
        Recipients: List of [Text];
        UserSetup: Record "User Setup";
        Emailobj: Codeunit Email;
        EmailMsg: Codeunit "Email Message";
        TxtDefaultCCMailList: List of [Text];
        TxtDefaultBCCMailList: List of [Text];
        Body: Text;
        InvStockMsg: Label 'Dear Execution Team, <br><br> The Item Inventory stock details are updated in SwiftLink for date %1.<br><br> Total Number of records updated : %2 <br><br>Total Number of records had issues while updating : %3 <br><br>Thanks, <br> Digital Team';
        SubjectMsg: Label 'Item Inventory stock update status - %1';
        Subject: Text;
        _CurrentDateTime: DateTime;
        _FileName: Text;
    begin
        UserSetup.RESET;
        UserSetup.SETFILTER(UserSetup."Backend User Email ID", '<>%1', '');
        IF UserSetup.FINDFIRST THEN
            REPEAT
                Recipients.add(UserSetup."Backend User Email ID");
            UNTIL UserSetup.NEXT = 0;
        _CurrentDateTime := CurrentDateTime;
        Body := StrSubstNo(InvStockMsg, Format(Today, 0, '<Day,2>/<Month,2>/<Year4>'), CounterOk, CounterErr);
        Subject := StrSubstNo(SubjectMsg, Format(Today, 0, '<Day,2>/<Month,2>/<Year4>'));
        EmailMsg.Create(Recipients, Subject, Body, true, TxtDefaultCCMailList, TxtDefaultBCCMailList);
        Emailobj.Send(EmailMsg, Enum::"Email Scenario"::Default);
        // _FileName := 'ConsolidatedBatchReport_' + Format(Today, 0, '<Day,2><Month,2><Year4>') + '.xlsx';
    end;

    var
        ExcelBuffer: Record "Excel Buffer" temporary;
        TempClosingStockDetails: Record "Closing Stock Details Buffer" temporary;
        EnvironmentBlocksErr: Label 'Environment blocks an outgoing HTTP request to ''%1''.', Comment = '%1 - url, e.g. https://microsoft.com';
        ConnectionErr: Label 'Connection to the remote service ''%1'' could not be established.', Comment = '%1 - url, e.g. https://microsoft.com';
        FileName: Text;
        ExcelFileDataErr: Label 'Excel File %1 do not have any data rows.';
}
