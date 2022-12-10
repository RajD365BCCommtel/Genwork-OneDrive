report 50200 "OneDrive Sales Order Report"
{
    //ApplicationArea = All;
    Caption = 'OneDrive Sales Order Report';
    UsageCategory = ReportsAndAnalysis;
    ProcessingOnly = true;
    //UseRequestPage = false;

    dataset
    {
        dataitem(CustomerLoop; Integer)
        {
            DataItemTableView = sorting(Number) order(ascending);
            MaxIteration = 1;

            trigger OnPreDataItem()
            begin
                TempCustomer.Reset();
                TempCustomer.DeleteAll();
            end;

            trigger OnAfterGetRecord()
            var
                _SalesHeader: Record "Sales Header";
                _SalesLine: Record "Sales Line";
            begin
                _SalesHeader.Reset();
                _SalesHeader.SetCurrentKey("Document Type", "Sell-to Customer No.");
                _SalesHeader.SetRange("Document Type", _SalesHeader."Document Type"::Order);
                _SalesHeader.SetRange(Status, _SalesHeader.Status::Open);
                _SalesHeader.SetRange("Order Date", Today);
                _SalesHeader.SetRange("Sales Order Export to Excel", false);
                if _SalesHeader.FindSet() then begin
                    repeat
                        if NOT TempCustomer.Get(_SalesHeader."Sell-to Customer No.") then
                            CreateTempCustomer(_SalesHeader."Sell-to Customer No.", _SalesHeader."Sell-to Customer Name");
                    until _SalesHeader.Next() = 0;
                end;

                if TempCustomer.FindSet() then
                    repeat
                        ExcelBuffer.Reset();
                        ExcelBuffer.DeleteAll();
                        ProcessSalesOrder(TempCustomer."No.");
                        if not ExcelBuffer.IsEmpty then begin
                            ExcelBuffer.CreateNewBook('Sheet1');
                            ExcelBuffer.WriteSheet(' ', CompanyName, UserId);
                            ExcelBuffer.CloseBook();
                            CreateAndSendEmail(ExcelBuffer, TempCustomer, 'SwiftLink Order Report');
                        end;
                    until TempCustomer.Next() = 0;
            end;
        }
    }


    trigger OnPostReport()
    begin
        IF GUIALLOWED THEN
            Message('Process Completed');
    end;

    procedure CreateAndSendEmail(var TempExcelBuf: Record "Excel Buffer" temporary; var _TempCustomer: Record Customer temporary; BookName: Text)
    var
        Recipients: List of [Text];
        UserSetup: Record "User Setup";
        Emailobj: Codeunit Email;
        EmailMsg: Codeunit "Email Message";
        TxtDefaultCCMailList: List of [Text];
        TxtDefaultBCCMailList: List of [Text];
        Body: Text;
        SalesPostedMsg: Label 'Dear Executive(s), <br><br> A new Sales Order(s) is placed by  %1 - %2 the %2_DigitalOrders.xlsx file is updated on the server at %3 .<br><br> Please process the orders. <br><br> Thanks, <br> Digital Team';
        SubjectMsg: Label 'Attention - New Digital Order(s) placed ';
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
        Body := StrSubstNo(SalesPostedMsg, _TempCustomer.Name, _TempCustomer."No.", _CurrentDateTime);
        EmailMsg.Create(Recipients, SubjectMsg + '[ ' + Format(_CurrentDateTime) + '-' + _TempCustomer.Name + ' ]', Body, true, TxtDefaultCCMailList, TxtDefaultBCCMailList);
        Emailobj.Send(EmailMsg, Enum::"Email Scenario"::Default);
        _FileName := _TempCustomer."No." + '_DigitalOrders.xlsx';
        StoreFileonOneDrive(TempExcelBuf, _FileName);
    end;

    local procedure StoreFileonOneDrive(var TempExcelBuf: Record "Excel Buffer" temporary; FileName: Text)
    var
        OneDriveMgt: Codeunit "One Drive Mgt.";
        TempBlob: Codeunit "Temp Blob";
        InStr: InStream;
    begin
        Clear(OneDriveMgt);
        ExportExcelFileToBlob(TempExcelBuf, TempBlob);
        TempBlob.CreateInStream(InStr);
        OneDriveMgt.UploadItem(FileName, InStr);
    end;

    local procedure ExportExcelFileToBlob(
    var TempExcelBuf: Record "Excel Buffer" temporary;
    var TempBlob: Codeunit "Temp Blob")
    var
        OutStr: OutStream;
    begin
        TempBlob.CreateOutStream(OutStr);
        TempExcelBuf.SaveToStream(OutStr, true);
    end;

    local procedure CreateTempCustomer(_CustNo: Code[20]; _CustName: text[100])
    begin
        TempCustomer.Init();
        TempCustomer."No." := _CustNo;
        TempCustomer.Name := _CustName;
        TempCustomer.Insert(true);
    end;

    local procedure ProcessSalesOrder(CustNo: Code[20])
    var
        _SalesHeader: Record "Sales Header";
        _SalesLine: Record "Sales Line";
    begin
        _SalesHeader.Reset();
        _SalesHeader.SetCurrentKey("Document Type", "Sell-to Customer No.");
        _SalesHeader.SetRange("Document Type", _SalesHeader."Document Type"::Order);
        _SalesHeader.SetRange("Sell-to Customer No.", CustNo);
        _SalesHeader.SetRange(Status, _SalesHeader.Status::Open);
        _SalesHeader.SetRange("Order Date", Today);
        _SalesHeader.SetRange("Sales Order Export to Excel", false);
        if _SalesHeader.FindSet() then begin
            CreateExcelHeader();
            repeat
                _SalesLine.SetRange("Document Type", _SalesHeader."Document Type");
                _SalesLine.SetRange("Document No.", _SalesHeader."No.");
                if _SalesLine.FindSet() then
                    repeat
                        CreateExcelBody(_SalesHeader, _SalesLine);
                    until _SalesLine.Next() = 0;
            until _SalesHeader.Next() = 0;
        end;
    end;


    local procedure CreateExcelHeader()
    var
    begin
        ExcelBuffer.AddColumn('Order ID', false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
        ExcelBuffer.AddColumn('Date Time', false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
        ExcelBuffer.AddColumn('Product Id', false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
        ExcelBuffer.AddColumn('Products', false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
        ExcelBuffer.AddColumn('Price', false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
        ExcelBuffer.AddColumn('Qty', false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
        ExcelBuffer.AddColumn('Customer Name', false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
        ExcelBuffer.AddColumn('Bill-to Address', false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
        ExcelBuffer.AddColumn('Ship-to Address', false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
        ExcelBuffer.AddColumn('Customer Order No', false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
        ExcelBuffer.AddColumn('Customer Order Date', false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
        ExcelBuffer.AddColumn('Customer Address', false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
        ExcelBuffer.AddColumn('Shipping State Name', false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
        ExcelBuffer.AddColumn('Order By', false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
        ExcelBuffer.AddColumn('Discount Per', false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
        ExcelBuffer.AddColumn('Discount Amount', false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
        //ExcelBuffer.AddColumn('Bank Details', false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
        //ExcelBuffer.AddColumn('Claim Purpose', false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
        ExcelBuffer.AddColumn('PO Received Date', false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
        ExcelBuffer.AddColumn('Sales Person Name', false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
        ExcelBuffer.AddColumn('District Name', false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
        ExcelBuffer.AddColumn('Buy Back', false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
        ExcelBuffer.AddColumn('Delivery Terms', false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
        //ExcelBuffer.AddColumn('Interest on Delayed Payments', false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
        ExcelBuffer.AddColumn('Buyback Value', false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
        ExcelBuffer.AddColumn('Legal', false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
        ExcelBuffer.AddColumn('Pndt', false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
        ExcelBuffer.AddColumn('Dealer Name', false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
        //ExcelBuffer.AddColumn('Lot No', false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
        ExcelBuffer.AddColumn('Warranty', false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
        ExcelBuffer.AddColumn('Payment Terms', false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
        ExcelBuffer.AddColumn('Order Type Reagent', false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
        ExcelBuffer.AddColumn('Employee ID', false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);//Added by Rajesh
    end;

    local procedure CreateExcelBody(var _SalesHdr: Record "Sales Header"; _SaleLine: Record "Sales Line")
    var
    begin
        ExcelBuffer.NewRow();
        ExcelBuffer.AddColumn(_SalesHdr."SwiftLink Document No.", false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
        ExcelBuffer.AddColumn(_SalesHdr."Order Date", false, '', false, false, false, '', ExcelBuffer."Cell Type"::Date);
        ExcelBuffer.AddColumn(_SaleLine."No.", false, '', false, false, false, '', ExcelBuffer."Cell Type"::text);
        ExcelBuffer.AddColumn(_SaleLine.Description, false, '', false, false, false, '', ExcelBuffer."Cell Type"::text);
        ExcelBuffer.AddColumn(_SaleLine."Unit Price", false, '', false, false, false, '', ExcelBuffer."Cell Type"::Number);
        ExcelBuffer.AddColumn(_SaleLine.Quantity, false, '', false, false, false, '', ExcelBuffer."Cell Type"::Number);
        ExcelBuffer.AddColumn(_SalesHdr."Sell-to Customer Name", false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);

        if BilltoCust.get(_SalesHdr."Bill-to Customer No.") then begin
            if State.Get(BilltoCust."State Code") then;
            if CountryRegion.get(_SalesHdr."Bill-to Country/Region Code") then;
            if (BilltoCust."Phone No." <> '') AND (BilltoCust."Mobile Phone No." <> '') then begin
                ExcelBuffer.AddColumn(_SalesHdr."Bill-to Address" + ',' + _SalesHdr."Bill-to Address 2" + ',' + _SalesHdr."Bill-to City" + ',' + State.Description + ',' + CountryRegion.Name + ',' + 'Pincode :' + _SalesHdr."Bill-to Post Code" + ',' + 'Contact: ' + BilltoCust."Mobile Phone No." + '/' + BilltoCust."Phone No.", false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
            end else
                if (BilltoCust."Phone No." <> '') AND (BilltoCust."Mobile Phone No." = '') then begin
                    ExcelBuffer.AddColumn(_SalesHdr."Bill-to Address" + ',' + _SalesHdr."Bill-to Address 2" + ',' + _SalesHdr."Bill-to City" + ',' + State.Description + ',' + CountryRegion.Name + ',' + 'Pincode :' + _SalesHdr."Bill-to Post Code" + ',' + 'Contact: ' + BilltoCust."Phone No.", false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
                end else
                    if (BilltoCust."Phone No." = '') AND (BilltoCust."Mobile Phone No." <> '') then begin
                        ExcelBuffer.AddColumn(_SalesHdr."Bill-to Address" + ',' + _SalesHdr."Bill-to Address 2" + ',' + _SalesHdr."Bill-to City" + ',' + State.Description + ',' + CountryRegion.Name + ',' + 'Pincode :' + _SalesHdr."Bill-to Post Code" + ',' + 'Contact: ' + BilltoCust."Mobile Phone No.", false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
                    end else
                        if (BilltoCust."Phone No." = '') AND (BilltoCust."Mobile Phone No." = '') then begin
                            ExcelBuffer.AddColumn(_SalesHdr."Bill-to Address" + ',' + _SalesHdr."Bill-to Address 2" + ',' + _SalesHdr."Bill-to City" + ',' + State.Description + ',' + CountryRegion.Name + ',' + 'Pincode :' + _SalesHdr."Bill-to Post Code" + ',' + 'Contact: ' + ' ', false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
                        end;
        end;
        if (_SalesHdr."Ship-to Code" <> '') then begin
            if ShiptoAddr.get(_SalesHdr."Sell-to Customer No.", _SalesHdr."Ship-to Code") then;
            if State.Get(ShiptoAddr.State) then;
            if CountryRegion.get(ShiptoAddr."Country/Region Code") then;

            ExcelBuffer.AddColumn(ShiptoAddr.Address + ',' + ShiptoAddr."Address 2" + ',' + ShiptoAddr.City + ',' + State.Description + ',' + CountryRegion.Name + ',' + 'Pincode :' + ShiptoAddr."Post Code" + ',' + 'Contact: ' + ShiptoAddr."Phone No.", false, '', false, false, false, '', ExcelBuffer."Cell Type"::text);
        end;
        if (_SalesHdr."Ship-to Code" = '') then begin
            if (BilltoCust."Phone No." <> '') AND (BilltoCust."Mobile Phone No." <> '') then begin
                ExcelBuffer.AddColumn(_SalesHdr."Bill-to Address" + ',' + _SalesHdr."Bill-to Address 2" + ',' + _SalesHdr."Bill-to City" + ',' + State.Description + ',' + CountryRegion.Name + ',' + 'Pincode :' + _SalesHdr."Bill-to Post Code" + ',' + 'Contact: ' + BilltoCust."Mobile Phone No." + '/' + BilltoCust."Phone No.", false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
            end else
                if (BilltoCust."Phone No." <> '') AND (BilltoCust."Mobile Phone No." = '') then begin
                    ExcelBuffer.AddColumn(_SalesHdr."Bill-to Address" + ',' + _SalesHdr."Bill-to Address 2" + ',' + _SalesHdr."Bill-to City" + ',' + State.Description + ',' + CountryRegion.Name + ',' + 'Pincode :' + _SalesHdr."Bill-to Post Code" + ',' + 'Contact: ' + BilltoCust."Phone No.", false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
                end else
                    if (BilltoCust."Phone No." = '') AND (BilltoCust."Mobile Phone No." <> '') then begin
                        ExcelBuffer.AddColumn(_SalesHdr."Bill-to Address" + ',' + _SalesHdr."Bill-to Address 2" + ',' + _SalesHdr."Bill-to City" + ',' + State.Description + ',' + CountryRegion.Name + ',' + 'Pincode :' + _SalesHdr."Bill-to Post Code" + ',' + 'Contact: ' + BilltoCust."Mobile Phone No.", false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
                    end else
                        if (BilltoCust."Phone No." = '') AND (BilltoCust."Mobile Phone No." = '') then begin
                            ExcelBuffer.AddColumn(_SalesHdr."Bill-to Address" + ',' + _SalesHdr."Bill-to Address 2" + ',' + _SalesHdr."Bill-to City" + ',' + State.Description + ',' + CountryRegion.Name + ',' + 'Pincode :' + _SalesHdr."Bill-to Post Code" + ',' + 'Contact: ' + ' ', false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
                        end;
        end;

        ExcelBuffer.AddColumn(_SalesHdr."PO Reference No.", false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
        ExcelBuffer.AddColumn('', false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
        ExcelBuffer.AddColumn('', false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
        ExcelBuffer.AddColumn('', false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
        ExcelBuffer.AddColumn('', false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
        ExcelBuffer.AddColumn(_SaleLine."Line Discount %", false, '', false, false, false, '', ExcelBuffer."Cell Type"::Number);
        ExcelBuffer.AddColumn('', false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
        //ExcelBuffer.AddColumn('', false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
        //ExcelBuffer.AddColumn('', false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
        ExcelBuffer.AddColumn('', false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
        if _SalesHdr."External Document No." <> '' then
            ExcelBuffer.AddColumn(GetSalesperson(_SalesHdr."External Document No.", '-'), false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);//Added by Rajesh
        ExcelBuffer.AddColumn('', false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
        ExcelBuffer.AddColumn('', false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
        ExcelBuffer.AddColumn('', false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
        //ExcelBuffer.AddColumn('', false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
        ExcelBuffer.AddColumn('', false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
        ExcelBuffer.AddColumn('', false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
        ExcelBuffer.AddColumn('', false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
        ExcelBuffer.AddColumn('', false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
        //ExcelBuffer.AddColumn('', false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
        ExcelBuffer.AddColumn('', false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
        ExcelBuffer.AddColumn('', false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
        ExcelBuffer.AddColumn('', false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);

        //Added by Rajesh
        if _SalesHdr."External Document No." <> '' then
            ExcelBuffer.AddColumn(GetEmployeeId(_SalesHdr."External Document No.", '-'), false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
        //Added by Rajesh
        _SalesHdr."Sales Order Export to Excel" := true;
        _SalesHdr.Modify();
        Commit();
    end;

    local procedure GetSalesperson(String: Code[35]; FindWhat: Code[10]) NewString: Code[35]
    var
        FindPos: Integer;
    begin
        FindPos := STRPOS(String, FindWhat);
        if FindPos <> 0 then begin
            if FindPos > 1 then
                NewString := CopyStr(String, 1, FindPos - 1)
        end else
            NewString := String;
    end;

    local procedure GetEmployeeId(String: Code[35]; FindWhat: Code[10]) NewString: Code[35]
    var
        FindPos: Integer;
    begin
        FindPos := STRPOS(String, FindWhat);
        if FindPos <> 0 then
            NewString := CopyStr(String, (FindPos + 1));
    end;

    var
        ExcelBuffer: Record "Excel Buffer";
        TempCustomer: Record Customer temporary;
        NewCustomer: Boolean;
        EmailSentTxt: Label 'File has been sent by email';
        ShiptoCust: Record Customer;
        BilltoCust: Record Customer;
        CountryRegion: Record "Country/Region";
        State: Record State;
        ShiptoAddr: Record "Ship-to Address";
        Emailobj: Codeunit Email;
        EmailMsg: Codeunit "Email Message";
        NoOfCopy: Integer;

}
