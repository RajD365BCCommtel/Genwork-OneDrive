//Description: This ia a new Codeunit "OneDrive Access Token Mgmnt."
codeunit 50202 "OneDrive Access Token Mgmnt."
{
    var
        EnvironmentBlocksErr: Label 'Environment blocks an outgoing HTTP request to ''%1''.', Comment = '%1 - url, e.g. https://microsoft.com';
        ConnectionErr: Label 'Connection to the remote service ''%1'' could not be established.', Comment = '%1 - url, e.g. https://microsoft.com';
        RefreshAccessTokenTxt: Label 'Refresh access token.', Locked = true;
        InvokeRequestTxt: Label 'Invoke %1 request.', Comment = '%1 - request type, e.g. GET, POST', Locked = true;
        RefreshSuccessfulTxt: Label 'Refresh token successful.';
        RefreshFailedTxt: Label 'Refresh token failed.';
        AuthorizationSuccessfulTxt: Label 'Authorization successful.';
        ReasonTxt: Label 'Reason: ';

    procedure RefreshAndSaveAccessToken(var OnedriveConnectorSetup: Record "Onedrive Connector Setup"; var MessageText: Text) Result: Boolean
    var
        AccessToken: Text;
    begin
        Result :=
          RefreshAccessToken(
            OnedriveConnectorSetup,
            AccessToken, MessageText);

        if Result then
            SaveTokens(OnedriveConnectorSetup, AccessToken);
    end;

    procedure RefreshAccessToken(var OnedriveConnectorSetup: Record "Onedrive Connector Setup"; var AccessToken: Text; var MessageText: Text): Boolean
    begin
        exit(
            RefreshAccessTokenWithGivenRequestJson(
                OnedriveConnectorSetup, MessageText, AccessToken));
    end;

    procedure RefreshAccessTokenWithGivenRequestJson(var OnedriveConnectorSetup: Record "Onedrive Connector Setup"; var MessageText: Text; var AccessToken: Text) Result: Boolean
    begin
        exit(RefreshAccessTokenWithGivenRequestJsonAndContentType(OnedriveConnectorSetup, MessageText, AccessToken, TRUE));
    end;

    local procedure RefreshAccessTokenWithGivenRequestJsonAndContentType(var OnedriveConnectorSetup: Record "Onedrive Connector Setup"; var MessageText: Text; var AccessToken: Text; UseUrlEncodedContentType: Boolean) Result: Boolean
    var
        RequestJsonContent: JsonObject;
        RequestUrlContent: Text;
        ResponseJson: Text;
        HttpError: Text;
        ExpireInSec: BigInteger;
    begin
        with OnedriveConnectorSetup do begin
            TestField("Client ID");
            TestField("Client Secret");
            TestField("Redirect URL");
            TestField(Scope);
            TestField("Authorization URL");

            if UseUrlEncodedContentType then
                CreateContentRequestForRefreshAccessToken(RequestUrlContent, OnedriveConnectorSetup);

            Result := RequestAccessAndRefreshTokens(OnedriveConnectorSetup."Access Token URL", RequestUrlContent, ResponseJson, AccessToken, HttpError);
            SaveResultForRequestAccessAndRefreshTokens(
              OnedriveConnectorSetup, MessageText, Result, RefreshAccessTokenTxt, RefreshSuccessfulTxt,
              RefreshFailedTxt, HttpError, ResponseJson);
        end;
    end;

    procedure InvokeAccessToken(var OnedriveConnectorSetup: Record "Onedrive Connector Setup"; var MessageText: Text; var AccessToken: Text; UseUrlEncodedContentType: Boolean) Result: Boolean
    var
        RequestJsonContent: JsonObject;
        RequestUrlContent: Text;
        ResponseJson: Text;
        HttpError: Text;
        ExpireInSec: BigInteger;
    begin
        with OnedriveConnectorSetup do begin
            TestField("Client ID");
            TestField("Client Secret");
            TestField("Authorization URL");
            TestField(Scope);
            TestField("Redirect URL");

            if UseUrlEncodedContentType then
                CreateContentRequestForRefreshAccessToken(RequestUrlContent, OnedriveConnectorSetup);
            Result := RequestAccessAndRefreshTokens(OnedriveConnectorSetup."Access Token URL", RequestUrlContent, ResponseJson, AccessToken, HttpError);
        end;
    end;

    local procedure SaveResultForRequestAccessAndRefreshTokens(var OnedriveConnectorSetup: Record "Onedrive Connector Setup"; var MessageText: Text; Result: Boolean; Context: Text; SuccessMsg: Text; ErrorMsg: Text; HttpError: Text; ResponseJson: Text)
    begin
        if Result then begin
            MessageText := SuccessMsg;
            OnedriveConnectorSetup.Status := OnedriveConnectorSetup.Status::Enabled;
        end else begin
            MessageText := ErrorMsg;
            if HttpError <> '' then
                MessageText += '\' + ReasonTxt + HttpError;
            OnedriveConnectorSetup.Status := OnedriveConnectorSetup.Status::Error;
        end;
    end;

    local procedure RequestAccessAndRefreshTokens(RequestURL: Text; RequestJson: Text; var ResponseJson: Text; var AccessToken: Text; var HttpError: Text): Boolean
    begin
        ResponseJson := '';
        if InvokeHttpJSONRequest(RequestURL, RequestJson, ResponseJson, HttpError) then
            exit(ParseAccessAndRefreshTokens(ResponseJson, AccessToken));
    end;


    local procedure InvokeHttpJSONRequest(RequestURL: Text; RequestJson: Text; var ResponseJson: Text; var HttpError: Text): Boolean
    var
        Client: HttpClient;
        RequestMessage: HttpRequestMessage;
        ResponseMessage: HttpResponseMessage;
        ErrorMessage: Text;
        ReqHttpContent: HttpContent;
        ReqHttpHeaders: HttpHeaders;
    begin
        ResponseJson := '';
        HttpError := '';

        Client.Clear();
        Client.Timeout(60000);
        ReqHttpContent.Clear();
        ReqHttpContent.WriteFrom(RequestJson);
        ReqHttpHeaders.Clear();
        ReqHttpContent.GetHeaders(ReqHttpHeaders);
        ReqHttpHeaders.Remove('Content-Type');
        ReqHttpHeaders.Add('Content-Type', 'application/x-www-form-urlencoded');
        ReqHttpContent.GetHeaders(ReqHttpHeaders);
        if not Client.Post(RequestURL, ReqHttpContent, ResponseMessage) then
            if ResponseMessage.IsBlockedByEnvironment() then
                ErrorMessage := StrSubstNo(EnvironmentBlocksErr, RequestMessage.GetRequestUri())
            else
                ErrorMessage := StrSubstNo(ConnectionErr, RequestMessage.GetRequestUri());

        if ErrorMessage <> '' then
            Error(ErrorMessage);

        exit(ProcessHttpResponseMessage(ResponseMessage, ResponseJson, HttpError));
    end;

    local procedure ProcessHttpResponseMessage(var ResponseMessage: HttpResponseMessage; var ResponseJson: Text; var HttpError: Text) Result: Boolean
    var
        ResponseJObject: JsonObject;
        ContentJObject: JsonObject;
        JToken: JsonToken;
        ResponseText: Text;
        JsonResponse: Boolean;
        StatusCode: Integer;
        StatusReason: Text;
        StatusDetails: Text;
    begin
        Result := ResponseMessage.IsSuccessStatusCode();
        StatusCode := ResponseMessage.HttpStatusCode();
        StatusReason := ResponseMessage.ReasonPhrase();

        if ResponseMessage.Content().ReadAs(ResponseText) then
            JsonResponse := ContentJObject.ReadFrom(ResponseText);
        if JsonResponse then
            ResponseJObject.Add('Content', ContentJObject)
        else
            ResponseJObject.Add('ContentText', ResponseText);

        if not Result then begin
            HttpError := StrSubstNo('HTTP error %1 (%2)', StatusCode, StatusReason);
            if JsonResponse then
                if ContentJObject.SelectToken('error_description', JToken) then begin
                    StatusDetails := JToken.AsValue().AsText();
                    HttpError += StrSubstNo('\%1', StatusDetails);
                end;
        end;

        SetHttpStatus(ResponseJObject, StatusCode, StatusReason, StatusDetails);
        ResponseJObject.WriteTo(ResponseJson);
    end;

    local procedure SetHttpStatus(var JObject: JsonObject; StatusCode: Integer; StatusReason: Text; StatusDetails: Text)
    var
        JObject2: JsonObject;
    begin
        JObject2.Add('code', StatusCode);
        JObject2.Add('reason', StatusReason);
        if StatusDetails <> '' then
            JObject2.Add('details', StatusDetails);
        JObject.Add('Status', JObject2);
    end;

    local procedure ParseAccessAndRefreshTokens(ResponseJson: Text; var AccessToken: Text): Boolean
    var
        JToken: JsonToken;
        NewAccessToken: Text;
        NewRefreshToken: Text;
    begin

        if JToken.ReadFrom(ResponseJson) then
            if JToken.SelectToken('Content', JToken) then
                foreach JToken in JToken.AsObject().Values() do
                    case JToken.Path() of
                        'Content.access_token':
                            NewAccessToken := JToken.AsValue().AsText();
                    end;
        if NewAccessToken = '' THEN
            exit(false);
        IF (AccessToken = NewAccessToken) then
            exit(true);

        AccessToken := NewAccessToken;
        exit(true);
    end;

    local procedure CreateContentRequestForRefreshAccessToken(var UrlString: Text; var OnedriveConnectorSetup: Record "Onedrive Connector Setup")
    var
        TypeHelper: Codeunit "Type Helper";
        DotNetUriBuilder: Codeunit Uri;
        AuthUrl: text;
        clientid: text;
        RedirectUrl: text;
        Scope: text;
        AccessTokenUrl: Text;
        clientsecret: text;
    begin
        clientid := OnedriveConnectorSetup."Client ID";
        clientsecret := OnedriveConnectorSetup."Client Secret";
        RedirectUrl := OnedriveConnectorSetup."Redirect URL";
        AuthUrl := OnedriveConnectorSetup."Authorization URL";
        AccessTokenUrl := OnedriveConnectorSetup."Access Token URL";
        Scope := OnedriveConnectorSetup.Scope;

        UrlString := StrSubstNo('grant_type=client_credentials&client_secret=%1&client_id=%2&scope=%3',
             DotNetUriBuilder.EscapeDataString(clientsecret), DotNetUriBuilder.EscapeDataString(clientid), DotNetUriBuilder.EscapeDataString(Scope));
    end;

    local procedure SaveTokens(var OnedriveConnectorSetup: Record "Onedrive Connector Setup"; AccessToken: Text)
    begin
        with OnedriveConnectorSetup do begin
            SetAccessToken(AccessToken);
            Commit(); // need to prevent rollback to save new keys
        end;
    end;
}
