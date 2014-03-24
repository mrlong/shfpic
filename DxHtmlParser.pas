(******************************************************)
(*                得闲工作室                          *)
(*              HTML解析单元库                        *)
(*                                                    *)
(*              DxHtmlParser Unit                     *)
(*    Copyright(c) 2008-2010  不得闲                  *)
(*    email:appleak46@yahoo.com.cn     QQ:75492895    *)
(******************************************************)
unit DxHtmlParser;

interface
uses Windows,MSHTML,ActiveX,DxHtmlElement,Forms;

type
  TDxHtmlParser = class
  private
    FHtmlDoc: IHTMLDocument2;
    FHTML: string;
    FWebTables: TDxTableCollection;
    FWebElements: TDxWebElementCollection;
    FWebComb: TDxWebCombobox;
    procedure SetHTML(const Value: string);
    function GetWebCombobox(AName: string): TDxWebCombobox;
  public
    constructor Create;
    destructor Destroy;override;
    property HTML: string read FHTML write SetHTML;
    property WebTables: TDxTableCollection read FWebTables;
    property WebElements: TDxWebElementCollection read FWebElements;
    property WebCombobox[Name: string]: TDxWebCombobox read GetWebCombobox;
  end;
implementation

{ TDxHtmlParser }

constructor TDxHtmlParser.Create;
begin
  CoInitialize(nil);
  //创建IHTMLDocument2接口
  CoCreateInstance(CLASS_HTMLDocument, nil, CLSCTX_INPROC_SERVER, IID_IHTMLDocument2, FHtmlDoc);
  Assert(FHtmlDoc<>nil,'构建HTMLDocument接口失败');
  FHtmlDoc.Set_designMode('On'); //设置为设计模式，不执行脚本
  while not (FHtmlDoc.readyState = 'complete') do
  begin
    sleep(1);
    Application.ProcessMessages;
  end;                   
  FWebTables := TDxTableCollection.Create(FHtmlDoc);
  FWebElements := TDxWebElementCollection.Create(nil);
  FWebComb := TDxWebCombobox.Create(nil);
end;

destructor TDxHtmlParser.Destroy;
begin
  FWebTables.Free;
  FWebElements.Free;
  FWebComb.Free;
  CoUninitialize;
  inherited;
end;

function TDxHtmlParser.GetWebCombobox(AName: string): TDxWebCombobox;
begin
   if FWebElements.Collection <> nil then
   begin
     FWebComb.CombInterface := FWebElements.ElementByName[AName] as IHTMLSelectElement;
     Result := FWebComb;
   end
   else Result := nil;
end;

procedure TDxHtmlParser.SetHTML(const Value: string);
begin
  if FHTML <> Value then
  begin
    FHTML := Value;
    FHtmlDoc.body.innerHTML := FHTML;
    FWebElements.Collection := FHtmlDoc.all;
  end;
end;

end.
