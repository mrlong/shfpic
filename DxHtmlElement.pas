(******************************************************)
(*                得闲工作室                          *)
(*              网页元素操作类库                      *)
(*                                                    *)
(*              DxHtmlElement Unit                    *)
(*    Copyright(c) 2008-2010  不得闲                  *)
(*    email:appleak46@yahoo.com.cn     QQ:75492895    *)
(******************************************************)
unit DxHtmlElement;

interface
uses Windows,sysUtils,Clipbrd,MSHTML,ActiveX,OleCtrls,Graphics,TypInfo;

{Get EleMent Type}
function IsSelectElement(eleElement: IHTMLElement): Boolean;
function IsPwdElement(eleElement: IHTMLElement): Boolean;
function IsTextElement(element: IHTMLElement): boolean;
function IsTableElement(element: IHTMLElement): Boolean;
function IsElementCollection(element: IHTMLElement): Boolean;
function IsChkElement(element: IHTMLElement): boolean;
function IsRadioBtnElement(element: IHTMLElement): boolean;
function IsMemoElement(element: IHTMLElement): boolean;
function IsFormElement(element: IHTMLElement): boolean;
function IsIMGElement(element: IHTMLElement): boolean;
function IsInIMGElement(element: IHTMLElement): boolean;
function IsLabelElement(element: IHTMLElement): boolean;
function IsLinkElement(element: IHTMLElement): boolean;
function IsListElement(element: IHTMLElement): boolean;
function IsControlElement(element: IHTMLElement): boolean;
function IsObjectElement(element: IHTMLElement): boolean;
function IsFrameElement(element: IHTMLElement): boolean;
function IsInPutBtnElement(element: IHTMLElement): boolean;
function IsInHiddenElement(element: IHTMLElement): boolean;
function IsSubmitElement(element: IHTMLElement): boolean;
{Get ImgElement Data}
function GetPicIndex(doc: IHTMLDocument2; Src: string; Alt: string): Integer;
function GetPicElement(doc: IHTMLDocument2;imgName: string;src: string;Alt: string): IHTMLImgElement;
function GetRegCodePic(doc: IHTMLDocument2;ImgName: string; Src: string; Alt: string): TPicture; overload;
function GetRegCodePic(doc: IHTMLDocument2;Index: integer): TPicture; overload;
function GetRegCodePic(doc: IHTMLDocument2;element: IHTMLIMGElement): TPicture;overload;

type
  TObjectFromLResult = function(LRESULT: lResult;const IID: TIID; WPARAM: wParam;out pObject): HRESULT; stdcall;
  TEleMentType = (ELE_UNKNOW,ELE_TEXT,ELE_PWD,ELE_SELECT,ELE_CHECKBOX,ELE_RADIOBTN,ELE_MEMO,ELE_FORM,ELE_IMAGE,
  ELE_LABEL,ELE_LINK,ELE_LIST,ELE_CONTROL,ELE_OBJECT,ELE_FRAME,ELE_INPUTBTN,ELE_INIMAGE,ELE_INHIDDEN);


function GetElementType(element: IHTMLELEMENT): TEleMentType;
function GetElementTypeName(element: IHTMLELEMENT): string;
function GetHtmlTableCell(aTable: IHTMLTable;aRow,aCol: Integer): IHTMLElement;
function GetHtmlTable(aDoc: IHTMLDocument2; aIndex: Integer): IHTMLTable;
function GetWebBrowserHtmlTableCellText(Doc: IHTMLDocument2;
         const TableIndex, RowIndex, ColIndex: Integer;var ResValue: string):   Boolean;
function GetHtmlTableRowHtml(aTable: IHTMLTable; aRow: Integer): IHTMLElement;

function GetWebBrowserHtmlTableCellHtml(Doc: IHTMLDocument2;
         const TableIndex,RowIndex,ColIndex: Integer;var ResValue: string):   Boolean;
function GeHtmlTableHtml(aTable: IHTMLTable; aRow: Integer): IHTMLElement;
function GetWebBrowserHtmlTableHtml(Doc: IHTMLDocument2;
         const TableIndex,RowIndex: Integer;var ResValue: string):   Boolean;

type
  TDxWebFrameCollection = class;
  TDxWebElementCollection = class;

 
  TLoadState = (Doc_Loading,Doc_Completed,Doc_Invalidate);

  TDxWebFrame = class
  private
    FFrame: IHTMLWINDOW2;
    FElementCollections: TDxWebElementCollection;
    FWebFrameCollections: TDxWebFrameCollection;
    function GetSrc: string;
    function GetElementCount: integer;
    function GetWebFrameCollections: TDxWebFrameCollection;
    function GetElementCollections: TDxWebElementCollection;
    function GetDocument: IHTMLDOCUMENT2;
    function GetReadState: TLoadState;
    function GetIsLoaded: boolean;
    procedure SetFrame(const Value: IHTMLWINDOW2);
    function GetName: string;
  public
    Constructor Create(IFrame: IHTMLWINDOW2);
    Destructor Destroy;override;
    property Frame: IHTMLWINDOW2 read FFrame write SetFrame;
    property Src: string read GetSrc;
    property Document: IHTMLDOCUMENT2 read GetDocument;
    property Name: string read GetName;
    property Frames: TDxWebFrameCollection read GetWebFrameCollections;
    property ElementCount: integer read GetElementCount;
    property ElementCollections: TDxWebElementCollection read GetElementCollections;
    property ReadyState: TLoadState read GetReadState;
    property IsLoaded: boolean read GetIsLoaded;  
  end;


  TDxWebFrameCollection = Class
  private
    FFrameCollection: IHTMLFramesCollection2;
    Frame: TDxWebFrame;
    function GetCount: integer;
    function GetFrameInterfaceByIndex(index: integer): IHTMLWINDOW2;
    function GetFrameInterfaceByName(Name: string): IHTMLWINDOW2;
    function GetFrameByIndex(index: integer): TDxWebFrame;
    function GetFrameByName(Name: string): TDxWebFrame;
    procedure SetFrameCollection(const Value: IHTMLFramesCollection2);
  public
    Constructor Create(ACollection: IHTMLFramesCollection2);
    Destructor Destroy;override;
    property FrameCollection: IHTMLFramesCollection2 read FFrameCollection write SetFrameCollection;
    property Count: integer read GetCount;
    property FrameInterfaceByIndex[index: integer]: IHTMLWINDOW2 read GetFrameInterfaceByIndex;
    property FrameInterfaceByName[Name: string]: IHTMLWINDOW2 read GetFrameInterfaceByName;

    property FrameByIndex[index: integer]: TDxWebFrame read GetFrameByIndex;
    property FrameByName[Name: string]: TDxWebFrame read GetFrameByName;
  end;
  
  TDxWebElementCollection = class
  private
    FCollection: IHTMLElementCollection;
    FChildCollection:  TDxWebElementCollection;
    function GetCollection(index: String): TDxWebElementCollection;
    function GetCount: integer;
    function GetElement(itemName: string; index: integer): IHTMLElement;
    function GetElementByName(itemName: string): IHTMLELEMENT;
    function GetElementByIndex(index: integer): IHTMLELEMENT;
    procedure SetCollection(const Value: IHTMLElementCollection);
  public
    Constructor Create(ACollection: IHTMLElementCollection);
    Destructor Destroy;override;
    property Collection: IHTMLElementCollection read FCollection write SetCollection;
    property ChildElementCollection[index: String]: TDxWebElementCollection read GetCollection;
    property ElementCount: integer read GetCount;
    property Element[itemName: string;index: integer]: IHTMLElement read GetElement;
    property ElementByName[itemName: string]: IHTMLELEMENT read GetElementByName;
    property ElementByIndex[index: integer]: IHTMLELEMENT read GetElementByIndex;
  end;

  TLinkCollection = class(TDxWebElementCollection)
  
  end;
  TDxWebTable = class;

  TDxTableCollection = class
  private
    FTableCollection: IHTMLElementCollection;
    FDocument: IHTMLDOCUMENT2;
    FWebTable: TDxWebTable;
    function GetTableInterfaceByName(AName: string): IHTMLTABLE;
    procedure SetDocument(Value: IHTMLDOCUMENT2);
    function GetTableInterfaceByIndex(index: integer): IHTMLTABLE;
    function GetCount: integer;
    function GetTableByIndex(index: integer): TDxWebTable;
    function GetTableByName(AName: string): TDxWebTable;
  public
    Constructor Create(Doc: IHTMLDOCUMENT2);
    destructor Destroy;override;
    property TableInterfaceByName[AName: string]: IHTMLTABLE read GetTableInterfaceByName;
    property TableInterfaceByIndex[index: integer]: IHTMLTABLE read GetTableInterfaceByIndex;

    property TableByName[AName: string]: TDxWebTable read GetTableByName;
    property TableByIndex[index: integer]: TDxWebTable read GetTableByIndex;
    
    property Document: IHTMLDOCUMENT2 read FDocument write SetDocument;
    property Count: integer read GetCount;
  end;

  TDxWebTable = class
  private
    FTableInterface: IHTMLTABLE;
    function GetRowCount: integer;
    procedure SetTableInterface(const Value: IHTMLTABLE);
    function GetCell(ACol, ARow: integer): string;
    function GetRowColCount(RowIndex: integer): integer;
    function GetInnerHtml: string;
    function GetInnerText: string;
    function GetCellElement(ACol, ARow: Integer): IHTMLTableCell;
  public
    Constructor Create(ATable: IHTMLTABLE);
    property TableInterface: IHTMLTABLE read FTableInterface write SetTableInterface;
    property RowCount: integer read GetRowCount;
    property Cell[ACol: integer;ARow: integer]: string read GetCell;
    property CellElement[ACol: Integer;ARow: Integer]: IHTMLTableCell read GetCellElement;
    property RowColCount[RowIndex: integer]: integer read GetRowColCount;
    property InnerHtml: string read GetInnerHtml;
    property InnerText: string read GetInnerText;
  end;

  TDxWebCombobox = class
  private
    FHtmlSelect: IHTMLSelectElement;
    function GetCount: Integer;
    procedure SetItemIndex(const Value: Integer);
    function GetItemIndex: Integer;
    function GetName: string;
    procedure SetName(const Value: string);
    function GetValue: string;
    procedure SetValue(const Value: string);
    procedure SetCombInterface(const Value: IHTMLSelectElement);
    function GetItemByName(EleName: string): string;
    function GetItemByIndex(index: integer): string;
    function GetItemAttribute(index: Integer; AttribName: string): OleVariant;
  public
    constructor Create(AWebCombo: IHTMLSelectElement);
    procedure Add(Ele: IHTMLElement);
    procedure Insert(Ele: IHTMLElement;Index: Integer);
    procedure Remove(index: Integer);

    property CombInterface: IHTMLSelectElement read FHtmlSelect write SetCombInterface;
    property Count: Integer read GetCount;
    property ItemIndex: Integer read GetItemIndex write SetItemIndex;
    property ItemByIndex[index: integer]: string read GetItemByIndex;
    property ItemByName[EleName: string]: string read GetItemByName;
    property ItemAttribute[index: Integer;AttribName: string]: OleVariant read GetItemAttribute;
    property Name: string read GetName write SetName;
    property value: string read GetValue write SetValue;
  end;


 
implementation

//***********************************
//名称：  IsSelectElement
//功能： 判断传入元素是否为选择标签
//作者： 不得闲
//日期： 2008-6-9
//***********************************
function IsSelectElement(eleElement: IHTMLElement): Boolean;
var
  selElement: IHTMLSelectElement;
begin
  result := Succeeded(eleElement.QueryInterface(IHTMLSelectElement,selElement));
end;
//***********************************
//名称：  IsPwdElement
//功能： 判断传入元素是否为密码输入框标签
//作者： 不得闲
//日期： 2008-6-9
//***********************************
function IsPwdElement(eleElement: IHTMLElement): Boolean;
var
  inElement: IHTMLInPutElement;
begin
  result := Succeeded(eleElement.QueryInterface(IHTMLInPutElement,InElement));
  if result then
    result := CompareText('PASSWORD',InElement.type_) = 0;
end;
//***********************************
//名称：  IsTextElement
//功能： 判断传入元素是否为文本输入框标签
//作者： 不得闲
//日期： 2008-6-9
//***********************************
function IsTextElement(element: IHTMLElement): boolean;
var
  inElement: IHTMLInPutElement;
begin
  result := Succeeded(element.QueryInterface(IHTMLInPutElement,InElement));
  if result then
    result := CompareText('TEXT',InElement.type_) = 0;
end;

function IsTableElement(element: IHTMLElement): Boolean;
var
  inElement: IHTMLTable;
begin
  result := Succeeded(element.QueryInterface(IHTMLTable,InElement));
end;

function IsElementCollection(element: IHTMLElement): Boolean;
var
  inElement: IHTMLElementCollection;
begin
  result := Succeeded(element.QueryInterface(IHTMLElementCollection,InElement));
end;

//***********************************
//名称：  IsChkElement
//功能： 判断传入元素是否为CheckBox输入框标签
//作者： 不得闲
//日期： 2008-6-9
//***********************************
function IsChkElement(element: IHTMLElement): boolean;
var
  inElement: IHTMLInPutElement;
begin
  result := Succeeded(element.QueryInterface(IHTMLInPutElement,InElement));
  if result then
    result := CompareText('CHECKBOX',InElement.type_) = 0;
end;
//***********************************
//名称：  IsRadioBtnElement
//功能： 判断传入元素是否为RadioButton输入框标签
//作者： 不得闲
//日期： 2008-6-9
//***********************************
function IsRadioBtnElement(element: IHTMLElement): boolean;
var
  inElement: IHTMLInPutElement;
  OpBtn: IHTMLOptionButtonElement;
begin
  result := Succeeded(element.QueryInterface(IHTMLInPutElement,InElement));
  if result then
    result := CompareText('RADIO',InElement.type_) = 0;
  if not result then
    result := Succeeded(element.QueryInterface(IHTMLOptionButtonElement,opBtn));
end;
//***********************************
//名称：  IsIMGElement
//功能： 判断传入元素是否为图片标签
//作者： 不得闲
//日期： 2008-6-9
//***********************************
function IsIMGElement(element: IHTMLElement): boolean;
var
  imgElement: IHTMLImgElement;
  inElement: IHTMLInputImage;
begin
   result := Succeeded(element.QueryInterface(IHTMLImgElement,imgElement));
   if not result then
   begin
     result := Succeeded(element.QueryInterface(IHTMLInputImage,InElement));
     //if result then
      //result := CompareText('image',InElement.type_) = 0;
   end;
end;
//***********************************
//名称：  IsInIMGElement
//功能： 判断传入元素是否为输入图片（用户点击会响应）标签
//作者： 不得闲
//日期： 2008-6-9
//***********************************
function IsInIMGElement(element: IHTMLElement): boolean;
var
  inElement: IHTMLInputImage;
begin
  result := Succeeded(element.QueryInterface(IHTMLInputImage,InElement));
end;

//***********************************
//名称：  IsLabelElement
//功能： 判断传入元素是否为Label标签
//作者： 不得闲
//日期： 2008-6-9
//***********************************
function IsLabelElement(element: IHTMLElement): boolean;
var
  Lab: IHTMLLABELELEMENT;
begin
  result := Succeeded(element.QueryInterface(IHTMLLABELELEMENT,Lab));
end;
//***********************************
//名称：  IsLinkElement
//功能： 判断传入元素是否为Link连接标签
//作者： 不得闲
//日期： 2008-6-9
//***********************************
function IsLinkElement(element: IHTMLElement): boolean;
var
  Link: IHTMLLINKELEMENT;
begin
  result := Succeeded(element.QueryInterface(IHTMLLINKELEMENT,LInk));
end;
//***********************************
//名称：  IsListElement
//功能： 判断传入元素是否为List标签
//作者： 不得闲
//日期： 2008-6-9
//***********************************
function IsListElement(element: IHTMLElement): boolean;
var
  List: IHTMLLISTELEMENT;
begin
  result := Succeeded(element.QueryInterface(IHTMLLISTELEMENT,List));
end;
//***********************************
//名称：  IsControlElement
//功能： 判断传入元素是否为Control标签
//作者： 不得闲
//日期： 2008-6-9
//***********************************
function IsControlElement(element: IHTMLElement): boolean;
var
  Control: IHTMLCONTROLELEMENT;
begin
  result := Succeeded(element.QueryInterface(IHTMLCONTROLELEMENT,Control));
end;
//***********************************
//名称：  IsObjectElement
//功能： 判断传入元素是否为Object标签
//作者： 不得闲
//日期： 2008-6-9
//***********************************
function IsObjectElement(element: IHTMLElement): boolean;
var
  Obj: IHTMLOBJECTELEMENT;
begin
  result := Succeeded(element.QueryInterface(IHTMLOBJECTELEMENT,Obj));
end;
//***********************************
//名称：  IsFrameElement
//功能： 判断传入元素是否为Frame标签
//作者： 不得闲
//日期： 2008-6-9
//***********************************
function IsFrameElement(element: IHTMLElement): boolean;
var
  fram: IHTMLFRAMEELEMENT;
begin
  result := Succeeded(element.QueryInterface(IHTMLFRAMEELEMENT,Fram));
end;
//***********************************
//名称：  IsInPutBtnElement
//功能： 判断传入元素是否为输入按扭标签
//作者： 不得闲
//日期： 2008-6-9
//***********************************
function IsInPutBtnElement(element: IHTMLElement): boolean;
var
  InBtn: IHTMLINPUTBUTTONELEMENT;
begin
  result := Succeeded(element.QueryInterface(IHTMLINPUTBUTTONELEMENT,InBtn));
end;
//***********************************
//名称：  IsInHiddenElement
//功能： 判断传入元素是否为当时隐藏不可见标签
//作者： 不得闲
//日期： 2008-6-9
//***********************************
function IsInHiddenElement(element: IHTMLElement): boolean;
var
  Hidden: IHTMLINPUTHIDDENELEMENT;
begin
  result := Succeeded(element.QueryInterface(IHTMLINPUTHIDDENELEMENT,Hidden));
end;

function IsSubmitElement(element: IHTMLElement): boolean;
var
  SubEle: IHTMLINPUTELEMENT;
begin
  result := Succeeded(element.QueryInterface(IHTMLINPUTELEMENT,SubEle));
  if result then
     result := CompareText('submit',SubEle.type_) = 0;
end;

//***********************************
//名称：  IsMemoElement
//功能： 判断传入元素是否为备注输入框标签
//作者： 不得闲
//日期： 2008-6-9
//***********************************
function IsMemoElement(element: IHTMLElement): boolean;
var
  inElement: IHTMLTEXTAreaElement;
begin
  result := Succeeded(element.QueryInterface(IHTMLTEXTAreaElement,InElement));
end;

//***********************************
//名称：  IsFormElement
//功能： 判断传入元素是否为表单
//作者： 不得闲
//日期： 2008-6-9
//***********************************
function IsFormElement(element: IHTMLElement): boolean;
var
  FormElement: IHTMLFormElement;
begin
  result := Succeeded(element.QueryInterface(IHTMLFormElement,FormElement));
end;

//***********************************
//名称：  GetElementType
//功能： 得到传递进来的元素的类型
//作者： 不得闲
//日期： 2008-6-9
//***********************************
function GetElementType(element: IHTMLELEMENT): TEleMentType;
begin
  if IsSelectEleMent(element) then
    result := ELE_SELECT
  else if IsTextEleMent(element) then
    result := ELE_TEXT
  else if IsPwdEleMent(element) then
    result := ELE_PWD
  else if IsFormElement(element) then
    result := ELE_FORM
  else if IsMemoElement(element) then
    result := ELE_MEMO
  else if IsRadioBtnElement(element) then
    result := ELE_RADIOBTN
  else if IsChkElement(element) then
    result := ELE_CHECKBOX
  else if IsInImgElement(element) then
    result := ELE_INIMAGE
  else if IsIMGElement(element) then
    result := ELE_IMAGE
  else if IsLabelElement(element) then
    result := ELE_LABEL      
  else if IsLinkElement(element) then
    result := ELE_LINK
  else if IsListElement(element) then
    result := ELE_LIST
  else if IsObjectElement(element) then
    result := ELE_OBJECT
  else if IsFrameElement(element) then
    result := ELE_FRAME
  else if IsInPutBtnElement(element) then
    result := ELE_INPUTBTN
  else if IsInHiddenElement(element) then
    result := ELE_INHIDDEN
  else if IsControlElement(element) then
    result := ELE_CONTROL
  else result := ELE_UNKNOW;
end;

//***********************************
//名称：  GetPicIndex
//功能： 根据Alt,Src得到图片标签的索引
//作者： 不得闲
//日期： 2008-6-9
//***********************************
function GetPicIndex(doc: IHTMLDocument2; Src: string; Alt: string): Integer;
var
  I: Integer;
  img: IHTMLImgElement;
begin
  Result := -1;
  for I := 0 to doc.images.length - 1 do
  begin
    img := doc.images.item(I, EmptyParam) as IHTMLImgElement;
    if Alt = '' then
    begin
      if SameText(img.src, Src) then
      begin
        Result := I;
        Break;
      end;
    end else
    begin
      if SameText(img.alt, Alt) then
      begin
        Result := I;
        Break;
      end;
    end;
  end;
end;

function GetPicElement(doc: IHTMLDocument2;imgName: string;src: string;Alt: string): IHTMLImgElement;
var
  I: Integer;
  img: IHTMLImgElement;
begin
  Result := nil;
  if ImgName <> '' then //如果没有图片的名字,通过Src或Alt中的关键字来取
  begin
     Img := doc.images.item(ImgName, EmptyParam) as IHTMLImgElement;
     result := Img;
  end
  else
  for I := 0 to doc.images.length - 1 do
  begin
    img := doc.images.item(I, EmptyParam) as IHTMLImgElement;
    if Alt = '' then
    begin
      if SameText(img.src, Src) then
      begin
        Result := img;
        Break;
      end;
    end else
    begin
      if SameText(img.alt, Alt) then
      begin
        Result := img;
        Break;
      end;
    end;
  end;
end;

//***********************************
//名称：  GetRegCodePic
//功能： 得到图片标签的图片数据
//作者： 不得闲
//日期： 2008-6-9
//***********************************
function GetRegCodePic(doc: IHTMLDocument2;element: IHTMLIMGElement): TPicture;
var
  rang: IHTMLControlRange;
  temp: IHTMLControlElement;
begin
  result := nil;
  OleInitialize(nil);
  rang := (doc.body as IHTMLElement2).createControlRange as IHTMLControlRange;
  if Succeeded(element.QueryInterface(IHTMLControlElement,temp)) then
  begin
    rang.add(temp);
    ClipBoard.Clear;
    result := TPicture.Create;
    rang.execCommand('Copy', False, EmptyParam);
    result.Assign(clipboard);
  end;
  rang := nil;
  OleUninitialize;
end;

function GetRegCodePic(doc: IHTMLDocument2;Index: integer): TPicture; overload;
var
  rang: IHTMLControlRange;
  temp: IHTMLControlElement;
  i: integer;
begin
  OleInitialize(nil);
  rang := (doc.body as IHTMLElement2).createControlRange as IHTMLControlRange;
  temp := doc.all.item(index,emptyParam) as IHTMLControlElement;
  rang.add(temp);
  ClipBoard.Clear;
  result := TPicture.Create;
  rang.execCommand('Copy', False, EmptyParam);
  result.Assign(clipboard);
 for i := 0 to rang.length - 1 do
    rang.remove(i);
  temp._Release;
  rang._Release;
  OleUninitialize;
end;
//***********************************
//名称：  GetRegCodePic
//功能： 得到图片标签的图片数据
//作者： 不得闲
//日期： 2008-6-9
//***********************************
function GetRegCodePic(doc: IHTMLDocument2;
  ImgName: string; Src: string; Alt: string): TPicture;
var
  body: IHTMLElement2;
  rang: IHTMLControlRange;
  Img: IHTMLControlElement;
  ImgNum: Integer;
begin
  OleInitialize(nil);
  Result := nil;
  body := doc.body as IHTMLElement2;
  rang := body.createControlRange() as IHTMLControlRange;
  if ImgName = '' then //如果没有图片的名字,通过Src或Alt中的关键字来取
  begin
    ImgNum := GetPicIndex(Doc, Src,Alt);
    if ImgNum < 0 then Exit;
    Img := doc.images.item(ImgNum, EmptyParam) as IHTMLControlElement;
  end else Img := doc.images.item(ImgName, EmptyParam) as IHTMLControlElement;
  rang.add(Img);
  ClipBoard.Clear;
  result := TPicture.Create;
  rang.execCommand('Copy', False, EmptyParam);
  result.Assign(clipboard);
  OleUninitialize;
end;

//***********************************
//名称：  GetElementTypeName
//功能： 获得当前的元素的标签类型
//作者： 不得闲
//日期： 2008-6-9
//***********************************
function GetElementTypeName(element: IHTMLELEMENT): string;
var  
  ti: PTypeInfo;
  //td: PTypeData;
begin
  ti := TypeInfo(TEleMentType);
  //td := GetTypeData(ti);
  result := GetEnumName(ti,integer(GetElementType(element)));
end;


function GetHtmlTableCell(aTable: IHTMLTable;aRow,aCol: Integer): IHTMLElement;
var
  Row: IHTMLTableRow;
begin
  Result := nil;
  if (aTable <> nil) and (aTable.rows <> nil) then
    Row := aTable.rows.item(aRow, aRow) as IHTMLTableRow;
  if Row <> nil   then
    Result := Row.cells.item(aCol, aCol) as IHTMLElement;
end;

function GetHtmlTable(aDoc: IHTMLDocument2; aIndex: Integer): IHTMLTable;
var
  list: IHTMLElementCollection;
begin
  Result := nil;
  if (aDoc <> nil) and (aDoc.all <> nil) then
    list := aDoc.all.tags('table') as IHTMLElementCollection;
  if list <> nil then
    Result := list.item(aIndex,aIndex) as IHTMLTable;
end;

function GetWebBrowserHtmlTableCellText(Doc: IHTMLDocument2;
         const TableIndex, RowIndex, ColIndex: Integer;var ResValue: string):   Boolean;
var
  tblintf: IHTMLTable;
  node: IHTMLElement;
begin
    ResValue   :=   '';
    tblintf := GetHtmlTable(Doc, TableIndex);
    node := GetHtmlTableCell(tblintf, RowIndex, ColIndex);
    Result := node <> nil;
    if Result then
        ResValue := Trim(node.innerText);
end;

function GetHtmlTableRowHtml(aTable: IHTMLTable; aRow: Integer): IHTMLElement;
var
  Row: IHTMLTableRow;
begin
  Result := nil;
  if (aTable <> nil) and (aTable.rows <> nil) then
    Row := aTable.rows.item(aRow, aRow) as IHTMLTableRow;
  if Row <> nil then
  Result := Row as IHTMLElement;
end;

function GetWebBrowserHtmlTableCellHtml(Doc: IHTMLDocument2;
         const TableIndex,RowIndex,ColIndex: Integer;var ResValue: string):   Boolean;
var
  tblintf: IHTMLTable;
  node: IHTMLElement;
begin
    ResValue := '';
    tblintf := GetHtmlTable(Doc,   TableIndex);
    node := GetHtmlTableCell(tblintf,   RowIndex,   ColIndex);
    Result := node   <>   nil;
    if Result then
      ResValue := Trim(node.innerHTML);
end;

function GeHtmlTableHtml(aTable: IHTMLTable; aRow: Integer): IHTMLElement;
var
  Row: IHTMLTableRow;
begin
    Result := nil;
    if (aTable <> nil) and (aTable.rows <> nil) then
      Row := aTable.rows.item(aRow, aRow) as IHTMLTableRow;
    if Row <> nil then
      Result := Row as IHTMLElement;
end;

function GetWebBrowserHtmlTableHtml(Doc: IHTMLDocument2;
         const TableIndex,RowIndex: Integer;var ResValue: string):   Boolean;
var
  tblintf: IHTMLTable;
  node: IHTMLElement;
begin
    ResValue := '';
    tblintf := GetHtmlTable(Doc, TableIndex);
    node := GeHtmlTableHtml(tblintf, RowIndex);
    Result := node <> nil;
    if Result then
      ResValue := node.innerHtml;
end;

{ TDxWebFrame }

constructor TDxWebFrame.Create(IFrame: IHTMLWINDOW2);
begin
  FFrame := IFrame;
end;

destructor TDxWebFrame.Destroy;
begin
  if FElementCollections <> nil then
    FreeAndNil(FElementCollections);
  if FWebFrameCollections <> nil then
     FreeandNil(FwebFrameCollections);
  inherited;
end;

function TDxWebFrame.GetDocument: IHTMLDOCUMENT2;
begin
  if FFrame <> nil then
    result := FFrame.document
  else Result := nil;
end;

function TDxWebFrame.GetElementCollections: TDxWebElementCollection;
begin
  if FFrame = nil then
  begin
    Result := nil;
    exit;
  end;
  if FElementCollections = nil then
     FElementCollections := TDxWebElementCollection.Create(FFrame.document.all)
  else FElementCollections.FCollection := FFrame.document.all;
  result := FElementCollections;
end;

function TDxWebFrame.GetElementCount: integer;
begin
  if Document <> nil then
    result := Document.all.length
  else Result := 0;
end;

function TDxWebFrame.GetIsLoaded: boolean;
begin
  if FFrame = nil then
    Result := True
  else result := CompareText(Frame.document.readyState,'complete') = 0;
end;

function TDxWebFrame.GetName: string;
begin
  if FFrame <> nil then
    Result := FFrame.name
  else Result := '';
end;

function TDxWebFrame.GetReadState: TLoadState;
begin
  if FFrame = nil then
  begin
    Result := Doc_Invalidate;
    exit;
  end;
  if CompareText(Frame.document.readyState,'loading') = 0 then
    result := Doc_Loading
  else if CompareText(Frame.document.readyState,'complete') = 0 then
     result := Doc_Completed
  else result := Doc_Invalidate;
end;

function TDxWebFrame.GetSrc: string;
begin
  if FFrame <> nil then
    result := FFrame.location.href
  else Result := '';
end;

function TDxWebFrame.GetWebFrameCollections: TDxWebFrameCollection;
begin
   if FFrame = nil then
   begin
     Result := nil;
     exit;
   end;
   if FWebFrameCollections = nil then
     FWebFrameCollections := TDxWebFrameCollection.Create(FFrame.document.frames)
   else  FWebFrameCollections.FFrameCollection := Document.frames;
   result := FWebFrameCollections; 
end;

procedure TDxWebFrame.SetFrame(const Value: IHTMLWINDOW2);
begin
  if FFrame <> Value then
    FFrame := Value;
end;

{ TDxWebFrameCollection }

constructor TDxWebFrameCollection.Create(
  ACollection: IHTMLFramesCollection2);
begin
  FFrameCollection := ACollection;
  Frame := TDxWebFrame.Create(nil);
end;

destructor TDxWebFrameCollection.Destroy;
begin
  Frame.Free;
  inherited;
end;

function TDxWebFrameCollection.GetCount: integer;
begin
  if FFrameCollection <> nil then
    result := FFrameCollection.length
  else Result := 0;
end;

function TDxWebFrameCollection.GetFrameByIndex(index: integer): TDxWebFrame;
begin  
  Frame.FFrame := FrameInterfaceByIndex[index];
  if Frame.FFrame <> nil then
    Result := Frame
  else Result := nil;
end;

function TDxWebFrameCollection.GetFrameByName(Name: string): TDxWebFrame;
begin
  Frame.FFrame := FrameInterfaceByName[Name];
  if Frame.FFrame <> nil then
    Result := Frame
  else Result := nil;
end;

function TDxWebFrameCollection.GetFrameInterfaceByIndex(index: integer): IHTMLWINDOW2;
var
  j: olevariant;
  temp: IDispatch;
begin
  if FFrameCollection = nil then
    Result := nil
  else
  begin
    j := index;
    temp :=  FFrameCollection.item(j);
    result := temp as IHTMLWINDOW2;
  end;
end;

function TDxWebFrameCollection.GetFrameInterfaceByName(Name: string): IHTMLWINDOW2;
var
  j: olevariant;
  temp: IDispatch;
begin
  if FFrameCollection = nil then
    Result := nil
  else
  begin
    j := Name;
    temp :=  FFrameCollection.item(j);
    temp.QueryInterface(IHTMLWINDOW2,result);
  end;
end;

procedure TDxWebFrameCollection.SetFrameCollection(
  const Value: IHTMLFramesCollection2);
begin
  FFrameCollection := Value;
end;

{ TDxWebElementCollection }

constructor TDxWebElementCollection.Create(
  ACollection: IHTMLElementCollection);
begin
  inherited Create;
  FCollection := ACollection;
end;

destructor TDxWebElementCollection.Destroy;
begin
  if FChildCollection <> nil then
    FreeandNil(FChildCollection);
  inherited;
end;

function TDxWebElementCollection.GetCollection(index: String): TDxWebElementCollection;
var
  temp: IHTMLElementCollection;
begin
   if FCollection = nil then
   begin
     Result := nil;
     exit;
   end;
   temp := FCollection.tags(index) as IHTMLElementCollection;
   if FChildCollection =  nil then
     FChildCollection := TDxWebElementCollection.Create(temp);
   FChildCollection.FCollection := temp;
   result := FChildCollection;
end;

function TDxWebElementCollection.GetCount: integer;
begin
   if FCollection = nil then
   begin
     Result := 0;
     exit;
   end;
   result := FCollection.length;
end;

function TDxWebElementCollection.GetElement(itemName: string;
  index: integer): IHTMLElement;
begin
  if FCollection = nil then
    Result := nil
  else result := FCollection.item(itemName,index) as IHTMLElement;
end;

function TDxWebElementCollection.GetElementByIndex(
  index: integer): IHTMLELEMENT;
begin
  if FCollection = nil then
    Result := nil
  else result := FCollection.item(index,0) as IHTMLElement;
end;

function TDxWebElementCollection.GetElementByName(
  itemName: string): IHTMLELEMENT;
begin
  if FCollection = nil then
    Result := nil
  else
    result := FCollection.item(itemName,0) as IHTMLElement;
end;

procedure TDxWebElementCollection.SetCollection(
  const Value: IHTMLElementCollection);
begin
  FCollection := Value;
end;

{ TDxTableCollection }

constructor TDxTableCollection.Create(Doc: IHTMLDOCUMENT2);
begin
  Document := Doc;
  FWebTable := TDxWebTable.Create(nil);
end;

destructor TDxTableCollection.Destroy;
begin
  FWebTable.TableInterface := nil;
  FWebTable.Free;
  inherited;
end;

function TDxTableCollection.GetCount: integer;
begin
   result := FTableCollection.length;
end;

function TDxTableCollection.GetTableByIndex(index: integer): TDxWebTable;
begin
  FWebTable.TableInterface := TableInterfaceByIndex[index];
  if FWebTable.TableInterface <> nil then
    Result := FWebTable
  else Result := nil;
end;

function TDxTableCollection.GetTableByName(AName: string): TDxWebTable;
begin
  FWebTable.TableInterface := TableInterfaceByName[AName];
  if FWebTable.TableInterface <> nil then
    Result := FWebTable
  else Result := nil;
end;

function TDxTableCollection.GetTableInterfaceByIndex(index: integer): IHTMLTABLE;
begin
  result := FTableCollection.item(index,0) as IHTMLTABLE;
end;

function TDxTableCollection.GetTableInterfaceByName(AName: string): IHTMLTABLE;
begin
  result := FTableCollection.item(AName,0) as IHTMLTABLE;
end;

procedure TDxTableCollection.SetDocument(Value: IHTMLDOCUMENT2);
begin
  if Value <> FDocument then
  begin
    FDocument := Value;       
    if FDocument.all <> nil then
      FTableCollection := FDocument.all.tags('table') as  IHTMLElementCollection;
  end;
end;

{ TDxWebTable }

constructor TDxWebTable.Create(ATable: IHTMLTABLE);
begin
  //if ATable = nil then
  //  Raise Exception.Create('错误！传递接口无效！');
  FTableInterface := ATable;
end;

function TDxWebTable.GetCell(ACol, ARow: integer): string;
var
  TableRow: IHTMLTableRow;
begin
   result := '';
   if (FTableInterface <> nil ) and (FTableInterface.rows <> nil) then
     TableRow := FTableInterface.rows.item(ARow,0) as IHTMLTableRow;
   if TableRow <> nil then
   begin
     result := trim((TableRow.cells.item(ACol,0) as IHTMLELEMENT).innerText);
   end;
end;


function TDxWebTable.GetCellElement(ACol, ARow: Integer): IHTMLTableCell;
var
  TableRow: IHTMLTableRow;
begin
   result := nil;   
   if (FTableInterface <> nil ) and (FTableInterface.rows <> nil) then
     TableRow := FTableInterface.rows.item(ARow,0) as IHTMLTableRow;
   if TableRow <> nil then
   begin
     result := TableRow.cells.item(ACol,0) as IHTMLTableCell;
   end;
end;

function TDxWebTable.GetInnerHtml: string;
begin
  if TableInterface <> nil then
    Result := (TableInterface as IHTMLEleMent).innerHTML
  else Result := '';
end;

function TDxWebTable.GetInnerText: string;
begin
  if TableInterface <> nil then
    Result := (TableInterface as IHTMLEleMent).innerText
  else Result := '';
end;

function TDxWebTable.GetRowColCount(RowIndex: integer): integer;
var
  TableRow: IHTMLTableRow;
begin
   result := 0;   
   if (FTableInterface <> nil ) and (FTableInterface.rows <> nil) then
     TableRow := FTableInterface.rows.item(RowIndex,0) as IHTMLTableRow;
   if TableRow <> nil then
     result := TableRow.cells.length;     
end;

function TDxWebTable.GetRowCount: integer;
begin
  result := FTableInterface.rows.length;
end;

procedure TDxWebTable.SetTableInterface(const Value: IHTMLTABLE);
begin
  if FTableInterface <> value then
    FTableInterface := Value;
end;


{ TDxWebCombobox }

procedure TDxWebCombobox.Add(ele: IHTMLElement);
begin
  if FHtmlSelect <> nil then
  begin
    FHtmlSelect.add(Ele,Count);
  end;
end;

constructor TDxWebCombobox.Create(AWebCombo: IHTMLSelectElement);
begin
  FHtmlSelect := AWebCombo;
end;

function TDxWebCombobox.GetCount: Integer;
begin
  if FHtmlSelect <> nil then
    Result := FHtmlSelect.length
  else Result := 0;
end;

function TDxWebCombobox.GetItemAttribute(index: Integer;
  AttribName: string): OleVariant;
begin
  if FHtmlSelect <> nil then
    Result := (FHtmlSelect.item(index,0) as IHTMLElement).getAttribute(AttribName,0)
  else Result := varEmpty;
end;

function TDxWebCombobox.GetItemByIndex(index: integer): string;
begin
  if FHtmlSelect <> nil then
    Result := (FHtmlSelect.item(index,0) as IHTMLElement).innerText
  else Result := '';
end;

function TDxWebCombobox.GetItemByName(EleName: string): string;
begin
  if FHtmlSelect <> nil then
    Result := (FHtmlSelect.item(EleName,0) as IHTMLElement).innerText
  else Result := '';
end;

function TDxWebCombobox.GetItemIndex: Integer;
begin
  if FHtmlSelect <> nil then
    Result := FHtmlSelect.selectedIndex
  else Result := -1;
end;

function TDxWebCombobox.GetName: string;
begin
  if FHtmlSelect <> nil then
    Result := FHtmlSelect.name
  else Result := '';
end;

function TDxWebCombobox.GetValue: string;
begin
  if FHtmlSelect <> nil then
    Result := FHtmlSelect.value
  else Result := '';
end;

procedure TDxWebCombobox.Insert(Ele: IHTMLElement; Index: Integer);
begin
  if FHtmlSelect <> nil then
    FHtmlSelect.add(Ele,index);
end;

procedure TDxWebCombobox.Remove(index: Integer);
begin
  if FHtmlSelect <> nil then
    FHtmlSelect.remove(index);
end;

procedure TDxWebCombobox.SetCombInterface(const Value: IHTMLSelectElement);
begin
  FHtmlSelect := Value;
end;

procedure TDxWebCombobox.SetItemIndex(const Value: Integer);
begin
  if FHtmlSelect <> nil then
    FHtmlSelect.selectedIndex := Value;
end;

procedure TDxWebCombobox.SetName(const Value: string);
begin
  if FHtmlSelect <> nil then
    FHtmlSelect.name := Value;  
end;

procedure TDxWebCombobox.SetValue(const Value: string);
begin
  if FHtmlSelect <> nil then
    FHtmlSelect.value := Value;
end;

end.
