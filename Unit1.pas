unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB, Grids, DBGrids, StdCtrls, IdBaseComponent,
  IdComponent, IdTCPConnection, IdTCPClient, IdHTTP;

type
  TForm1 = class(TForm)
    ds1: TDataSource;
    con1: TADOConnection;
    tbl1: TADOTable;
    dbgrd1: TDBGrid;
    btn1: TButton;
    idhtp1: TIdHTTP;
    chk1: TCheckBox;
    mmo1: TMemo;
    mmo2: TMemo;
    lbl1: TLabel;
    procedure btn1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation
uses
  shellapi,
  DxHtmlElement,
  DxHtmlParser,
  PerlRegEx,
  IdURI;

{$R *.dfm}

function NewGuid():string;
var
  LTep: TGUID;
begin
  CreateGUID(LTep);
  Result := GUIDToString(LTep);
end;

function getimge(url,filename:string):Boolean;
var
  myidhttp : TIdHTTP;
  myms : TMemoryStream;
begin
  Result := False;
  myidhttp := TIdHTTP.Create(nil);
  myms := TMemoryStream.Create;
  try
    try
      myidhttp.Get(TIdURI.PathEncode(UTF8Encode(url)),myms);
    except
      Exit;
    end;
    
    if (myidhttp.ResponseCode = 200) and (myms.Size>900) then
    begin
      myms.Position := 0;
      myms.SaveToFile(filename);
      Result := True;
    end;
  finally
    myidhttp.Free;
    myms.Free;
  end;
end;


Function DelTree(DirName : string): Boolean;
var
SHFileOpStruct : TSHFileOpStruct;
DirBuf        : array [0..255] of char;
begin
  try
    Fillchar(SHFileOpStruct,Sizeof(SHFileOpStruct),0);
    FillChar(DirBuf, Sizeof(DirBuf), 0 );
    StrPCopy(DirBuf, DirName);
    with SHFileOpStruct do begin
    Wnd    := 0;
    pFrom := @DirBuf;
    wFunc := FO_DELETE;
    fFlags := FOF_ALLOWUNDO;
    fFlags := fFlags or FOF_NOCONFIRMATION;
    fFlags := fFlags or FOF_SILENT;
    end;
    Result := (SHFileOperation(SHFileOpStruct) = 0);
  except
    Result := False;
  end;
end;


function parsehtml(AFileName:string):Boolean;
var
  mydir: string;
  myparser : TDxHtmlParser;
  mysl : TStringList;
  i,r,c : Integer;
  myTable: TDxWebTable;
  mystr,mystr2 : string;
  reg,reg2 : TPerlRegEx; //声明正则表达式变量
  myNameList : TStringList;
  myindex : Integer;
  mypicdir,mypicname : string;
  mypicurl : string;
  myfilenlist : TStringList;

  //取出中文
  function IsIncludeChinese( s:string): string;
  var
    i : Integer;
    mystr : string;
  begin
    mystr := '';
    for i := 1 to  length(s) do
    begin
      if not (ByteType(s,i) = mbSingleByte) then
      begin
        mystr := mystr + s[i];
      end
    end;
    result := mystr;
  end;

begin
  Result := False;
  if not FileExists(AFileName) then Exit;

  mydir := ExtractFileDir(AFileName);
  myparser := TDxHtmlParser.Create;
  mysl := TStringList.Create;
  myNameList := TStringList.Create;
  myfilenlist := TStringList.Create;
  try

    mysl.LoadFromFile(AFileName);
    myparser.HTML := mysl.Text;

    for i:=0 to myparser.WebTables.Count -1 do
    begin
      myTable := myparser.WebTables.TableByIndex[i];
      mystr := LowerCase(myTable.InnerHtml);


      reg2 := TPerlRegEx.Create;
      reg2.Subject := mystr;
      reg2.RegEx := '<td.style.+?>?</td>';
      {
      【匹配结果：12】
      (1)<td style="font-size: 14px; color: red" align=center border="0">褚遂良<br></td>
      (2)<td style="font-size: 14px; color: red" align=center border="0">褚遂良<br></td>
      (3)<td style="font-size: 14px; color: red" align=center border="0">高贞碑<br></td>
      (4)<td style="font-size: 14px; color: red" align=center border="0">柳公权<br></td>
      (5)<td style="font-size: 14px; color: red" align=center border="0">欧阳询<br></td>
      (6)<td style="font-size: 14px; color: red" align=center border="0">苏孝慈墓志<br></td>
      (7)<td style="font-size: 14px; color: red" align=center border="0">吴宽<br></td>
      (8)<td style="font-size: 14px; color: red" align=center border="0">颜真卿<br></td>
      (9)<td style="font-size: 14px; color: red" align=center border="0">颜真卿<br></td>
      (10)<td style="font-size: 14px; color: red" align=center border="0">智永<br></td>
      (11)<td style="font-size: 14px; color: red" align=center border="0">智永<br></td>
      (12)<td style="font-size: 14px; color: red" align=center border="0"><br></td>
      }
      myNameList.Clear;
      while reg2.MatchAgain do
      begin
        mystr2 := reg2.MatchedText;
        mystr2 := IsIncludeChinese(mystr2);
        if mystr2 <> '' then
          myNameList.Add(mystr2);
      end;

      reg2.Free;

      reg:= TPerlRegEx.Create;
      reg.Subject := mystr;
      reg.RegEx := '<img.src.+?>';

      {
      【匹配结果：22】
      (1)<a href="about:show.php?wordtype=k&amp;filename=108690.png">
      (2)<img src="about:thumb/k/108690.png" width=auto height=auto>
      (3)<a href="about:show.php?wordtype=k&amp;filename=108691.png">
      (4)<img src="about:thumb/k/108691.png" width=auto height=auto>
      (5)<a href="about:show.php?wordtype=k&amp;filename=108692.png">
      (6)<img src="about:thumb/k/108692.png" width=auto height=auto>
      (7)<a href="about:show.php?wordtype=k&amp;filename=108693.png">
      (8)<img src="about:thumb/k/108693.png" width=auto height=auto>
      (9)<a href="about:show.php?wordtype=k&amp;filename=108694.png">
      (10)<img src="about:thumb/k/108694.png" width=auto height=auto>
      (11)<a href="about:show.php?wordtype=k&amp;filename=108695.png">
      (12)<img src="about:thumb/k/108695.png" width=auto height=auto>
      (13)<a href="about:show.php?wordtype=k&amp;filename=108696.png">
      (14)<img src="about:thumb/k/108696.png" width=auto height=auto>
      (15)<a href="about:show.php?wordtype=k&amp;filename=108697.png">
      (16)<img src="about:thumb/k/108697.png" width=auto height=auto>
      (17)<a href="about:show.php?wordtype=k&amp;filename=108698.png">
      (18)<img src="about:thumb/k/108698.png" width=auto height=auto>
      (19)<a href="about:show.php?wordtype=k&amp;filename=108699.png">
      (20)<img src="about:thumb/k/108699.png" width=auto height=auto>
      (21)<a href="about:show.php?wordtype=k&amp;filename=108700.png">
      (22)<img src="about:thumb/k/108700.png" width=auto height=auto>
      }

      {
      楷书: k
      草书: c
      行书: x
      隶书：l
      篆书: z
      }

      c := 0;
      while reg.MatchAgain do
      begin

        if c > myNameList.Count-1 then Break;
        
        mystr2 := reg.MatchedText;
        mystr2 := StringReplace(mystr2,'amp;','',[rfReplaceAll]);
        if Pos('thumb/k/',mystr2) > 0 then
        begin
          mypicdir := 'k';

        end
        else if Pos('thumb/x/',mystr2) > 0 then
        begin
          mypicdir :=  'x';
        end
        else if Pos('thumb/c/',mystr2) > 0 then
        begin
          mypicdir := 'c';
        end
        else if Pos('thumb/l/',mystr2) > 0 then
        begin
          mypicdir := 'l';
        end
        else if Pos('thumb/z/',mystr2) > 0 then
        begin
          mypicdir := 'z';
        end
        else
          Exit;

        myindex := Pos('about:thumb/' + mypicdir + '/',mystr2);
        mystr2 := Copy(mystr2,myindex+14,Pos('.png',mystr2)-myindex-14)+'.png';



        mypicurl := Format('http://dianshibaidu.cn/shufa/large/%s/%s',[mypicdir,mystr2]);
        mypicname := Format('%s\%s_%s.png',[mydir,NewGuid(),myNameList.Strings[c]]);
        if getimge(mypicurl,mypicname) then
          myfilenlist.Add(Format('%sb%s=%s',[mypicdir,myNameList.Strings[c],ExtractFileName(mypicname)]));

        mypicurl := Format('http://dianshibaidu.cn/shufa/thumb/%s/%s',[mypicdir,mystr2]); 
        mypicname := Format('%s\%s__%s.png',[mydir,NewGuid(),myNameList.Strings[c]]);
        if  getimge(mypicurl,mypicname) then
          myfilenlist.Add(Format('%ss%s=%s',[mypicdir,myNameList.Strings[c],ExtractFileName(mypicname)]));


        c := c + 1;
      end;

      reg.Free;
    end;

    Result := myfilenlist.Count>0;
    myfilenlist.SaveToFile(ChangeFileExt(AFileName,'.txt'));
  finally
    myparser.Free;
    mysl.Free;
    myNameList.Free;
    myfilenlist.Free;
  end;
end;

{
procedure TForm1.Button1Click(Sender: TObject);
var
  i: Integer;
  Table: TDxWebTable;
  strText,LastStr: string;
  Item: TListItem;
begin
  HtmlParser.HTML := (WebBrowser1.Document as IHTMLDocument2).body.innerHTML;
  ListView1.Clear;
  for i := 0 to HtmlParser.WebTables.Count - 1 do
  begin
    Table := HtmlParser.WebTables.TableByIndex[i];
    strText := Table.innerText;
    if (Table.RowCount > 2) and  (CompareStr(strText,LastStr)<>0) and (Pos('地址：',strText) <> 0) and (Pos('邮编:',strText)<>0) and (Pos('移动黄页:',strText)=0) then
    begin
      Item := ListView1.Items.Add;
      Item.Caption := Table.Cell[0,1];
      Item.SubItems.Add(Table.Cell[0,2]);
      Item.SubItems.Add(Table.Cell[0,3]);          
      LastStr := strText;
    end;
  end;  
end;
}

procedure TForm1.btn1Click(Sender: TObject);
var
  mypicdir : string;
  myidhttp : TIdHTTP;
  myurl : string;
  mysl : TStringList;
  myss : TStringStream;
  myzi : string;
  myfilename : string;
  mycount ,i : Integer;
const
  gc_url = 'http://dianshibaidu.cn/shufa/search.php?word=%s';
begin
  mypicdir := ExtractFileDir(System.ParamStr(0)) + '\pic';
  if not DirectoryExists(mypicdir) then
    CreateDir(mypicdir);

  mycount := tbl1.RecordCount;
  i := 1;
  tbl1.First;
  while not tbl1.Eof do
  begin
    tbl1.Edit;
    tbl1.FieldByName('pic').AsBoolean := False;

    myidhttp := TIdHTTP.Create(nil);
    mysl := TStringList.Create;
    myss := TStringStream.Create('');
    try
      myzi  := tbl1.FieldByName('zi').AsString;
      DelTree(mypicdir+'/' + myzi);
      
      myurl := Format(gc_url,[myzi{'字'}]);
      myidhttp.request.AcceptCharSet:='utf8';
      myidhttp.Request.AcceptEncoding := 'utf8';

      try
        myidhttp.Get(TIdURI.PathEncode(UTF8Encode(myurl)),myss);
      except
        tbl1.Post;
        tbl1.Next;
        Continue;
      end;

      if (myidhttp.ResponseCode = 200) and (myss.Size>0) then
      begin
        //分解文件内容。
        if not DirectoryExists(mypicdir+'\' + myzi) then
          CreateDir(mypicdir+'\' + myzi);
        mysl.Text := Utf8ToAnsi(myss.DataString);
        myfilename := Format('%s\%s\%s.html',[mypicdir,myzi,myzi]);
        mysl.SaveToFile(myfilename);
        if parsehtml(myfilename) then
        begin
          tbl1.FieldByName('pic').AsBoolean := True;  
        end
        else
          mmo2.Lines.Add(tbl1.fieldbyName('zi').AsString);
      end;
    finally
      myidhttp.Free;
      mysl.Free;
      myss.Free;
    end;

    tbl1.Post;
    if chk1.Checked then Break;

    lbl1.Caption := Format('%d/%d',[i,mycount]);
    Inc(i);
    Application.ProcessMessages;
    tbl1.Next;
  end;
end;

end.
