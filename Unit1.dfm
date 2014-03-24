object Form1: TForm1
  Left = 396
  Top = 107
  Width = 838
  Height = 604
  Caption = 'Form1'
  Color = clBtnFace
  Font.Charset = GB2312_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = #23435#20307
  Font.Style = []
  OldCreateOrder = False
  Position = poDesktopCenter
  PixelsPerInch = 96
  TextHeight = 13
  object lbl1: TLabel
    Left = 352
    Top = 184
    Width = 28
    Height = 13
    Caption = 'lbl1'
  end
  object dbgrd1: TDBGrid
    Left = 8
    Top = 48
    Width = 801
    Height = 120
    DataSource = ds1
    TabOrder = 0
    TitleFont.Charset = GB2312_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -13
    TitleFont.Name = #23435#20307
    TitleFont.Style = []
  end
  object btn1: TButton
    Left = 8
    Top = 176
    Width = 153
    Height = 25
    Caption = #33719#21462#20070#27861#22270
    TabOrder = 1
    OnClick = btn1Click
  end
  object chk1: TCheckBox
    Left = 232
    Top = 184
    Width = 97
    Height = 17
    Caption = 'stop'
    TabOrder = 2
  end
  object mmo1: TMemo
    Left = 16
    Top = 216
    Width = 657
    Height = 89
    Lines.Strings = (
      'mmo1')
    TabOrder = 3
  end
  object mmo2: TMemo
    Left = 16
    Top = 320
    Width = 537
    Height = 89
    Lines.Strings = (
      'mmo2')
    TabOrder = 4
  end
  object ds1: TDataSource
    DataSet = tbl1
    Left = 32
    Top = 16
  end
  object con1: TADOConnection
    Connected = True
    ConnectionString = 
      'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\xhzd.mdb;Persist ' +
      'Security Info=False'
    LoginPrompt = False
    Mode = cmShareDenyNone
    Provider = 'Microsoft.Jet.OLEDB.4.0'
    Left = 64
    Top = 16
  end
  object tbl1: TADOTable
    Active = True
    Connection = con1
    CursorType = ctStatic
    TableName = 'xhzd_surnfu'
    Left = 104
    Top = 16
  end
  object idhtp1: TIdHTTP
    MaxLineAction = maException
    ReadTimeout = 0
    AuthRetries = 0
    AllowCookies = True
    ProxyParams.BasicAuthentication = False
    ProxyParams.ProxyPort = 0
    Request.ContentLength = -1
    Request.ContentRangeEnd = 0
    Request.ContentRangeStart = 0
    Request.ContentType = 'text/html'
    Request.Accept = 'text/html, */*'
    Request.BasicAuthentication = False
    Request.UserAgent = 'Mozilla/3.0 (compatible; Indy Library)'
    HTTPOptions = [hoForceEncodeParams]
    Left = 192
    Top = 176
  end
end
