object frmMain: TfrmMain
  Left = 0
  Top = 0
  Caption = #1069#1082#1089#1087#1086#1088#1090#1077#1088' '#1073#1072#1079#1099' '#1076#1072#1085#1085#1099#1093' '#1055#1054' MineSCADA'
  ClientHeight = 349
  ClientWidth = 527
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  PixelsPerInch = 96
  TextHeight = 13
  object dOPCServer: TdOPCServer
    Active = False
    ClientName = 'dOPC DA Client'
    KeepAlive = 0
    Version = '4.32 trial'
    Protocol = coCOM
    Params.Strings = (
      'xml-user='
      'xml-pass='
      'xml-proxy=')
    OPCGroups = <>
    OPCGroupDefault.IsActive = True
    OPCGroupDefault.UpdateRate = 1000
    OPCGroupDefault.LocaleId = 0
    OPCGroupDefault.TimeBias = 0
    ConnectDelay = 300
    OnDatachange = dOPCServerDatachange
    Left = 24
    Top = 16
  end
  object ADOConnection: TADOConnection
    ConnectionString = 
      'Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security In' +
      'fo=False;Initial Catalog=RADatabase;Data Source=127.0.0.1'
    LoginPrompt = False
    Mode = cmReadWrite
    Provider = 'SQLOLEDB.1'
    Left = 104
    Top = 16
  end
  object tObjects: TADOTable
    Connection = ADOConnection
    CursorType = ctStatic
    TableName = 'objects'
    Left = 104
    Top = 80
  end
  object tSensors: TADOTable
    Connection = ADOConnection
    CursorType = ctStatic
    TableName = 'sensors'
    Left = 104
    Top = 128
  end
  object tData: TADOTable
    Connection = ADOConnection
    CursorType = ctStatic
    TableName = 'data'
    Left = 104
    Top = 176
  end
  object tCurData: TADOTable
    Connection = ADOConnection
    CursorType = ctStatic
    TableName = 'curdata'
    Left = 104
    Top = 224
  end
  object tSensTypes: TADOTable
    Connection = ADOConnection
    TableName = 'stypes'
    Left = 104
    Top = 272
  end
end
