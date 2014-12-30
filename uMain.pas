unit uMain;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, ZAbstractConnection, ZConnection, Registry,
  dOPCIntf, dOPCComn, dOPCDA, dOPC, IniFiles, RegExpr, TLHELP32, Vcl.StdCtrls, math,
  Data.DB, Data.Win.ADODB;

type
  TValue = record
    val   : double;
    dt    : TDateTime;
  end;

  PTTecObject = ^TTechObject;

  TSensor = record
    id          : integer;
    pObject     : PTTecObject;
    oid         : integer;
    address     : string[13];
    name        : string;
    low         : byte;
    high        : byte;
    raw         : double;
    units       : string;
    isMaxRaw    : boolean;
    sensType    : byte;
    value       : TValue;
    dt          : TDateTime;
  end;

  TSensors = array of TSensor;

  TTechObject = record
    id          : integer;
    address     : string[7];
    name        : string;
    dt          : TDateTime;
    node        : byte;
  end;

  TTechObjects = array of TTechObject;

  TconfigInfo = record
    path : string;
    dt   : TDateTime;
  end;

  //TSensTypes = set of byte;
  TSensType = record
    confId    : integer;
    name      : string;
    lowAlarms : boolean;
  end;
  TSensTypes = array of TSensType;

  TMainClass = class
    private
      //settings : TIniFile;
      //config   : TMemIniFile;
      function prepareObjAddress(sap : byte; ring : byte; number : byte) : string;
      function rawToDouble(input : string) : double;
      function getActiveConf : TconfigInfo;
      function getConfAvailablity : boolean;
      function getConfLastChangeDateTime(path : string) : TDateTime;
      procedure confAnalyze;
      procedure fillSensTypes(var sTypes : TSensTypes; settings : TIniFile);
      //procedure fillAirTypes(var airTypes : TSensTypes; settings : TIniFile);
      procedure configOPCServer(OPCServer : TdOPCServer);
      procedure resetOPC(OPCServer : TdOPCServer);
      function checkMSOPC : boolean;
      procedure initOPC(var s : TSensors; OPCServer : TdOPCServer);
    public
      S : TSensors;
      O : TTechObjects;
      constructor Create;
      destructor Destroy;
      property activeConf : TconfigInfo read getActiveConf; //находит доступный файл конфигурации на серверах
      property confAvailable : boolean read getConfAvailablity; //доступность файла конфигурации, проверять при обращении
      property MSOPCisOffline : boolean read checkMSOPC;
  end;

  {TDatabase = class
    private
      procedure readInformation(var O : TTechObjects; var S : TSensors);
    public
      constructor Create;
      destructor Destroy;
  end;}

  TfrmMain = class(TForm)
    dOPCServer: TdOPCServer;
    ADOConnection: TADOConnection;
    tObjects: TADOTable;
    tSensors: TADOTable;
    tData: TADOTable;
    tCurData: TADOTable;
    tSensTypes: TADOTable;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure dOPCServerDatachange(Sender: TObject; ItemList: TdOPCItemList);
  private
    { Private declarations }
    function prepareConnection : boolean;
  public
    { Public declarations }
    dbS : TSensors;
    dbO : TTechObjects;
    procedure getValueFromOPC(var sens : TSensor);
    procedure dbAnalyze;
    property connectionReady : boolean read prepareConnection;
  end;

var
  frmMain   : TfrmMain;
  MainClass : TMainClass;
  //Database  : TDatabase;

implementation

{$R *.dfm}

function TfrmMain.prepareConnection;
begin
  ADOConnection.Close;
  ADOConnection.ConnectionString := 'Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=RADatabase;Data Source=127.0.0.1';
  try
    ADOConnection.Open;
    Result := true;
  except
    Result := false;
  end;
end;

procedure TfrmMain.dbAnalyze;
var
  i, j : integer;
  //dbO  : TTechObjects;
  //dbS  : TSensors;
begin
  if connectionReady then
  begin
    tObjects.Close;
    tObjects.Open;
    tSensors.Close;
    tSensors.Open;
    if (tObjects.IsEmpty) then
    begin
      with MainClass do
      begin
        for i := 0 to length(O) - 1 do
        begin
          tObjects.AppendRecord([nil,  O[i].name, O[i].address, now]);
          O[i].id := tObjects.FieldByName('id').AsInteger;
          O[i].dt := tObjects.FieldByName('dt').AsFloat;
        end;
        for i := 0 to length(S) - 1 do
        begin
          tSensors.AppendRecord([nil, S[i].pObject.id, S[i].name, S[i].address,
                                 S[i].low, S[i].high, S[i].raw, S[i].units, now]);
          S[i].id := tSensors.FieldByName('id').AsInteger;
          S[i].dt := tSensors.FieldByName('dt').AsDateTime;
        end;
      end;
    end
    else begin //таблицы не пустые, сравнение данных
      SetLength(dbO, tObjects.RecordCount);
      i := 0;
      tObjects.First;
      while not tObjects.Eof do
      begin
        dbO[i].id := tObjects.FieldByName('id').AsInteger;
        dbO[i].address := tObjects.FieldByName('address').AsString;
        dbO[i].name := tObjects.FieldByName('name').AsString;
        dbO[i].dt := tObjects.FieldByName('dt').AsFloat;
        inc(i);
        tObjects.Next;
      end;
      SetLength(dbS, tSensors.RecordCount);
      i := 0;
      while not tSensors.Eof do
      begin
        dbS[i].id       := tSensors.FieldByName('id').AsInteger;
        dbS[i].oid      := tSensors.FieldByName('oid').AsInteger;
        dbS[i].address  := tSensors.FieldByName('address').AsString;
        dbS[i].name     := tSensors.FieldByName('name').AsString;
        dbS[i].low      := tSensors.FieldByName('low').AsInteger;
        dbS[i].high     := tSensors.FieldByName('high').AsInteger;
        dbS[i].raw      := tSensors.FieldByName('raw').AsFloat;
        dbS[i].units    := tSensors.FieldByName('units').AsString;
        dbS[i].senstype := tSensors.FieldByName('stype').AsInteger;
      end;
    end;
  end
  else begin
    MessageDlg('Не удалось соединиться с сервером баз данных. Проверьте настройки подключения.', mtError, [mbOk], 0);
    exit;
  end;
end;

{procedure TDatabase.readInformation(var O : TTechObjects; var S : TSensors);
var
  i, j : integer;
begin
  with frmMain do
  begin
    tObjects.Close;
    tObjects.Open;

    tSensors.Close;
    tSensors.Open;

    if tObjects.IsEmpty and tSensors.IsEmpty then
    begin
      for i := 0 to length(O) - 1 do
      begin
        tObjects.AppendRecord([nil, O[i].address, O[i].name, now]);
        O[i].id := tObjects.FieldByName('id').AsInteger;
        O[i].dt := tObjects.FieldByName('dt').AsDateTime;
      end;
      for i := 0 to length(S) - 1 do
      begin
        tSensors.AppendRecord([nil, S[i].pObject.id, S[i].name, S[i].address,
                               S[i].low, S[i].high, S[i].raw, S[i].units, now]);
        S[i].id := tSensors.FieldByName('id').AsInteger;
        S[i].dt := tSensors.FieldByName('dt').AsDateTime;
      end;
    end else
    begin
      //записи существуют
    end;
  end;
end;
}
{constructor TDatabase.Create;
begin
  inherited;
  with frmMain.ADOConnection do
  begin
    Close;
    ConnectionString := 'Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=RADatabase;Data Source=127.0.0.1';
    try
      Open;
    except
      MessageDlg('Не удалось соединиться с сервером баз данных. Проверьте настройки подключения.', mtError, [mbOk], 0);
      exit;
    end;
  end;
  with frmMain do
  begin
    tObjects.TableName := 'objects';
    tSensors.TableName := 'sensors';
    tData.TableName    := 'data';
    tCurData.TableName := 'curdata';
  end;
  readInformation(MainClass.O, MainClass.S);
end;

destructor TDatabase.Destroy;
begin
  frmMain.ADOConnection.Close;
  inherited;
end;
}
procedure TfrmMain.getValueFromOPC(var sens : TSensor);
var
  i       : integer;
  props   : TdOPCItemProperties;
begin
  props := TdOPCItemProperties.create(dOPCServer, sens.address);
  try
    for i := 0 to props.Count - 1 do
    if props[i].Id = 2 then
    begin
      sens.value.val := RoundTo(double(props[i].Value), -2);
      sens.value.dt  := now;
      //добавить данные в БД
    end;
  finally
    props.Free;
  end;
end;

procedure TMainClass.initOPC(var s : TSensors; OPCServer : TdOPCServer);
var
  MyOPCGroup : TdOPCGroup;
  i, j       : integer;
  props      : TdOPCItemProperties;
begin
  OPCServer.OPCGroups.RemoveAll;
  MyOPCGroup := OPCServer.OPCGroups.Add('SCADA');
  for i := 0 to length(S) - 1 do
  begin
    MyOPCGroup.OPCItems.AddItem(S[i].address);
    props := TdOPCItemProperties.create(OPCServer, S[i].address);
    for j := 0 to props.Count - 1 do
    case props[j].Id of
      100 : s[i].units := string(props[j].Value);
      102 : s[i].high  := integer(props[j].Value);
      103 : s[i].low   := integer(props[j].Value);
    end;
  end;
  MyOPCGroup.Free;
  OPCServer.OPCGroups.RemoveAll;
end;

function TMainClass.checkMSOPC : boolean;
var
  H     : THandle;
  proc  : TProcessEntry32;
begin
  result := true;
  H := CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
  if H = INVALID_HANDLE_VALUE then exit;
  proc.dwSize := sizeof(TprocessEntry32);
  if Process32First(H, proc) then
  repeat
    if pos('MSOPCSRV', UpperCase(proc.szExeFile)) > 0 then
    begin
      result := false;
      break;
    end;
  until not process32next(H, proc);
  closehandle(H);
end;

procedure TMainClass.resetOPC(OPCServer: TdOPCServer);
begin
  //if assigned(OPCServer.OPCGroups.GetOPCGroup('SCADA')) then OPCServer.OPCGroups.RemoveAll;
  OPCServer.Active := false;
  //Проверить, выключен ли сервер OPC
  while not MSOPCisOffline do Application.ProcessMessages;
  OPCServer.Active := true;
end;

function TMainClass.rawToDouble(input : string) : double;
const
  EXPR = '^\d+(\.)\d+$';
  TEMPLATE = '$1';
begin
  if ExecRegExpr(EXPR, input) then
  begin
    result := StrToFloat(ReplaceRegExpr('\.', input, ',', false));
  end
  else result := StrToFloat(input);
end;

function TMainClass.prepareObjAddress(sap: Byte; ring: Byte; number: Byte) : string;
begin
  result := '.' + IntToStr(sap) + '_' + IntToStr(ring) + '_' + IntToStr(number);
end;

{procedure TMainClass.fillAirTypes(var airTypes: TSensTypes; settings: TIniFile);
begin
  with settings do
  begin
    airTypes := [
      ReadInteger('SensTypes', 'СДСВ', 0),
      ReadInteger('SensTypes', 'WMA', 0)
    ];
  end;
end;      }

procedure TMainClass.fillSensTypes(var sTypes: TSensTypes; settings: TIniFile);
var
  count : integer;
  i     : integer;
begin
  {with settings do
  begin
    sTypes := [
      ReadInteger('SensTypes', 'СДСВ', 0),
      ReadInteger('SensTypes', 'WMA', 0),
      ReadInteger('SensTypes', 'M', 0),
      ReadInteger('SensTypes', 'M100', 0),
      ReadInteger('SensTypes', 'CO', 0)
    ];
  end;}
  count := settings.ReadInteger('Info', 'typeCount', 0);
  SetLength(sTypes, count);
  for i := 0 to count - 1 do
  begin
    sTypes[i].confId    := settings.ReadInteger('type' + IntToStr(i), 'confId', 0);
    sTypes[i].name      := settings.ReadString('type' + IntToStr(i), 'name', '');
    sTypes[i].lowAlarms := settings.ReadBool('type' + IntToStr(i), 'lowAlarm', false);
  end;
end;

procedure TMainClass.confAnalyze;
var
  config          : TMemIniFile;
  settings        : TIniFile;
  nodeCount       : byte;
  analogCount     : integer;
  //sensTypes       : TSensTypes;
  //airTypes        : TSensTypes;
  i, j            : integer;
  objRing         : byte;
  curNode         : byte;
  sensTypes       : TSensTypes;
begin
  while not confAvailable do Application.ProcessMessages;     //проверить доступность файла конфигурации
  config := TMemIniFile.Create(activeConf.path);
  settings := TIniFile.Create(ExtractFilePath(Application.ExeName) + 'settings.ini');
  try
    nodeCount := config.ReadInteger('Header', 'Nodes', 0);
    analogCount := config.ReadInteger('Header', 'AnalIns', 0);
    //sensTypes := [];
    //airTypes := [];
    fillSensTypes(sensTypes, settings);
    //fillAirTypes(airTypes, settings);
    SetLength(O, 0);
    SetLength(S, 0);
    for i := 0 to nodeCount - 1 do
    if (config.ReadInteger('Node' + IntToStr(i), 'not connected', 0) = 0) then
    begin
      SetLength(O, length(O) + 1);
      O[length(O) - 1].name := trim(config.ReadString('Node' + IntToStr(i), 'Name', 'Отсутствует имя объекта'));
      objRing := config.ReadInteger('Node' + IntToStr(i), 'Network', 0);
      O[length(O) - 1].address := prepareObjAddress(
        config.ReadInteger('Network' + IntToStr(objRing), 'Interface', 0) + 1,
        config.ReadInteger('Network' + IntToStr(objRing), 'Number',    1),
        config.ReadInteger('Node'    + IntToStr(i),       'Number',    1)
      );
      O[length(O) - 1].node := i;
    end;

    for i  := 0 to analogCount - 1 do
    begin
      if config.ReadInteger('AnalIn' + IntToStr(i), 'Sensor', 0) in sensTypes then
      begin
        curNode := config.ReadInteger('AnalIn' + IntToStr(i), 'Node', 0);
        for j := 0 to length(O) - 1 do
        begin
          if curNode = O[j].node then
          begin
            if config.ReadInteger('AnalIn' + IntToStr(i), 'Number', 1) <> 0 then
            begin
              setlength(S, length(S) + 1);
              with S[length(S) - 1] do
              begin
                name := config.ReadString('AnalIn' + IntToStr(i), 'Name', 'Имя не указано');
                sensType := config.ReadInteger('AnalIn' + IntToStr(i), 'Sensor', 0);
                pObject  := @O[j];
                if sensType in airTypes then
                begin
                  raw := rawToDouble(config.ReadString('AnalIn' + IntToStr(i), 'Low alarm', '0'));
                  isMaxRaw := false;
                end
                else begin
                  raw := rawToDouble(config.ReadString('AnalIn' + IntTOStr(i), 'High alarm', '0'));
                  isMaxRaw := true;
                end;
                address := 'AI' + O[j].address + '_' + IntToStr(config.ReadInteger('AnalIn' + IntTostr(i), 'Number', 1));
              end;  
            end;  
          end;          
        end;
      end;
    end;
  finally
    settings.Free;
    config.Free;
  end;
end;

constructor TMainClass.Create;
begin
  inherited;
  confAnalyze;
  configOPCServer(frmMain.dOPCServer);
  resetOPC(frmMain.dOPCServer);
  initOPC(S, frmMain.dOPCServer);
end;

destructor TMainClass.Destroy;
begin
  SetLength(S, 0);
  SetLength(O, 0);
  inherited;
end;

procedure TMainClass.configOPCServer(OPCServer : TdOPCServer);
const
  OPC_SERVER_NAME = 'MineSCADA.OPCServer';
begin
  OPCServer.ComputerName := 'localhost';
  OPCServer.ServerName := OPC_SERVER_NAME;
end;

function TMainClass.getConfAvailablity;
var
  F : TFileStream;
begin
  try
    F := TFileStream.Create(activeConf.path, fmOpenReadWrite or fmShareExclusive);
    result := true;
    F.Free;
  except
    result := false;
  end;
end;

function TMainClass.getConfLastChangeDateTime(path: string) : TDateTime;
var
  FHandle : integer;
begin
  FHandle := FileOpen(path, 0);
  if FHandle > 0 then
    try
      result := FileDateToDateTime(FileGetDate(FHandle));
    finally
      FileClose(FHandle);
    end
  else result := 0;
end;

function TMainClass.getActiveConf;
type
  TPathes = record
    serv1, serv2 : string;
  end;
var
  registry   : TRegistry;
  pathes   : TPathes;
begin
  registry := TRegistry.Create;
  registry.RootKey := HKEY_CURRENT_USER;
  registry.OpenKey('SOFTWARE\MineSCADA', false);
  pathes.serv1 := registry.ReadString('Server 1 Path') + 'config\' + registry.ReadString('Config file') + '.cfg';
  pathes.serv2 := registry.ReadString('Server 2 Path') + 'config\' + registry.ReadString('Config file') + '.cfg';
  registry.Free;
  while true do
  begin
    Application.ProcessMessages;
    if FileExists(pathes.serv1) then
    begin
      result.path := pathes.serv1;
      result.dt := getConfLastChangeDateTime(pathes.serv1);
      break;
    end
    else
    if FileExists(pathes.serv2) then
    begin
      result.path := pathes.serv2;
      result.dt := getConfLastChangeDateTime(pathes.serv2);
      break;
    end;
  end;
end;

procedure TfrmMain.dOPCServerDatachange(Sender: TObject;
  ItemList: TdOPCItemList);
var
  i, j : integer;
begin
  if assigned(MainClass) then
    if (assigned(MainClass.S)) and (assigned(MainClass.O)) then
      for i := 0 to ItemList.Count - 1 do
        for j := 0 to length(MainClass.S) - 1 do
        if ItemList.Items[i].ItemID = MainClass.S[j].address then
        begin
          getValueFromOPC(MainClass.S[j]);//получить показания из OPC
          //добавить в БД значения
          break;
        end;
end;

procedure TfrmMain.FormCreate(Sender: TObject);
begin
  {try
    with Database.ADOConnection do
    begin
      Close;
      ConnectionString := 'Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=RADatabase;Data Source=127.0.0.1';
      Open;
    end;
  except
    MessageDlg('Невозможно соединиться с базой данных.', mtError, [mbOk], 0);
    exit;
  end;
  tObjects.TableName := 'objects';
  tSensors.TableName := 'sensors';
  tData.TableName    := 'data';
  tCurData.TableName := 'curdata';    }
  MainClass := TMainClass.Create;
  dbAnalyze;
  //Database  := TDatabase.Create;
end;

procedure TfrmMain.FormDestroy(Sender: TObject);
begin
  //if Assigned(Database) then Database.Destroy;
  dOPCServer.Disconnect;
  if Assigned(MainClass) then MainClass.Destroy;
end;

end.
