unit ExportCSV;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons, ExtCtrls, Menus, db, Oracle, OracleData, RXCtrls, Animate, GIFCtrl, ImgList;

type
  TFmExportCSV = class(TForm)
    Panel1: TPanel;
    OracleQuery1: TOracleQuery;
    Image1: TImage;
    RxGIFAnimator1: TRxGIFAnimator;
    Image2: TImage;
    Label1: TLabel;
    Label2: TLabel;
    ImageList1: TImageList;
    BitBtn1: TBitBtn;
    procedure OracleQuery1ThreadExecuted(Sender: TOracleQuery);
    procedure OracleQuery1ThreadFinished(Sender: TOracleQuery);
    procedure OracleQuery1ThreadRecord(Sender: TOracleQuery);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormCreate(Sender: TObject);
  private
    FFilename: string;
    FFieldNames: array of string;
    FHeader: string;
    FS: TFileStream;
  public
    { Public declarations }
  end;

procedure OracleDataSetToCSV(const ADataSet: TOracleDataSet; const AFilename: string; const AFieldNames: array of string);

implementation

{$R *.DFM}

const
  SEPARATOR_CSV = ';';

procedure OracleDataSetToCSV(const ADataSet: TOracleDataSet; const AFilename: string; const AFieldNames: array of string);
var
  LC: Integer;
begin
  with TFmExportCSV.Create(Application) do
  try
    OracleQuery1.Session := ADataSet.Session;
    OracleQuery1.SQL.Assign(ADataSet.SQL);

    FFilename := AFilename;
    SetLength(FFieldNames, Length(AFieldNames));
    for LC := 0 to Length(AFieldNames)-1 do
      FFieldNames[LC] := AFieldNames[LC];
    if Length(FFieldNames) > 0 then
    begin
      FHeader := '';
      for LC := 0 to Length(FFieldNames)-1 do
        FHeader := FHeader + ADataSet.FieldByName(FFieldNames[LC]).DisplayLabel + SEPARATOR_CSV;
      FHeader := FHeader + #13#10;

      RxGIFAnimator1.Animate := True;
      OracleQuery1.Execute;
      ShowModal;
    end;
  finally
    Free;
  end;
end;

procedure TFmExportCSV.OracleQuery1ThreadExecuted(Sender: TOracleQuery);
var
  LC: Integer;
  Linia: string;
begin
  FS := TFileStream.Create(FFilename, fmCreate);
  FS.Size := 0;
  FS.Position := 0;

  if Length(FHeader) > 0 then
    FS.Write(PChar(FHeader)^, Length(FHeader));
end;

procedure TFmExportCSV.OracleQuery1ThreadFinished(Sender: TOracleQuery);
var
  TmpIcon: TIcon;
begin
  FS.Free;

  RxGIFAnimator1.Animate := False;
  RxGIFAnimator1.Hide;
  Label2.Hide;
  Label1.Caption := 'Data export finished';
  BitBtn1.Show;
  TmpIcon := TIcon.Create;
  try
    ImageList1.GetIcon(2, TmpIcon);
    Image1.Picture.Icon := TmpIcon;
    ImageList1.GetIcon(3, TmpIcon);
    Image2.Picture.Icon := TmpIcon;
  finally
    TmpIcon.Free;
  end;
end;

procedure TFmExportCSV.OracleQuery1ThreadRecord(Sender: TOracleQuery);
var
  LC: Integer;
  Line: string;
begin
  Line := '';
  for LC := 0 to Length(FFieldNames)-1 do
    Line := Line + Sender.FieldAsString(FFieldNames[LC]) + SEPARATOR_CSV;

  if (Sender.RowsProcessed and 63) = 63 then
  begin
    Label2.Caption := 'Processed records: ' + IntToStr(Sender.RowsProcessed);
    Label2.Refresh;
  end;

  Line := Line + #13#10;
  FS.Write(PChar(Line)^, Length(Line));
end;

procedure TFmExportCSV.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  CanClose := ModalResult <> mrNone;
end;

procedure TFmExportCSV.FormCreate(Sender: TObject);
var
  TmpIcon: TIcon;
begin
  TmpIcon := TIcon.Create;
  try
    ImageList1.GetIcon(0, TmpIcon);
    Image1.Picture.Icon := TmpIcon;
    ImageList1.GetIcon(1, TmpIcon);
    Image2.Picture.Icon := TmpIcon;
  finally
    TmpIcon.Free;
  end;
end;

end.
