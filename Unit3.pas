unit Unit3;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Buttons, StdCtrls, ExtCtrls, DB, MemDS, DBAccess, MSAccess, Grids,
  AdvObj, BaseGrid, AdvGrid;

type
  TForm3 = class(TForm)
    Panel5: TPanel;
    Panel6: TPanel;
    Grid1: TAdvStringGrid;
    QrMetas: TMSQuery;
    SpeedButton2: TSpeedButton;
    SpeedButton5: TSpeedButton;
    SpeedButton6: TSpeedButton;
    UpMetas: TMSQuery;
    procedure Grid1CanEditCell(Sender: TObject; ARow, ACol: Integer;
      var CanEdit: Boolean);
    procedure SpeedButton6Click(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
    procedure SpeedButton5Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form3: TForm3;
  Menux: Integer;

implementation

{$R *.dfm}

uses unit1;

procedure TForm3.Grid1CanEditCell(Sender: TObject; ARow, ACol: Integer;
  var CanEdit: Boolean);
begin
if acol=13 then canedit:=false;

end;

procedure TForm3.SpeedButton2Click(Sender: TObject);
var
y: integer;
begin
if MessageDlg('Esta Seguro de querer guardar los datos actuales?',
    mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      for y := 1 to grid1.RowCount-1 do
        begin
          UpMetas.Close;
          upmetas.ParamByName('MetaO').AsInteger:=grid1.Ints[1,y];
          upmetas.ParamByName('MetaE').AsInteger:=grid1.Ints[2,y];
          upmetas.ParamByName('MetaHD').AsInteger:=grid1.Ints[3,y];
          upmetas.ParamByName('MetaL').AsInteger:=grid1.Ints[4,y];
          upmetas.ParamByName('MetaT').AsInteger:=grid1.Ints[5,y];
          upmetas.ParamByName('MetaC').AsInteger:=grid1.Ints[6,y];
          upmetas.ParamByName('MetaI').AsInteger:=grid1.Ints[7,y];
          upmetas.ParamByName('MetaS').AsInteger:=grid1.Ints[8,y];
          upmetas.ParamByName('MetaB').AsInteger:=grid1.Ints[9,y];
          upmetas.ParamByName('MetaPT').AsInteger:=grid1.Ints[10,y];
          upmetas.ParamByName('MetaPL').AsInteger:=grid1.Ints[11,y];
          upmetas.ParamByName('MetaQ').AsInteger:=grid1.Ints[12,y];
          upmetas.ParamByName('MetaQHD').AsInteger:=grid1.Ints[13,y];
          upmetas.ParamByName('Nombre').Asstring:=grid1.cells[0,y];
          upMetas.ExecSQL;
        end;
        speedbutton5.Enabled:=true;
        showmessage('Cambios grabados con Exito!');
    end;

end;

procedure TForm3.SpeedButton5Click(Sender: TObject);
var
x,y : integer;
begin

grid1.ClearRows(1,grid1.RowCount-1);
grid1.RowCount:=2;
y:=1;
grid1.Visible:=true;
qrmetas.Close;
qrmetas.ParamByName('tienda').AsString:=form1.combobox1.Text;
qrmetas.Open;

while not qrmetas.Eof do
  begin
    grid1.cells[0,y]:=qrmetas.FieldByName('nombre').AsString;
    grid1.ints[1,y]:=qrmetas.FieldByName('MetaUO').AsInteger;
    grid1.ints[2,y]:=qrmetas.FieldByName('MetaUE').AsInteger;
    grid1.ints[3,y]:=qrmetas.FieldByName('MetaUEHD').AsInteger;
    grid1.ints[4,y]:=qrmetas.FieldByName('MetaLL').AsInteger;
    grid1.ints[5,y]:=qrmetas.FieldByName('MetaT').AsInteger;
    grid1.ints[6,y]:=qrmetas.FieldByName('MetaC').AsInteger;
    grid1.ints[7,y]:=qrmetas.FieldByName('MetaI').AsInteger;
    grid1.ints[8,y]:=qrmetas.FieldByName('MetaS').AsInteger;
    grid1.ints[9,y]:=qrmetas.FieldByName('MetaB').AsInteger;
    grid1.ints[10,y]:=qrmetas.FieldByName('MetaPT').AsInteger;
    grid1.ints[11,y]:=qrmetas.FieldByName('MetaPL').AsInteger;
    grid1.ints[12,y]:=qrmetas.FieldByName('MetaQ').AsInteger;
    grid1.ints[13,y]:=qrmetas.FieldByName('MetaQHD').AsInteger;
    grid1.ints[14,y]:=qrmetas.FieldByName('MetaTotal').AsInteger;
    grid1.RowCount:=grid1.RowCount+1;
    inc(y);
    qrmetas.Next;
  end;
  grid1.RowCount:=grid1.RowCount-1;

  //formato de celdas
  for y := 1 to grid1.RowCount-1 do
    for x := 1 to grid1.ColCount-1 do
     begin
       grid1.Alignments[x,y]:=tarightjustify;
     end;

 // for y := 1 to grid1.RowCount-1 do
    // grid1.Cells[13,y]:=FormatFloat('#,##0',grid1.Floats[13,y]);

  speedbutton5.Enabled:=false;
end;

procedure TForm3.SpeedButton6Click(Sender: TObject);
begin
Form3.Visible:=false;
form1.Visible:=true;
end;

end.
