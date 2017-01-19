unit Unit10;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, Grids, AdvObj, BaseGrid, AdvGrid, ComCtrls, Buttons,
  StdCtrls, tmsAdvGridExcel, DB, MemDS, DBAccess, MSAccess,DateUtils;

type
  TForm10 = class(TForm)
    Panel2: TPanel;
    Panel1: TPanel;
    grid1: TAdvStringGrid;
    SpeedButton1: TSpeedButton;
    Panel3: TPanel;
    Panel4: TPanel;
    SpeedButton2: TSpeedButton;
    SpeedButton3: TSpeedButton;
    Excel1: TAdvGridExcelIO;
    Save1: TSaveDialog;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    QrFechas: TMSQuery;
    procedure SpeedButton1Click(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
    procedure SpeedButton3Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form10: TForm10;

implementation

{$R *.dfm}

uses unit1,unit2;

procedure TForm10.SpeedButton1Click(Sender: TObject);
var
x,y,yearActual,YearLast,mes1n,mes2n,mes3n: integer;
prop : single;
mes1,mes2,mes3,ano1,ano2,ano3 : string;
begin
grid1.ClearRows(1,grid1.RowCount-1);
grid1.RowCount:=2;
y:=1;

qrFechas.Close;
qrfechas.ParamByName('FechaActual').AsDate:=date;
qrfechas.Open;

if qrfechas.FieldByName('mesactual').AsString='Enero' then
 begin
   mes1:=qrfechas.FieldByName('mespasado1').AsString;
   mes2:=qrfechas.FieldByName('mespasado2').AsString;
   mes3:=qrfechas.FieldByName('mespasado3').AsString;
   ano1:=inttostr(qrfechas.FieldByName('anopasado').AsInteger);
   ano2:=inttostr(qrfechas.FieldByName('anopasado').AsInteger);
   ano3:=inttostr(qrfechas.FieldByName('anopasado').AsInteger);
   mes1n:=qrfechas.FieldByName('mespasado1n').AsInteger;
   mes2n:=qrfechas.FieldByName('mespasado2n').AsInteger;
   mes3n:=qrfechas.FieldByName('mespasado3n').AsInteger;
 end;

if qrfechas.FieldByName('mesactual').AsString='Febrero' then
 begin
   mes1:=qrfechas.FieldByName('mespasado1').AsString;
   mes2:=qrfechas.FieldByName('mespasado2').AsString;
   mes3:=qrfechas.FieldByName('mespasado3').AsString;
   ano1:=inttostr(qrfechas.FieldByName('anoactual').AsInteger);
   ano2:=inttostr(qrfechas.FieldByName('anopasado').AsInteger);
   ano3:=inttostr(qrfechas.FieldByName('anopasado').AsInteger);
   mes1n:=qrfechas.FieldByName('mespasado1n').AsInteger;
   mes2n:=qrfechas.FieldByName('mespasado2n').AsInteger;
   mes3n:=qrfechas.FieldByName('mespasado3n').AsInteger;
 end;

if qrfechas.FieldByName('mesactual').AsString='Marzo' then
 begin
   mes1:=qrfechas.FieldByName('mespasado1').AsString;
   mes2:=qrfechas.FieldByName('mespasado2').AsString;
   mes3:=qrfechas.FieldByName('mespasado3').AsString;
   ano1:=inttostr(qrfechas.FieldByName('anoactual').AsInteger);
   ano2:=inttostr(qrfechas.FieldByName('anoactual').AsInteger);
   ano3:=inttostr(qrfechas.FieldByName('anopasado').AsInteger);
   mes1n:=qrfechas.FieldByName('mespasado1n').AsInteger;
   mes2n:=qrfechas.FieldByName('mespasado2n').AsInteger;
   mes3n:=qrfechas.FieldByName('mespasado3n').AsInteger;
 end;

if (qrfechas.FieldByName('mesactual').AsString<>'Enero') and (qrfechas.FieldByName('mesactual').AsString<>'Febrero') and (qrfechas.FieldByName('mesactual').AsString<>'Marzo') then
begin
   mes1:=qrfechas.FieldByName('mespasado1').AsString;
   mes2:=qrfechas.FieldByName('mespasado2').AsString;
   mes3:=qrfechas.FieldByName('mespasado3').AsString;
   ano1:=inttostr(qrfechas.FieldByName('anoactual').AsInteger);
   ano2:=inttostr(qrfechas.FieldByName('anoactual').AsInteger);
   ano3:=inttostr(qrfechas.FieldByName('anoactual').AsInteger);
   mes1n:=qrfechas.FieldByName('mespasado1n').AsInteger;
   mes2n:=qrfechas.FieldByName('mespasado2n').AsInteger;
   mes3n:=qrfechas.FieldByName('mespasado3n').AsInteger;
end;




form1.qrmetas.Close;
form1.qrmetas.ParamByName('tienda').AsString:=form1.combobox1.Text;
form1.qrmetas.Open;

for x := 1 to 3 do
  begin
    while not form1.qrmetas.Eof do
     begin
      grid1.cells[0,y]:=form1.qrmetas.FieldByName('nombre').AsString;
      if x=1 then begin grid1.Cells[1,y]:=mes1+' '+ano1; grid1.Ints[7,y]:=mes1n; grid1.Cells[8,y]:=ano1 end;
      if x=2 then begin grid1.Cells[1,y]:=mes2+' '+ano2; grid1.Ints[7,y]:=mes2n; grid1.Cells[8,y]:=ano2 end;
      if x=3 then begin grid1.Cells[1,y]:=mes3+' '+ano3; grid1.Ints[7,y]:=mes3n; grid1.Cells[8,y]:=ano3 end;
      grid1.RowCount:=grid1.RowCount+1;
      inc(y);
      form1.qrmetas.Next;
     end;
    form1.QrMetas.First;
  end;
    grid1.RowCount:=grid1.RowCount-1;
  //formato de celdas
  for y := 1 to grid1.RowCount-1 do
    for x := 1 to grid1.ColCount-1 do
     begin
       grid1.Alignments[x,y]:=tarightjustify;
     end;
   for y := 1 to grid1.RowCount-1 do
     begin
       form2.qrGenerales.Close;
       form2.qrGenerales.ParamByName('nombre').AsString:=grid1.Cells[0,y];
       form2.qrGenerales.Open;


       form2.QrCruzado1.close;
       form2.QrCruzado1.ParamByName('Tienda').AsString:=form2.QrGenerales.FieldByName('Factura').AsString;
       form2.QrCruzado1.ParamByName('Fecha1').AsDate:=encodedate(grid1.Ints[8,y],grid1.Ints[7,y],1);
       form2.QrCruzado1.ParamByName('Fecha2').AsDate:=encodedate(grid1.Ints[8,y],grid1.Ints[7,y],daysinamonth(grid1.Ints[8,y],grid1.Ints[7,y]));
       form2.QrCruzado1.ParamByName('Usuario').AsString:=form2.QrGenerales.FieldByName('Usuario').AsString;
       form2.QrCruzado1.Open;

       form2.QrCruzado2.close;
       form2.QrCruzado2.ParamByName('Tienda').AsString:=form2.QrGenerales.FieldByName('Factura').AsString;
       form2.QrCruzado2.ParamByName('Fecha1').AsDate:=encodedate(grid1.Ints[8,y],grid1.Ints[7,y],1);
       form2.QrCruzado2.ParamByName('Fecha2').AsDate:=encodedate(grid1.Ints[8,y],grid1.Ints[7,y],daysinamonth(grid1.Ints[8,y],grid1.Ints[7,y]));
       form2.QrCruzado2.ParamByName('Usuario').AsString:=form2.QrGenerales.FieldByName('Usuario').AsString;
       form2.QrCruzado2.Open;

       form2.QrCruzado3.close;
       form2.QrCruzado3.ParamByName('Tienda').AsString:=form2.QrGenerales.FieldByName('Factura').AsString;
       form2.QrCruzado3.ParamByName('Fecha1').AsDate:=encodedate(grid1.Ints[8,y],grid1.Ints[7,y],1);
       form2.QrCruzado3.ParamByName('Fecha2').AsDate:=encodedate(grid1.Ints[8,y],grid1.Ints[7,y],daysinamonth(grid1.Ints[8,y],grid1.Ints[7,y]));
       form2.QrCruzado3.ParamByName('Usuario').AsString:=form2.QrGenerales.FieldByName('Usuario').AsString;
       form2.QrCruzado3.Open;

       form2.QrCruzado4.close;
       form2.QrCruzado4.ParamByName('Tienda').AsString:=form2.QrGenerales.FieldByName('Factura').AsString;
       form2.QrCruzado4.ParamByName('Fecha1').AsDate:=encodedate(grid1.Ints[8,y],grid1.Ints[7,y],1);
       form2.QrCruzado4.ParamByName('Fecha2').AsDate:=encodedate(grid1.Ints[8,y],grid1.Ints[7,y],daysinamonth(grid1.Ints[8,y],grid1.Ints[7,y]));
       form2.QrCruzado4.ParamByName('Usuario').AsString:=form2.QrGenerales.FieldByName('Usuario').AsString;
       form2.QrCruzado4.Open;


       if form2.QrCruzado1.isempty=true then grid1.ints[2,y]:=0 else
       grid1.ints[2,y]:=form2.QrCruzado1.FieldByName('Facturas').AsInteger;

       if form2.qrCruzado2.isempty=true then grid1.ints[3,y]:=0 else
       grid1.ints[3,y]:=form2.qrCruzado2.FieldByName('Facturas').AsInteger;

       if form2.qrCruzado3.isempty=true then grid1.ints[4,y]:=0 else
       grid1.ints[4,y]:=form2.qrCruzado3.FieldByName('Facturas').AsInteger;

       if form2.qrCruzado4.isempty=true then grid1.ints[4,y]:=grid1.ints[4,y]+0 else
       grid1.ints[4,y]:=grid1.ints[4,y]+form2.qrCruzado4.FieldByName('Facturas').AsInteger;


       //comienzan calculos

       if grid1.Floats[2,y]=0  then grid1.Floats[5,y]:=0 else
       grid1.floats[5,y]:=(grid1.Floats[3,y]*100)/grid1.Floats[2,y];

       if grid1.Floats[2,y]=0  then grid1.Floats[6,y]:=0 else
       grid1.floats[6,y]:=(grid1.Floats[4,y]*100)/grid1.Floats[2,y];


     end;

     grid1.HideColumns(7,8);
     Grid1.SortIndexes.Clear;
     Grid1.SortIndexes.Add(0);
     Grid1.SortIndexes.Add(7);
     Grid1.SortIndexes.Add(8);
     Grid1.QSortIndexed;

    

end;

procedure TForm10.SpeedButton2Click(Sender: TObject);
begin
Form10.Visible:=false;
form1.Visible:=true;
end;

procedure TForm10.SpeedButton3Click(Sender: TObject);
begin
save1.Execute;
Excel1.XLSExport(save1.FileName);

showmessage('Reporte Exportado con Exito!');

end;

end.
