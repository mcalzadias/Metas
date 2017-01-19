unit Unit7;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, AdvObj, BaseGrid, AdvGrid, ExtCtrls, StdCtrls, Buttons,
  ComCtrls, tmsAdvGridExcel,dateutils;

type
  TForm7 = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    Grid1: TAdvStringGrid;
    SpeedButton1: TSpeedButton;
    Panel3: TPanel;
    Panel4: TPanel;
    SpeedButton2: TSpeedButton;
    SpeedButton3: TSpeedButton;
    Excel1: TAdvGridExcelIO;
    Save1: TSaveDialog;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    procedure SpeedButton1Click(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
    procedure SpeedButton3Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form7: TForm7;

implementation

{$R *.dfm}

uses unit1,unit2,unit6;

procedure TForm7.SpeedButton1Click(Sender: TObject);
var
 x,y : integer;
 yearActual,YearLast,mes1n,mes2n,mes3n: integer;
 mes1,mes2,mes3,ano1,ano2,ano3 : string;
prop : single;
begin
grid1.ClearRows(1,grid1.RowCount-1);
grid1.RowCount:=2;
y:=1;

form6.qrFechas.Close;
form6.qrfechas.ParamByName('FechaActual').AsDate:=date;
form6.qrfechas.Open;

if form6.qrfechas.FieldByName('mesactual').AsString='Enero' then
 begin
   mes1:=form6.qrfechas.FieldByName('mespasado1').AsString;
   mes2:=form6.qrfechas.FieldByName('mespasado2').AsString;
   mes3:=form6.qrfechas.FieldByName('mespasado3').AsString;
   ano1:=inttostr(form6.qrfechas.FieldByName('anopasado').AsInteger);
   ano2:=inttostr(form6.qrfechas.FieldByName('anopasado').AsInteger);
   ano3:=inttostr(form6.qrfechas.FieldByName('anopasado').AsInteger);
   mes1n:=form6.qrfechas.FieldByName('mespasado1n').AsInteger;
   mes2n:=form6.qrfechas.FieldByName('mespasado2n').AsInteger;
   mes3n:=form6.qrfechas.FieldByName('mespasado3n').AsInteger;
 end;

if form6.qrfechas.FieldByName('mesactual').AsString='Febrero' then
 begin
   mes1:=form6.qrfechas.FieldByName('mespasado1').AsString;
   mes2:=form6.qrfechas.FieldByName('mespasado2').AsString;
   mes3:=form6.qrfechas.FieldByName('mespasado3').AsString;
   ano1:=inttostr(form6.qrfechas.FieldByName('anoactual').AsInteger);
   ano2:=inttostr(form6.qrfechas.FieldByName('anopasado').AsInteger);
   ano3:=inttostr(form6.qrfechas.FieldByName('anopasado').AsInteger);
   mes1n:=form6.qrfechas.FieldByName('mespasado1n').AsInteger;
   mes2n:=form6.qrfechas.FieldByName('mespasado2n').AsInteger;
   mes3n:=form6.qrfechas.FieldByName('mespasado3n').AsInteger;
 end;

if form6.qrfechas.FieldByName('mesactual').AsString='Marzo' then
 begin
   mes1:=form6.qrfechas.FieldByName('mespasado1').AsString;
   mes2:=form6.qrfechas.FieldByName('mespasado2').AsString;
   mes3:=form6.qrfechas.FieldByName('mespasado3').AsString;
   ano1:=inttostr(form6.qrfechas.FieldByName('anoactual').AsInteger);
   ano2:=inttostr(form6.qrfechas.FieldByName('anoactual').AsInteger);
   ano3:=inttostr(form6.qrfechas.FieldByName('anopasado').AsInteger);
   mes1n:=form6.qrfechas.FieldByName('mespasado1n').AsInteger;
   mes2n:=form6.qrfechas.FieldByName('mespasado2n').AsInteger;
   mes3n:=form6.qrfechas.FieldByName('mespasado3n').AsInteger;
 end;

if (form6.qrfechas.FieldByName('mesactual').AsString<>'Enero') and (form6.qrfechas.FieldByName('mesactual').AsString<>'Febrero') and (form6.qrfechas.FieldByName('mesactual').AsString<>'Marzo') then
begin
   mes1:=form6.qrfechas.FieldByName('mespasado1').AsString;
   mes2:=form6.qrfechas.FieldByName('mespasado2').AsString;
   mes3:=form6.qrfechas.FieldByName('mespasado3').AsString;
   ano1:=inttostr(form6.qrfechas.FieldByName('anoactual').AsInteger);
   ano2:=inttostr(form6.qrfechas.FieldByName('anoactual').AsInteger);
   ano3:=inttostr(form6.qrfechas.FieldByName('anoactual').AsInteger);
   mes1n:=form6.qrfechas.FieldByName('mespasado1n').AsInteger;
   mes2n:=form6.qrfechas.FieldByName('mespasado2n').AsInteger;
   mes3n:=form6.qrfechas.FieldByName('mespasado3n').AsInteger;
end;




form1.qrmetas.Close;
form1.qrmetas.ParamByName('tienda').AsString:=form1.combobox1.Text;
form1.qrmetas.Open;

for x := 1 to 3 do
begin
 while not form1.qrmetas.Eof do
  begin
    grid1.cells[0,y]:=form1.qrmetas.FieldByName('nombre').AsString;
    if x=1 then begin grid1.Cells[1,y]:=mes1+' '+ano1; grid1.Ints[5,y]:=mes1n; grid1.Cells[6,y]:=ano1 end;
    if x=2 then begin grid1.Cells[1,y]:=mes2+' '+ano2; grid1.Ints[5,y]:=mes2n; grid1.Cells[6,y]:=ano2 end;
    if x=3 then begin grid1.Cells[1,y]:=mes3+' '+ano3; grid1.Ints[5,y]:=mes3n; grid1.Cells[6,y]:=ano3 end;
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

       form2.qrventasOT.close;
       form2.QrVentasOT.ParamByName('DevVenta').AsString:=form2.QrGenerales.FieldByName('Devolucion').AsString;
       form2.QrVentasOT.ParamByName('Tienda').AsString:=form2.QrGenerales.FieldByName('Factura').AsString;
       form2.QrVentasOT.ParamByName('Fecha1').AsDate:=encodedate(grid1.Ints[6,y],grid1.Ints[5,y],1);
       form2.QrVentasOT.ParamByName('Fecha2').AsDate:=encodedate(grid1.Ints[6,y],grid1.Ints[5,y],daysinamonth(grid1.Ints[6,y],grid1.Ints[5,y]));
       form2.QrVentasOT.ParamByName('Usuario').AsString:=form2.QrGenerales.FieldByName('Usuario').AsString;
       form2.QrVentasOT.Open;

       form2.qrventasCBHD.close;
       form2.QrVentasCBHD.ParamByName('DevVenta').AsString:=form2.QrGenerales.FieldByName('Devolucion').AsString;
       form2.QrVentasCBHD.ParamByName('Tienda').AsString:=form2.QrGenerales.FieldByName('Factura').AsString;
       form2.QrVentasCBHD.ParamByName('Fecha1').AsDate:=encodedate(grid1.Ints[6,y],grid1.Ints[5,y],1);
       form2.QrVentasCBHD.ParamByName('Fecha2').AsDate:=encodedate(grid1.Ints[6,y],grid1.Ints[5,y],daysinamonth(grid1.Ints[6,y],grid1.Ints[5,y]));
       form2.QrVentasCBHD.ParamByName('Usuario').AsString:=form2.QrGenerales.FieldByName('Usuario').AsString;
       form2.QrVentasCBHD.Open;

       form2.qrventasE.close;
       form2.QrVentasE.ParamByName('DevVenta').AsString:=form2.QrGenerales.FieldByName('Devolucion').AsString;
       form2.QrVentasE.ParamByName('Tienda').AsString:=form2.QrGenerales.FieldByName('Factura').AsString;
       form2.QrVentasE.ParamByName('Fecha1').AsDate:=encodedate(grid1.Ints[6,y],grid1.Ints[5,y],1);
       form2.QrVentasE.ParamByName('Fecha2').AsDate:=encodedate(grid1.Ints[6,y],grid1.Ints[5,y],daysinamonth(grid1.Ints[6,y],grid1.Ints[5,y]));
       form2.QrVentasE.ParamByName('Usuario').AsString:=form2.QrGenerales.FieldByName('Usuario').AsString;
       form2.QrVentasE.Open;

       if form2.qrVentasOT.isempty=true then grid1.ints[2,y]:=0 else
       grid1.ints[2,y]:=form2.qrVentasOT.FieldByName('Cantidad').AsInteger;

       if form2.qrVentasCBHD.isempty=true then grid1.ints[3,y]:=0 else
       grid1.ints[3,y]:=form2.qrVentasCBHD.FieldByName('Cantidad').AsInteger;

       if form2.qrVentasE.isempty=true then grid1.ints[3,y]:=grid1.ints[3,y]+0 else
       grid1.ints[3,y]:=grid1.ints[3,y]+form2.qrVentasE.FieldByName('Cantidad').AsInteger;

       // comienzan calculos
       if grid1.Ints[2,y]=0 then grid1.Floats[4,y]:=0 else
       grid1.Floats[4,y]:=(grid1.Floats[3,y]/(grid1.Floats[2,y]+grid1.Floats[3,y]))*100;

     end;



     for y := 1 to grid1.RowCount-1 do
   begin
     grid1.Cells[4,y]:=FormatFloat('##0.0%',grid1.Floats[4,y]);
   end;

     grid1.HideColumns(5,6);
     Grid1.SortIndexes.Clear;
     Grid1.SortIndexes.Add(0);
     Grid1.SortIndexes.Add(5);
     Grid1.SortIndexes.Add(6);
     Grid1.QSortIndexed;


end;

procedure TForm7.SpeedButton2Click(Sender: TObject);
begin
Form7.Visible:=false;
form1.Visible:=true;
end;

procedure TForm7.SpeedButton3Click(Sender: TObject);
begin
save1.Execute;
Excel1.XLSExport(save1.FileName,'Datos');

showmessage('Reporte Exportado con Exito!');
end;

end.
