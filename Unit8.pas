unit Unit8;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, Grids, AdvObj, BaseGrid, AdvGrid, Buttons, StdCtrls,
  ComCtrls, DB, MemDS, DBAccess, MSAccess, tmsAdvGridExcel, ChartLink, TeEngine,
  Series, TeeProcs, Chart,DateUtils;

type
  TForm8 = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    Grid1: TAdvStringGrid;
    SpeedButton1: TSpeedButton;
    Panel3: TPanel;
    Panel4: TPanel;
    SpeedButton2: TSpeedButton;
    QrVentasO: TMSQuery;
    QrGenerales: TMSQuery;
    QrVentasE: TMSQuery;
    QrVentasHD: TMSQuery;
    QrVentasT: TMSQuery;
    QrVentasL: TMSQuery;
    QrVentasL2: TMSQuery;
    QrVentasC: TMSQuery;
    QrVentasI: TMSQuery;
    QrVentasS: TMSQuery;
    QrVentasB: TMSQuery;
    QrVentasPT: TMSQuery;
    QrVentasPL: TMSQuery;
    QrVentasPL2: TMSQuery;
    QrVentasP: TMSQuery;
    SpeedButton3: TSpeedButton;
    Excel1: TAdvGridExcelIO;
    Save1: TSaveDialog;
    QrVentasOT: TMSQuery;
    QrVentasCBHD: TMSQuery;
    QrVentasCub: TMSQuery;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    QrVentasQ: TMSQuery;
    QrVentasQHD: TMSQuery;
    QrImportesCABA: TMSQuery;
    QrImporteEV: TMSQuery;
    QrImporteHD: TMSQuery;
    QrImporteLl: TMSQuery;
    QrImporteTA: TMSQuery;
    QrImporteCO: TMSQuery;
    QrImporteIN: TMSQuery;
    QrImporteSL: TMSQuery;
    QrImporteK: TMSQuery;
    QrImporteLL2: TMSQuery;
    QrImportePT: TMSQuery;
    QrImportePL: TMSQuery;
    QrImportePL2: TMSQuery;
    QrImporteQ: TMSQuery;
    QrImporteQHD: TMSQuery;

    procedure SpeedButton1Click(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
    procedure SpeedButton3Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form8: TForm8;

implementation

{$R *.dfm}

uses unit1,unit6;

procedure TForm8.SpeedButton1Click(Sender: TObject);
var
x,y : integer;
prop,SumaImportes : single;
yearActual,YearLast,mes1n,mes2n,mes3n: integer;
mes1,mes2,mes3,ano1,ano2,ano3 : string;
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
    if form1.qrmetas.FieldByName('MetaUO').AsInteger=null then   grid1.ints[3,y]:=0 else
    grid1.ints[3,y]:=form1.qrmetas.FieldByName('MetaUO').AsInteger;
    if form1.qrmetas.FieldByName('MetaUE').AsInteger=null then   grid1.ints[7,y]:=0 else
    grid1.ints[7,y]:=form1.qrmetas.FieldByName('MetaUE').AsInteger;
    if form1.qrmetas.FieldByName('MetaUEHD').AsInteger=null then   grid1.ints[11,y]:=0 else
    grid1.ints[11,y]:=form1.qrmetas.FieldByName('MetaUEHD').AsInteger;
    if form1.qrmetas.FieldByName('MetaLL').AsInteger=null then   grid1.ints[15,y]:=0 else
    grid1.ints[15,y]:=form1.qrmetas.FieldByName('MetaLL').AsInteger;
    if form1.qrmetas.FieldByName('MetaT').AsInteger=null then   grid1.ints[19,y]:=0 else
    grid1.ints[19,y]:=form1.qrmetas.FieldByName('MetaT').AsInteger;
    if form1.qrmetas.FieldByName('MetaC').AsInteger=null then   grid1.ints[23,y]:=0 else
    grid1.ints[23,y]:=form1.qrmetas.FieldByName('MetaC').AsInteger;
    if form1.qrmetas.FieldByName('MetaI').value=null then   grid1.floats[27,y]:=0 else
    grid1.floats[27,y]:=form1.qrmetas.FieldByName('MetaI').value;
    if form1.qrmetas.FieldByName('MetaS').AsInteger=null then   grid1.ints[31,y]:=0 else
    grid1.ints[31,y]:=form1.qrmetas.FieldByName('MetaS').AsInteger;
    if form1.qrmetas.FieldByName('MetaB').AsInteger=null then   grid1.ints[35,y]:=0 else
    grid1.ints[35,y]:=form1.qrmetas.FieldByName('MetaB').AsInteger;
    if form1.qrmetas.FieldByName('MetaPT').AsInteger=null then   grid1.ints[39,y]:=0 else
    grid1.ints[39,y]:=form1.qrmetas.FieldByName('MetaPT').AsInteger;
    if form1.qrmetas.FieldByName('MetaPL').AsInteger=null then   grid1.ints[43,y]:=0 else
    grid1.ints[43,y]:=form1.qrmetas.FieldByName('MetaPL').AsInteger;
    if form1.qrmetas.FieldByName('MetaQ').AsInteger=null then   grid1.ints[47,y]:=0 else
    grid1.ints[47,y]:=form1.qrmetas.FieldByName('MetaQ').AsInteger;
    if form1.qrmetas.FieldByName('MetaQHD').AsInteger=null then   grid1.ints[51,y]:=0 else
    grid1.ints[51,y]:=form1.qrmetas.FieldByName('MetaQHD').AsInteger;
    if form1.qrmetas.FieldByName('MetaTotal').AsInteger=null then   grid1.ints[55,y]:=0 else
    grid1.ints[55,y]:=form1.qrmetas.FieldByName('MetaTotal').AsInteger;

    if x=1 then begin grid1.Cells[1,y]:=mes1+' '+ano1; grid1.Ints[58,y]:=mes1n; grid1.Cells[59,y]:=ano1 end;
    if x=2 then begin grid1.Cells[1,y]:=mes2+' '+ano2; grid1.Ints[58,y]:=mes2n; grid1.Cells[59,y]:=ano2 end;
    if x=3 then begin grid1.Cells[1,y]:=mes3+' '+ano3; grid1.Ints[58,y]:=mes3n; grid1.Cells[59,y]:=ano3 end;

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
       qrGenerales.Close;
       qrGenerales.ParamByName('nombre').AsString:=grid1.Cells[0,y];
       qrGenerales.Open;

       qrventasO.close;
       QrVentasO.ParamByName('DevVenta').AsString:=QrGenerales.FieldByName('Devolucion').AsString;
       QrVentasO.ParamByName('Tienda').AsString:=QrGenerales.FieldByName('Factura').AsString;
       QrVentasO.ParamByName('Fecha1').AsDate:=encodedate(grid1.Ints[59,y],grid1.Ints[58,y],1);
       QrVentasO.ParamByName('Fecha2').AsDate:=encodedate(grid1.Ints[59,y],grid1.Ints[58,y],daysinamonth(grid1.Ints[59,y],grid1.Ints[58,y]));
       QrVentasO.ParamByName('Usuario').AsString:=QrGenerales.FieldByName('Usuario').AsString;
       QrVentasO.Open;

       qrventasE.close;
       QrVentasE.ParamByName('DevVenta').AsString:=QrGenerales.FieldByName('Devolucion').AsString;
       QrVentasE.ParamByName('Tienda').AsString:=QrGenerales.FieldByName('Factura').AsString;
       QrVentasE.ParamByName('Fecha1').AsDate:=encodedate(grid1.Ints[59,y],grid1.Ints[58,y],1);
       QrVentasE.ParamByName('Fecha2').AsDate:=encodedate(grid1.Ints[59,y],grid1.Ints[58,y],daysinamonth(grid1.Ints[59,y],grid1.Ints[58,y]));
       QrVentasE.ParamByName('Usuario').AsString:=QrGenerales.FieldByName('Usuario').AsString;
       QrVentasE.Open;

       qrventasHD.close;
       QrVentasHD.ParamByName('DevVenta').AsString:=QrGenerales.FieldByName('Devolucion').AsString;
       QrVentasHD.ParamByName('Tienda').AsString:=QrGenerales.FieldByName('Factura').AsString;
       QrVentasHD.ParamByName('Fecha1').AsDate:=encodedate(grid1.Ints[59,y],grid1.Ints[58,y],1);
       QrVentasHD.ParamByName('Fecha2').AsDate:=encodedate(grid1.Ints[59,y],grid1.Ints[58,y],daysinamonth(grid1.Ints[59,y],grid1.Ints[58,y]));
       QrVentasHD.ParamByName('Usuario').AsString:=QrGenerales.FieldByName('Usuario').AsString;
       QrVentasHD.Open;

       qrventasT.close;
       QrVentasT.ParamByName('DevVenta').AsString:=QrGenerales.FieldByName('Devolucion').AsString;
       QrVentasT.ParamByName('Tienda').AsString:=QrGenerales.FieldByName('Factura').AsString;
       QrVentasT.ParamByName('Fecha1').AsDate:=encodedate(grid1.Ints[59,y],grid1.Ints[58,y],1);
       QrVentasT.ParamByName('Fecha2').AsDate:=encodedate(grid1.Ints[59,y],grid1.Ints[58,y],daysinamonth(grid1.Ints[59,y],grid1.Ints[58,y]));
       QrVentasT.ParamByName('Usuario').AsString:=QrGenerales.FieldByName('Usuario').AsString;
       QrVentasT.Open;

       qrventasL.close;
       QrVentasL.ParamByName('DevVenta').AsString:=QrGenerales.FieldByName('Devolucion').AsString;
       QrVentasL.ParamByName('Tienda').AsString:=QrGenerales.FieldByName('Factura').AsString;
       QrVentasL.ParamByName('Fecha1').AsDate:=encodedate(grid1.Ints[59,y],grid1.Ints[58,y],1);
       QrVentasL.ParamByName('Fecha2').AsDate:=encodedate(grid1.Ints[59,y],grid1.Ints[58,y],daysinamonth(grid1.Ints[59,y],grid1.Ints[58,y]));
       QrVentasL.ParamByName('Usuario').AsString:=QrGenerales.FieldByName('Usuario').AsString;
       QrVentasL.Open;

       qrventasL2.close;
       QrVentasL2.ParamByName('DevVenta').AsString:=QrGenerales.FieldByName('Devolucion').AsString;
       QrVentasL2.ParamByName('Tienda').AsString:=QrGenerales.FieldByName('Factura').AsString;
       QrVentasL2.ParamByName('Fecha1').AsDate:=encodedate(grid1.Ints[59,y],grid1.Ints[58,y],1);
       QrVentasL2.ParamByName('Fecha2').AsDate:=encodedate(grid1.Ints[59,y],grid1.Ints[58,y],daysinamonth(grid1.Ints[59,y],grid1.Ints[58,y]));
       QrVentasL2.ParamByName('Usuario').AsString:=QrGenerales.FieldByName('Usuario').AsString;
       QrVentasL2.Open;

       qrventasC.close;
       QrVentasC.ParamByName('DevVenta').AsString:=QrGenerales.FieldByName('Devolucion').AsString;
       QrVentasC.ParamByName('Tienda').AsString:=QrGenerales.FieldByName('Factura').AsString;
       QrVentasC.ParamByName('Fecha1').AsDate:=encodedate(grid1.Ints[59,y],grid1.Ints[58,y],1);
       QrVentasC.ParamByName('Fecha2').AsDate:=encodedate(grid1.Ints[59,y],grid1.Ints[58,y],daysinamonth(grid1.Ints[59,y],grid1.Ints[58,y]));
       QrVentasC.ParamByName('Usuario').AsString:=QrGenerales.FieldByName('Usuario').AsString;
       QrVentasC.Open;

       qrventasI.close;
       QrVentasI.ParamByName('DevVenta').AsString:=QrGenerales.FieldByName('Devolucion').AsString;
       QrVentasI.ParamByName('Tienda').AsString:=QrGenerales.FieldByName('Factura').AsString;
       QrVentasI.ParamByName('Fecha1').AsDate:=encodedate(grid1.Ints[59,y],grid1.Ints[58,y],1);
       QrVentasI.ParamByName('Fecha2').AsDate:=encodedate(grid1.Ints[59,y],grid1.Ints[58,y],daysinamonth(grid1.Ints[59,y],grid1.Ints[58,y]));
       QrVentasI.ParamByName('Usuario').AsString:=QrGenerales.FieldByName('Usuario').AsString;
       QrVentasI.Open;

       qrventasS.close;
       QrVentasS.ParamByName('DevVenta').AsString:=QrGenerales.FieldByName('Devolucion').AsString;
       QrVentasS.ParamByName('Tienda').AsString:=QrGenerales.FieldByName('Factura').AsString;
       QrVentasS.ParamByName('Fecha1').AsDate:=encodedate(grid1.Ints[59,y],grid1.Ints[58,y],1);
       QrVentasS.ParamByName('Fecha2').AsDate:=encodedate(grid1.Ints[59,y],grid1.Ints[58,y],daysinamonth(grid1.Ints[59,y],grid1.Ints[58,y]));
       QrVentasS.ParamByName('Usuario').AsString:=QrGenerales.FieldByName('Usuario').AsString;
       QrVentasS.Open;

       qrventasB.close;
       QrVentasB.ParamByName('DevVenta').AsString:=QrGenerales.FieldByName('Devolucion').AsString;
       QrVentasB.ParamByName('Tienda').AsString:=QrGenerales.FieldByName('Factura').AsString;
       QrVentasB.ParamByName('Fecha1').AsDate:=encodedate(grid1.Ints[59,y],grid1.Ints[58,y],1);
       QrVentasB.ParamByName('Fecha2').AsDate:=encodedate(grid1.Ints[59,y],grid1.Ints[58,y],daysinamonth(grid1.Ints[59,y],grid1.Ints[58,y]));
       QrVentasB.ParamByName('Usuario').AsString:=QrGenerales.FieldByName('Usuario').AsString;
       QrVentasB.Open;

       qrventasPT.close;
       QrVentasPT.ParamByName('DevVenta').AsString:=QrGenerales.FieldByName('Devolucion').AsString;
       QrVentasPT.ParamByName('Tienda').AsString:=QrGenerales.FieldByName('Factura').AsString;
       QrVentasPT.ParamByName('Fecha1').AsDate:=encodedate(grid1.Ints[59,y],grid1.Ints[58,y],1);
       QrVentasPT.ParamByName('Fecha2').AsDate:=encodedate(grid1.Ints[59,y],grid1.Ints[58,y],daysinamonth(grid1.Ints[59,y],grid1.Ints[58,y]));
       QrVentasPT.ParamByName('Usuario').AsString:=QrGenerales.FieldByName('Usuario').AsString;
       QrVentasPT.Open;

       qrventasPL.close;
       QrVentasPL.ParamByName('DevVenta').AsString:=QrGenerales.FieldByName('Devolucion').AsString;
       QrVentasPL.ParamByName('Tienda').AsString:=QrGenerales.FieldByName('Factura').AsString;
       QrVentasPL.ParamByName('Fecha1').AsDate:=encodedate(grid1.Ints[59,y],grid1.Ints[58,y],1);
       QrVentasPL.ParamByName('Fecha2').AsDate:=encodedate(grid1.Ints[59,y],grid1.Ints[58,y],daysinamonth(grid1.Ints[59,y],grid1.Ints[58,y]));
       QrVentasPL.ParamByName('Usuario').AsString:=QrGenerales.FieldByName('Usuario').AsString;
       QrVentasPL.Open;

       qrventasPL2.close;
       QrVentasPL2.ParamByName('DevVenta').AsString:=QrGenerales.FieldByName('Devolucion').AsString;
       QrVentasPL2.ParamByName('Tienda').AsString:=QrGenerales.FieldByName('Factura').AsString;
       QrVentasPL2.ParamByName('Fecha1').AsDate:=encodedate(grid1.Ints[59,y],grid1.Ints[58,y],1);
       QrVentasPL2.ParamByName('Fecha2').AsDate:=encodedate(grid1.Ints[59,y],grid1.Ints[58,y],daysinamonth(grid1.Ints[59,y],grid1.Ints[58,y]));
       QrVentasPL2.ParamByName('Usuario').AsString:=QrGenerales.FieldByName('Usuario').AsString;
       QrVentasPL2.Open;

       qrventasQ.close;
       QrVentasQ.ParamByName('DevVenta').AsString:=QrGenerales.FieldByName('Devolucion').AsString;
       QrVentasQ.ParamByName('Tienda').AsString:=QrGenerales.FieldByName('Factura').AsString;
       QrVentasQ.ParamByName('Fecha1').AsDate:=encodedate(grid1.Ints[59,y],grid1.Ints[58,y],1);
       QrVentasQ.ParamByName('Fecha2').AsDate:=encodedate(grid1.Ints[59,y],grid1.Ints[58,y],daysinamonth(grid1.Ints[59,y],grid1.Ints[58,y]));
       QrVentasQ.ParamByName('Usuario').AsString:=QrGenerales.FieldByName('Usuario').AsString;
       QrVentasQ.Open;

       qrventasQHD.close;
       QrVentasQHD.ParamByName('DevVenta').AsString:=QrGenerales.FieldByName('Devolucion').AsString;
       QrVentasQHD.ParamByName('Tienda').AsString:=QrGenerales.FieldByName('Factura').AsString;
       QrVentasQHD.ParamByName('Fecha1').AsDate:=encodedate(grid1.Ints[59,y],grid1.Ints[58,y],1);
       QrVentasQHD.ParamByName('Fecha2').AsDate:=encodedate(grid1.Ints[59,y],grid1.Ints[58,y],daysinamonth(grid1.Ints[59,y],grid1.Ints[58,y]));
       QrVentasQHD.ParamByName('Usuario').AsString:=QrGenerales.FieldByName('Usuario').AsString;
       QrVentasQHD.Open;



       if qrVentasO.isempty=true then grid1.ints[2,y]:=0 else
       grid1.ints[2,y]:=qrVentasO.FieldByName('Cantidad').AsInteger;

       if qrVentasE.isempty=true then grid1.ints[6,y]:=0 else
       grid1.ints[6,y]:=qrVentasE.FieldByName('Cantidad').AsInteger;

       if qrVentasHD.isempty=true then grid1.ints[10,y]:=0 else
       grid1.ints[10,y]:=qrVentasHD.FieldByName('Cantidad').AsInteger;

       if qrVentasL.isempty=true then grid1.ints[14,y]:=0 else
       grid1.ints[14,y]:=qrVentasL.FieldByName('Cantidad').AsInteger+qrVentasL2.FieldByName('Cantidad').AsInteger;

       if qrVentasT.isempty=true then grid1.ints[18,y]:=0 else
       grid1.ints[18,y]:=qrVentasT.FieldByName('Cantidad').AsInteger;

       if qrVentasC.isempty=true then grid1.ints[22,y]:=0 else
       grid1.ints[22,y]:=qrVentasC.FieldByName('Cantidad').AsInteger;

       if qrVentasI.isempty=true then grid1.ints[26,y]:=0 else
       grid1.floats[26,y]:=qrVentasI.FieldByName('Cantidad').value;

       if qrVentasS.isempty=true then grid1.ints[30,y]:=0 else
       grid1.ints[30,y]:=qrVentasS.FieldByName('Cantidad').asinteger;

       if qrVentasB.isempty=true then grid1.ints[34,y]:=0 else
       grid1.ints[34,y]:=qrVentasB.FieldByName('Cantidad').asinteger;

       if qrVentasPT.isempty=true then grid1.ints[38,y]:=0 else
       grid1.ints[38,y]:=qrVentasPT.FieldByName('Cantidad').asinteger;

       if qrVentasPL.isempty=true then grid1.ints[42,y]:=0 else
       grid1.ints[42,y]:=qrVentasPL.FieldByName('Cantidad').AsInteger+qrVentasPL2.FieldByName('Cantidad').AsInteger;

       if qrVentasQ.isempty=true then grid1.ints[46,y]:=0 else
       grid1.ints[46,y]:=qrVentasQ.FieldByName('Cantidad').asinteger;

       if qrVentasQHD.isempty=true then grid1.ints[50,y]:=0 else
       grid1.ints[50,y]:=qrVentasQHD.FieldByName('Cantidad').asinteger;


       //suma de piezas totales
       grid1.Ints[54,y]:=grid1.Ints[2,y]+grid1.Ints[6,y]+grid1.Ints[14,y]+grid1.Ints[18,y]+grid1.Ints[22,y]+grid1.Ints[30,y]+grid1.Ints[34,y]+grid1.Ints[38,y]+grid1.Ints[42,y]+grid1.Ints[46,y]+grid1.Ints[50,y];
       //comienzan las operaciones   -siempre es meta - venta --cambio el 9 de junio a venta - meta
       grid1.Ints[4,y]:=grid1.Ints[2,y]-grid1.Ints[3,y];
       grid1.Ints[8,y]:=grid1.Ints[6,y]-grid1.Ints[7,y];
       grid1.Ints[12,y]:=grid1.Ints[10,y]-grid1.Ints[11,y];
       grid1.Ints[16,y]:=grid1.Ints[14,y]-grid1.Ints[15,y];
       grid1.Ints[20,y]:=grid1.Ints[18,y]-grid1.Ints[19,y];
       grid1.Ints[24,y]:=grid1.Ints[22,y]-grid1.Ints[23,y];
       grid1.Ints[28,y]:=grid1.Ints[26,y]-grid1.Ints[27,y];
       grid1.Ints[32,y]:=grid1.Ints[30,y]-grid1.Ints[31,y];
       grid1.Ints[36,y]:=grid1.Ints[34,y]-grid1.Ints[35,y];
       grid1.Ints[40,y]:=grid1.Ints[38,y]-grid1.Ints[39,y];
       grid1.Ints[44,y]:=grid1.Ints[42,y]-grid1.Ints[43,y];
       grid1.Ints[48,y]:=grid1.Ints[46,y]-grid1.Ints[47,y];
       grid1.Ints[52,y]:=grid1.Ints[50,y]-grid1.Ints[51,y];
       grid1.Ints[56,y]:=grid1.Ints[54,y]-grid1.Ints[55,y];
       //porcentajes
       if grid1.Floats[3,y]=0 then  grid1.Floats[5,y]:=0 else
       grid1.Floats[5,y]:=(grid1.Floats[2,y]*100)/grid1.Floats[3,y];

       if grid1.Floats[7,y]=0 then  grid1.Floats[9,y]:=0 else
       grid1.Floats[9,y]:=(grid1.Floats[6,y]*100)/grid1.Floats[7,y];

       if grid1.Floats[11,y]=0 then  grid1.Floats[13,y]:=0 else
       grid1.Floats[13,y]:=(grid1.Floats[10,y]*100)/grid1.Floats[11,y];

       if grid1.Floats[15,y]=0 then  grid1.Floats[17,y]:=0 else
       grid1.Floats[17,y]:=(grid1.Floats[14,y]*100)/grid1.Floats[15,y];

       if grid1.Floats[19,y]=0 then  grid1.Floats[21,y]:=0 else
       grid1.Floats[21,y]:=(grid1.Floats[18,y]*100)/grid1.Floats[19,y];

       if grid1.Floats[23,y]=0 then  grid1.Floats[25,y]:=0 else
       grid1.Floats[25,y]:=(grid1.Floats[22,y]*100)/grid1.Floats[23,y];

       if grid1.Floats[27,y]=0 then  grid1.Floats[29,y]:=0 else
       grid1.Floats[29,y]:=(grid1.Floats[26,y]*100)/grid1.Floats[27,y];

       if grid1.Floats[31,y]=0 then  grid1.Floats[33,y]:=0 else
       grid1.Floats[33,y]:=(grid1.Floats[30,y]*100)/grid1.Floats[31,y];

       if grid1.Floats[35,y]=0 then  grid1.Floats[37,y]:=0 else
       grid1.Floats[37,y]:=(grid1.Floats[34,y]*100)/grid1.Floats[35,y];

       if grid1.Floats[39,y]=0 then  grid1.Floats[41,y]:=0 else
       grid1.Floats[41,y]:=(grid1.Floats[38,y]*100)/grid1.Floats[39,y];

       if grid1.Floats[43,y]=0 then  grid1.Floats[45,y]:=0 else
       grid1.Floats[45,y]:=(grid1.Floats[42,y]*100)/grid1.Floats[43,y];

       if grid1.Floats[47,y]=0 then  grid1.Floats[49,y]:=0 else
       grid1.Floats[49,y]:=(grid1.Floats[46,y]*100)/grid1.Floats[47,y];

       if grid1.Floats[51,y]=0 then  grid1.Floats[53,y]:=0 else
       grid1.Floats[53,y]:=(grid1.Floats[50,y]*100)/grid1.Floats[51,y];

       if grid1.Floats[55,y]=0 then  grid1.Floats[57,y]:=0 else
       grid1.Floats[57,y]:=(grid1.Floats[54,y]*100)/grid1.Floats[55,y];

     end;




     //formateo de %
 for y := 1 to grid1.RowCount-1 do
   begin
     grid1.Cells[55,y]:=FormatFloat('#,##0',grid1.Floats[55,y]);

     grid1.Cells[5,y]:=FormatFloat('##0%',grid1.Floats[5,y]);
     grid1.Cells[9,y]:=FormatFloat('##0%',grid1.Floats[9,y]);
     grid1.Cells[13,y]:=FormatFloat('##0%',grid1.Floats[13,y]);
     grid1.Cells[17,y]:=FormatFloat('##0%',grid1.Floats[17,y]);
     grid1.Cells[21,y]:=FormatFloat('##0%',grid1.Floats[21,y]);
     grid1.Cells[25,y]:=FormatFloat('##0%',grid1.Floats[25,y]);
     grid1.Cells[29,y]:=FormatFloat('##0%',grid1.Floats[29,y]);
     grid1.Cells[33,y]:=FormatFloat('##0%',grid1.Floats[33,y]);
     grid1.Cells[37,y]:=FormatFloat('##0%',grid1.Floats[37,y]);
     grid1.Cells[41,y]:=FormatFloat('##0%',grid1.Floats[41,y]);
     grid1.Cells[45,y]:=FormatFloat('##0%',grid1.Floats[45,y]);
     grid1.Cells[49,y]:=FormatFloat('##0%',grid1.Floats[49,y]);
     grid1.Cells[53,y]:=FormatFloat('##0%',grid1.Floats[53,y]);

     grid1.Cells[57,y]:=FormatFloat('##0%',grid1.Floats[57,y]);


   end;

     //aqui va venta en pesos

     for y := 1 to grid1.RowCount-1 do
     begin
       SumaImportes:=0;
       qrGenerales.Close;
       qrGenerales.ParamByName('nombre').AsString:=grid1.Cells[0,y];
       qrGenerales.Open;

       QrImportesCABA.close;
       QrImportesCABA.ParamByName('DevVenta').AsString:=QrGenerales.FieldByName('Devolucion').AsString;
       QrImportesCABA.ParamByName('Tienda').AsString:=QrGenerales.FieldByName('Factura').AsString;
       QrImportesCABA.ParamByName('Fecha1').AsDate:=encodedate(grid1.Ints[59,y],grid1.Ints[58,y],1);
       QrImportesCABA.ParamByName('Fecha2').AsDate:=encodedate(grid1.Ints[59,y],grid1.Ints[58,y],daysinamonth(grid1.Ints[59,y],grid1.Ints[58,y]));
       QrImportesCABA.ParamByName('Usuario').AsString:=QrGenerales.FieldByName('Usuario').AsString;
       QrImportesCABA.Open;

       QrImporteEV.close;
       QrImporteEV.ParamByName('DevVenta').AsString:=QrGenerales.FieldByName('Devolucion').AsString;
       QrImporteEV.ParamByName('Tienda').AsString:=QrGenerales.FieldByName('Factura').AsString;
       QrImporteEV.ParamByName('Fecha1').AsDate:=encodedate(grid1.Ints[59,y],grid1.Ints[58,y],1);
       QrImporteEV.ParamByName('Fecha2').AsDate:=encodedate(grid1.Ints[59,y],grid1.Ints[58,y],daysinamonth(grid1.Ints[59,y],grid1.Ints[58,y]));
       QrImporteEV.ParamByName('Usuario').AsString:=QrGenerales.FieldByName('Usuario').AsString;
       QrImporteEV.Open;

       QrImporteHD.close;
       QrImporteHD.ParamByName('DevVenta').AsString:=QrGenerales.FieldByName('Devolucion').AsString;
       QrImporteHD.ParamByName('Tienda').AsString:=QrGenerales.FieldByName('Factura').AsString;
       QrImporteHD.ParamByName('Fecha1').AsDate:=encodedate(grid1.Ints[59,y],grid1.Ints[58,y],1);
       QrImporteHD.ParamByName('Fecha2').AsDate:=encodedate(grid1.Ints[59,y],grid1.Ints[58,y],daysinamonth(grid1.Ints[59,y],grid1.Ints[58,y]));
       QrImporteHD.ParamByName('Usuario').AsString:=QrGenerales.FieldByName('Usuario').AsString;
       QrImporteHD.Open;

       QrImporteLL.close;
       QrImporteLL.ParamByName('DevVenta').AsString:=QrGenerales.FieldByName('Devolucion').AsString;
       QrImporteLL.ParamByName('Tienda').AsString:=QrGenerales.FieldByName('Factura').AsString;
       QrImporteLL.ParamByName('Fecha1').AsDate:=encodedate(grid1.Ints[59,y],grid1.Ints[58,y],1);
       QrImporteLL.ParamByName('Fecha2').AsDate:=encodedate(grid1.Ints[59,y],grid1.Ints[58,y],daysinamonth(grid1.Ints[59,y],grid1.Ints[58,y]));
       QrImporteLL.ParamByName('Usuario').AsString:=QrGenerales.FieldByName('Usuario').AsString;
       QrImporteLL.Open;

       QrImporteLL2.close;
       QrImporteLL2.ParamByName('DevVenta').AsString:=QrGenerales.FieldByName('Devolucion').AsString;
       QrImporteLL2.ParamByName('Tienda').AsString:=QrGenerales.FieldByName('Factura').AsString;
       QrImporteLL2.ParamByName('Fecha1').AsDate:=encodedate(grid1.Ints[59,y],grid1.Ints[58,y],1);
       QrImporteLL2.ParamByName('Fecha2').AsDate:=encodedate(grid1.Ints[59,y],grid1.Ints[58,y],daysinamonth(grid1.Ints[59,y],grid1.Ints[58,y]));
       QrImporteLL2.ParamByName('Usuario').AsString:=QrGenerales.FieldByName('Usuario').AsString;
       QrImporteLL2.Open;

       QrImporteTA.close;
       QrImporteTA.ParamByName('DevVenta').AsString:=QrGenerales.FieldByName('Devolucion').AsString;
       QrImporteTA.ParamByName('Tienda').AsString:=QrGenerales.FieldByName('Factura').AsString;
       QrImporteTA.ParamByName('Fecha1').AsDate:=encodedate(grid1.Ints[59,y],grid1.Ints[58,y],1);
       QrImporteTA.ParamByName('Fecha2').AsDate:=encodedate(grid1.Ints[59,y],grid1.Ints[58,y],daysinamonth(grid1.Ints[59,y],grid1.Ints[58,y]));
       QrImporteTA.ParamByName('Usuario').AsString:=QrGenerales.FieldByName('Usuario').AsString;
       QrImporteTA.Open;

       QrImporteCO.close;
       QrImporteCO.ParamByName('DevVenta').AsString:=QrGenerales.FieldByName('Devolucion').AsString;
       QrImporteCO.ParamByName('Tienda').AsString:=QrGenerales.FieldByName('Factura').AsString;
       QrImporteCO.ParamByName('Fecha1').AsDate:=encodedate(grid1.Ints[59,y],grid1.Ints[58,y],1);
       QrImporteCO.ParamByName('Fecha2').AsDate:=encodedate(grid1.Ints[59,y],grid1.Ints[58,y],daysinamonth(grid1.Ints[59,y],grid1.Ints[58,y]));
       QrImporteCO.ParamByName('Usuario').AsString:=QrGenerales.FieldByName('Usuario').AsString;
       QrImporteCO.Open;

       QrImporteIN.close;
       QrImporteIN.ParamByName('DevVenta').AsString:=QrGenerales.FieldByName('Devolucion').AsString;
       QrImporteIN.ParamByName('Tienda').AsString:=QrGenerales.FieldByName('Factura').AsString;
       QrImporteIN.ParamByName('Fecha1').AsDate:=encodedate(grid1.Ints[59,y],grid1.Ints[58,y],1);
       QrImporteIN.ParamByName('Fecha2').AsDate:=encodedate(grid1.Ints[59,y],grid1.Ints[58,y],daysinamonth(grid1.Ints[59,y],grid1.Ints[58,y]));
       QrImporteIN.ParamByName('Usuario').AsString:=QrGenerales.FieldByName('Usuario').AsString;
       QrImporteIN.Open;

       QrImporteSL.close;
       QrImporteSL.ParamByName('DevVenta').AsString:=QrGenerales.FieldByName('Devolucion').AsString;
       QrImporteSL.ParamByName('Tienda').AsString:=QrGenerales.FieldByName('Factura').AsString;
       QrImporteSL.ParamByName('Fecha1').AsDate:=encodedate(grid1.Ints[59,y],grid1.Ints[58,y],1);
       QrImporteSL.ParamByName('Fecha2').AsDate:=encodedate(grid1.Ints[59,y],grid1.Ints[58,y],daysinamonth(grid1.Ints[59,y],grid1.Ints[58,y]));
       QrImporteSL.ParamByName('Usuario').AsString:=QrGenerales.FieldByName('Usuario').AsString;
       QrImporteSL.Open;

       QrImporteK.close;
       QrImporteK.ParamByName('DevVenta').AsString:=QrGenerales.FieldByName('Devolucion').AsString;
       QrImporteK.ParamByName('Tienda').AsString:=QrGenerales.FieldByName('Factura').AsString;
       QrImporteK.ParamByName('Fecha1').AsDate:=encodedate(grid1.Ints[59,y],grid1.Ints[58,y],1);
       QrImporteK.ParamByName('Fecha2').AsDate:=encodedate(grid1.Ints[59,y],grid1.Ints[58,y],daysinamonth(grid1.Ints[59,y],grid1.Ints[58,y]));
       QrImporteK.ParamByName('Usuario').AsString:=QrGenerales.FieldByName('Usuario').AsString;
       QrImporteK.Open;

       QrImportePT.close;
       QrImportePT.ParamByName('DevVenta').AsString:=QrGenerales.FieldByName('Devolucion').AsString;
       QrImportePT.ParamByName('Tienda').AsString:=QrGenerales.FieldByName('Factura').AsString;
       QrImportePT.ParamByName('Fecha1').AsDate:=encodedate(grid1.Ints[59,y],grid1.Ints[58,y],1);
       QrImportePT.ParamByName('Fecha2').AsDate:=encodedate(grid1.Ints[59,y],grid1.Ints[58,y],daysinamonth(grid1.Ints[59,y],grid1.Ints[58,y]));
       QrImportePT.ParamByName('Usuario').AsString:=QrGenerales.FieldByName('Usuario').AsString;
       QrImportePT.Open;

       QrImportePL.close;
       QrImportePL.ParamByName('DevVenta').AsString:=QrGenerales.FieldByName('Devolucion').AsString;
       QrImportePL.ParamByName('Tienda').AsString:=QrGenerales.FieldByName('Factura').AsString;
       QrImportePL.ParamByName('Fecha1').AsDate:=encodedate(grid1.Ints[59,y],grid1.Ints[58,y],1);
       QrImportePL.ParamByName('Fecha2').AsDate:=encodedate(grid1.Ints[59,y],grid1.Ints[58,y],daysinamonth(grid1.Ints[59,y],grid1.Ints[58,y]));
       QrImportePL.ParamByName('Usuario').AsString:=QrGenerales.FieldByName('Usuario').AsString;
       QrImportePL.Open;

       QrImportePL2.close;
       QrImportePL2.ParamByName('DevVenta').AsString:=QrGenerales.FieldByName('Devolucion').AsString;
       QrImportePL2.ParamByName('Tienda').AsString:=QrGenerales.FieldByName('Factura').AsString;
       QrImportePL2.ParamByName('Fecha1').AsDate:=encodedate(grid1.Ints[59,y],grid1.Ints[58,y],1);
       QrImportePL2.ParamByName('Fecha2').AsDate:=encodedate(grid1.Ints[59,y],grid1.Ints[58,y],daysinamonth(grid1.Ints[59,y],grid1.Ints[58,y]));
       QrImportePL2.ParamByName('Usuario').AsString:=QrGenerales.FieldByName('Usuario').AsString;
       QrImportePL2.Open;

       QrImporteQ.close;
       QrImporteQ.ParamByName('DevVenta').AsString:=QrGenerales.FieldByName('Devolucion').AsString;
       QrImporteQ.ParamByName('Tienda').AsString:=QrGenerales.FieldByName('Factura').AsString;
       QrImporteQ.ParamByName('Fecha1').AsDate:=encodedate(grid1.Ints[59,y],grid1.Ints[58,y],1);
       QrImporteQ.ParamByName('Fecha2').AsDate:=encodedate(grid1.Ints[59,y],grid1.Ints[58,y],daysinamonth(grid1.Ints[59,y],grid1.Ints[58,y]));
       QrImporteQ.ParamByName('Usuario').AsString:=QrGenerales.FieldByName('Usuario').AsString;
       QrImporteQ.Open;

       QrImporteQHD.close;
       QrImporteQHD.ParamByName('DevVenta').AsString:=QrGenerales.FieldByName('Devolucion').AsString;
       QrImporteQHD.ParamByName('Tienda').AsString:=QrGenerales.FieldByName('Factura').AsString;
       QrImporteQHD.ParamByName('Fecha1').AsDate:=encodedate(grid1.Ints[59,y],grid1.Ints[58,y],1);
       QrImporteQHD.ParamByName('Fecha2').AsDate:=encodedate(grid1.Ints[59,y],grid1.Ints[58,y],daysinamonth(grid1.Ints[59,y],grid1.Ints[58,y]));
       QrImporteQHD.ParamByName('Usuario').AsString:=QrGenerales.FieldByName('Usuario').AsString;
       QrImporteQHD.Open;

       if QrImportesCABA.isempty=true then grid1.floats[60,y]:=0 else
       grid1.floats[60,y]:=QrImportesCABA.FieldByName('Importe').value;
       SumaImportes:=SumaImportes+grid1.floats[60,y];

       if QrImporteEV.isempty=true then grid1.floats[61,y]:=0 else
       grid1.floats[61,y]:=QrImporteEV.FieldByName('Importe').value;
       SumaImportes:=SumaImportes+grid1.floats[61,y];

       if QrImporteHD.isempty=true then grid1.floats[62,y]:=0 else
       grid1.floats[62,y]:=QrImporteHD.FieldByName('Importe').value;
       SumaImportes:=SumaImportes+grid1.floats[62,y];

       if QrImporteLL.isempty=true then grid1.floats[63,y]:=0 else
       grid1.floats[63,y]:=QrImporteLL.FieldByName('Importe').value;
       SumaImportes:=SumaImportes+grid1.floats[63,y];

       if QrImporteLL2.isempty=true then grid1.floats[63,y]:=grid1.floats[63,y]+0 else
       grid1.floats[63,y]:=grid1.floats[63,y]+QrImporteLL.FieldByName('Importe').value;
       SumaImportes:=SumaImportes+grid1.floats[64,y];

       if QrImporteTA.isempty=true then grid1.floats[64,y]:=0 else
       grid1.floats[64,y]:=QrImporteTA.FieldByName('Importe').value;
       SumaImportes:=SumaImportes+grid1.floats[65,y];

       if QrImporteCO.isempty=true then grid1.floats[65,y]:=0 else
       grid1.floats[65,y]:=QrImporteCO.FieldByName('Importe').value;
       SumaImportes:=SumaImportes+grid1.floats[65,y];

       if QrImporteIN.isempty=true then grid1.floats[66,y]:=0 else
       grid1.floats[66,y]:=QrImporteIN.FieldByName('Importe').value;
       SumaImportes:=SumaImportes+grid1.floats[66,y];

       if QrImporteSL.isempty=true then grid1.floats[67,y]:=0 else
       grid1.floats[67,y]:=QrImporteSL.FieldByName('Importe').value;
       SumaImportes:=SumaImportes+grid1.floats[67,y];

       if QrImporteK.isempty=true then grid1.floats[68,y]:=0 else
       grid1.floats[68,y]:=QrImporteK.FieldByName('Importe').value;
       SumaImportes:=SumaImportes+grid1.floats[68,y];

       if QrImportePT.isempty=true then grid1.floats[69,y]:=0 else
       grid1.floats[69,y]:=QrImportePT.FieldByName('Importe').value;
       SumaImportes:=SumaImportes+grid1.floats[69,y];

       if QrImportePL.isempty=true then grid1.floats[70,y]:=0 else
       grid1.floats[70,y]:=QrImportePL.FieldByName('Importe').value;


       if QrImportePL2.isempty=true then grid1.floats[70,y]:=grid1.floats[70,y]+0 else
       grid1.floats[70,y]:=grid1.floats[70,y]+QrImportePL2.FieldByName('Importe').value;
       SumaImportes:=SumaImportes+grid1.floats[70,y];

       if QrImporteQ.isempty=true then grid1.floats[71,y]:=0 else
       grid1.floats[71,y]:=QrImporteQ.FieldByName('Importe').value;
       SumaImportes:=SumaImportes+grid1.floats[71,y];

       if QrImporteQHD.isempty=true then grid1.floats[72,y]:=0 else
       grid1.floats[72,y]:=QrImporteQHD.FieldByName('Importe').value;
       SumaImportes:=SumaImportes+grid1.floats[72,y];

       Grid1.Floats[73,y]:=SumaImportes;


     end;

     grid1.MoveColumn(60,6); //CABA
     grid1.MoveColumn(61,11); //EVOKE
     grid1.MoveColumn(62,16); //HD
     grid1.MoveColumn(63,21); //Ll
     grid1.MoveColumn(64,26); //Tarjas
     grid1.MoveColumn(65,31); //Complementos
     grid1.MoveColumn(66,36); //Innovika
     grid1.MoveColumn(67,41); //Slim
     grid1.MoveColumn(68,46); //Basi-K
     grid1.MoveColumn(69,51); //Plados Tarjas
     grid1.MoveColumn(70,56); //Plados LLaves
     grid1.MoveColumn(71,61); //Q
     grid1.MoveColumn(72,66); //Q HD


     grid1.HideColumns(71,72);
     Grid1.SortIndexes.Clear;
     Grid1.SortIndexes.Add(0);
     Grid1.SortIndexes.Add(71);
     Grid1.SortIndexes.Add(72);
     Grid1.QSortIndexed;




end;

procedure TForm8.SpeedButton2Click(Sender: TObject);
begin
Form8.Visible:=false;
form1.Visible:=true;
end;

procedure TForm8.SpeedButton3Click(Sender: TObject);
var
x,y,j,k: integer;
begin
{j:=0;
k:=1;
gridex.ClearAll;
gridex.RowCount:=1;
Gridex.RowCount:=gridex.RowCount+1;
for y := 0 to gridg.RowCount-1 do
 begin
   for x := 0 to 10 do
     begin
       gridex.Cells[j,k]:=gridg.Cells[x,y];
       inc(j);
     end;
    inc(k);
    gridex.rowcount:=gridex.RowCount+1;
    j:=0;
 end;}



save1.Execute;
Excel1.XLSExport(save1.FileName,'Datos');


showmessage('Reporte Exportado con Exito!');
end;

end.
