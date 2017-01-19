unit Unit2;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, Grids, AdvObj, BaseGrid, AdvGrid, Buttons, StdCtrls,
  ComCtrls, DB, MemDS, DBAccess, MSAccess, tmsAdvGridExcel, ChartLink, TeEngine,
  Series, TeeProcs, Chart;

type
  TForm2 = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    Grid1: TAdvStringGrid;
    Fecha1: TDateTimePicker;
    Fecha2: TDateTimePicker;
    Label1: TLabel;
    Label2: TLabel;
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
    TabSheet2: TTabSheet;
    GridG: TAdvStringGrid;
    GridE: TAdvStringGrid;
    GridH: TAdvStringGrid;
    GridL: TAdvStringGrid;
    ScrollBox1: TScrollBox;
    GridT: TAdvStringGrid;
    GridC: TAdvStringGrid;
    GridI: TAdvStringGrid;
    GridS: TAdvStringGrid;
    GridB: TAdvStringGrid;
    gridPL: TAdvStringGrid;
    GridPT: TAdvStringGrid;
    GridTT: TAdvStringGrid;
    QrCruzado1: TMSQuery;
    QrCruzado2: TMSQuery;
    Qrcruzado3: TMSQuery;
    QrCruzado4: TMSQuery;
    QrVentasQ: TMSQuery;
    QrVentasQHD: TMSQuery;
    gridQ: TAdvStringGrid;
    GridQHD: TAdvStringGrid;
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
  Form2: TForm2;

implementation

{$R *.dfm}

uses unit1;

procedure TForm2.SpeedButton1Click(Sender: TObject);
var
x,y : integer;
prop,SumaImportes : single;
begin
grid1.ColCount:=71;
grid1.RowCount:=2;
grid1.LoadFromFile('config.cx2');
SumaImportes:=0;
grid1.ClearRows(1,grid1.RowCount-1);
grid1.RowCount:=2;
y:=1;
grid1.Visible:=true;
form1.qrmetas.Close;
form1.qrmetas.ParamByName('tienda').AsString:=form1.combobox1.Text;
form1.qrmetas.Open;

while not form1.qrmetas.Eof do
  begin
    grid1.cells[0,y]:=form1.qrmetas.FieldByName('nombre').AsString;
    if form1.qrmetas.FieldByName('MetaUO').AsInteger=null then   grid1.ints[2,y]:=0 else
    grid1.ints[2,y]:=form1.qrmetas.FieldByName('MetaUO').AsInteger;
    if form1.qrmetas.FieldByName('MetaUE').AsInteger=null then   grid1.ints[6,y]:=0 else
    grid1.ints[6,y]:=form1.qrmetas.FieldByName('MetaUE').AsInteger;
    if form1.qrmetas.FieldByName('MetaUEHD').AsInteger=null then   grid1.ints[10,y]:=0 else
    grid1.ints[10,y]:=form1.qrmetas.FieldByName('MetaUEHD').AsInteger;
    if form1.qrmetas.FieldByName('MetaLL').AsInteger=null then   grid1.ints[14,y]:=0 else
    grid1.ints[14,y]:=form1.qrmetas.FieldByName('MetaLL').AsInteger;
    if form1.qrmetas.FieldByName('MetaT').AsInteger=null then   grid1.ints[18,y]:=0 else
    grid1.ints[18,y]:=form1.qrmetas.FieldByName('MetaT').AsInteger;
    if form1.qrmetas.FieldByName('MetaC').AsInteger=null then   grid1.ints[22,y]:=0 else
    grid1.ints[22,y]:=form1.qrmetas.FieldByName('MetaC').AsInteger;
    if form1.qrmetas.FieldByName('MetaI').value=null then   grid1.floats[26,y]:=0 else
    grid1.floats[26,y]:=form1.qrmetas.FieldByName('MetaI').value;
    if form1.qrmetas.FieldByName('MetaS').AsInteger=null then   grid1.ints[30,y]:=0 else
    grid1.ints[30,y]:=form1.qrmetas.FieldByName('MetaS').AsInteger;
    if form1.qrmetas.FieldByName('MetaB').AsInteger=null then   grid1.ints[34,y]:=0 else
    grid1.ints[34,y]:=form1.qrmetas.FieldByName('MetaB').AsInteger;
    if form1.qrmetas.FieldByName('MetaPT').AsInteger=null then   grid1.ints[38,y]:=0 else
    grid1.ints[38,y]:=form1.qrmetas.FieldByName('MetaPT').AsInteger;
    if form1.qrmetas.FieldByName('MetaPL').AsInteger=null then   grid1.ints[42,y]:=0 else
    grid1.ints[42,y]:=form1.qrmetas.FieldByName('MetaPL').AsInteger;
    if form1.qrmetas.FieldByName('MetaQ').AsInteger=null then   grid1.ints[46,y]:=0 else
    grid1.ints[46,y]:=form1.qrmetas.FieldByName('MetaQ').AsInteger;
    if form1.qrmetas.FieldByName('MetaQHD').AsInteger=null then   grid1.ints[50,y]:=0 else
    grid1.ints[50,y]:=form1.qrmetas.FieldByName('MetaQHD').AsInteger;
    if form1.qrmetas.FieldByName('MetaTotal').AsInteger=null then   grid1.ints[54,y]:=0 else
    grid1.ints[54,y]:=form1.qrmetas.FieldByName('MetaTotal').AsInteger;

    grid1.RowCount:=grid1.RowCount+1;
    inc(y);
    form1.qrmetas.Next;
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
       QrVentasO.ParamByName('Fecha1').AsDate:=Fecha1.Date;
       QrVentasO.ParamByName('Fecha2').AsDate:=Fecha2.date;
       QrVentasO.ParamByName('Usuario').AsString:=QrGenerales.FieldByName('Usuario').AsString;
       QrVentasO.Open;

       qrventasE.close;
       QrVentasE.ParamByName('DevVenta').AsString:=QrGenerales.FieldByName('Devolucion').AsString;
       QrVentasE.ParamByName('Tienda').AsString:=QrGenerales.FieldByName('Factura').AsString;
       QrVentasE.ParamByName('Fecha1').AsDate:=Fecha1.Date;
       QrVentasE.ParamByName('Fecha2').AsDate:=Fecha2.date;
       QrVentasE.ParamByName('Usuario').AsString:=QrGenerales.FieldByName('Usuario').AsString;
       QrVentasE.Open;

       qrventasHD.close;
       QrVentasHD.ParamByName('DevVenta').AsString:=QrGenerales.FieldByName('Devolucion').AsString;
       QrVentasHD.ParamByName('Tienda').AsString:=QrGenerales.FieldByName('Factura').AsString;
       QrVentasHD.ParamByName('Fecha1').AsDate:=Fecha1.Date;
       QrVentasHD.ParamByName('Fecha2').AsDate:=Fecha2.date;
       QrVentasHD.ParamByName('Usuario').AsString:=QrGenerales.FieldByName('Usuario').AsString;
       QrVentasHD.Open;

       qrventasT.close;
       QrVentasT.ParamByName('DevVenta').AsString:=QrGenerales.FieldByName('Devolucion').AsString;
       QrVentasT.ParamByName('Tienda').AsString:=QrGenerales.FieldByName('Factura').AsString;
       QrVentasT.ParamByName('Fecha1').AsDate:=Fecha1.Date;
       QrVentasT.ParamByName('Fecha2').AsDate:=Fecha2.date;
       QrVentasT.ParamByName('Usuario').AsString:=QrGenerales.FieldByName('Usuario').AsString;
       QrVentasT.Open;

       qrventasL.close;
       QrVentasL.ParamByName('DevVenta').AsString:=QrGenerales.FieldByName('Devolucion').AsString;
       QrVentasL.ParamByName('Tienda').AsString:=QrGenerales.FieldByName('Factura').AsString;
       QrVentasL.ParamByName('Fecha1').AsDate:=Fecha1.Date;
       QrVentasL.ParamByName('Fecha2').AsDate:=Fecha2.date;
       QrVentasL.ParamByName('Usuario').AsString:=QrGenerales.FieldByName('Usuario').AsString;
       QrVentasL.Open;

       qrventasL2.close;
       QrVentasL2.ParamByName('DevVenta').AsString:=QrGenerales.FieldByName('Devolucion').AsString;
       QrVentasL2.ParamByName('Tienda').AsString:=QrGenerales.FieldByName('Factura').AsString;
       QrVentasL2.ParamByName('Fecha1').AsDate:=Fecha1.Date;
       QrVentasL2.ParamByName('Fecha2').AsDate:=Fecha2.date;
       QrVentasL2.ParamByName('Usuario').AsString:=QrGenerales.FieldByName('Usuario').AsString;
       QrVentasL2.Open;

       qrventasC.close;
       QrVentasC.ParamByName('DevVenta').AsString:=QrGenerales.FieldByName('Devolucion').AsString;
       QrVentasC.ParamByName('Tienda').AsString:=QrGenerales.FieldByName('Factura').AsString;
       QrVentasC.ParamByName('Fecha1').AsDate:=Fecha1.Date;
       QrVentasC.ParamByName('Fecha2').AsDate:=Fecha2.date;
       QrVentasC.ParamByName('Usuario').AsString:=QrGenerales.FieldByName('Usuario').AsString;
       QrVentasC.Open;

       qrventasI.close;
       QrVentasI.ParamByName('DevVenta').AsString:=QrGenerales.FieldByName('Devolucion').AsString;
       QrVentasI.ParamByName('Tienda').AsString:=QrGenerales.FieldByName('Factura').AsString;
       QrVentasI.ParamByName('Fecha1').AsDate:=Fecha1.Date;
       QrVentasI.ParamByName('Fecha2').AsDate:=Fecha2.date;
       QrVentasI.ParamByName('Usuario').AsString:=QrGenerales.FieldByName('Usuario').AsString;
       QrVentasI.Open;

       qrventasS.close;
       QrVentasS.ParamByName('DevVenta').AsString:=QrGenerales.FieldByName('Devolucion').AsString;
       QrVentasS.ParamByName('Tienda').AsString:=QrGenerales.FieldByName('Factura').AsString;
       QrVentasS.ParamByName('Fecha1').AsDate:=Fecha1.Date;
       QrVentasS.ParamByName('Fecha2').AsDate:=Fecha2.date;
       QrVentasS.ParamByName('Usuario').AsString:=QrGenerales.FieldByName('Usuario').AsString;
       QrVentasS.Open;

       qrventasB.close;
       QrVentasB.ParamByName('DevVenta').AsString:=QrGenerales.FieldByName('Devolucion').AsString;
       QrVentasB.ParamByName('Tienda').AsString:=QrGenerales.FieldByName('Factura').AsString;
       QrVentasB.ParamByName('Fecha1').AsDate:=Fecha1.Date;
       QrVentasB.ParamByName('Fecha2').AsDate:=Fecha2.date;
       QrVentasB.ParamByName('Usuario').AsString:=QrGenerales.FieldByName('Usuario').AsString;
       QrVentasB.Open;

       qrventasPT.close;
       QrVentasPT.ParamByName('DevVenta').AsString:=QrGenerales.FieldByName('Devolucion').AsString;
       QrVentasPT.ParamByName('Tienda').AsString:=QrGenerales.FieldByName('Factura').AsString;
       QrVentasPT.ParamByName('Fecha1').AsDate:=Fecha1.Date;
       QrVentasPT.ParamByName('Fecha2').AsDate:=Fecha2.date;
       QrVentasPT.ParamByName('Usuario').AsString:=QrGenerales.FieldByName('Usuario').AsString;
       QrVentasPT.Open;

       qrventasPL.close;
       QrVentasPL.ParamByName('DevVenta').AsString:=QrGenerales.FieldByName('Devolucion').AsString;
       QrVentasPL.ParamByName('Tienda').AsString:=QrGenerales.FieldByName('Factura').AsString;
       QrVentasPL.ParamByName('Fecha1').AsDate:=Fecha1.Date;
       QrVentasPL.ParamByName('Fecha2').AsDate:=Fecha2.date;
       QrVentasPL.ParamByName('Usuario').AsString:=QrGenerales.FieldByName('Usuario').AsString;
       QrVentasPL.Open;

       qrventasPL2.close;
       QrVentasPL2.ParamByName('DevVenta').AsString:=QrGenerales.FieldByName('Devolucion').AsString;
       QrVentasPL2.ParamByName('Tienda').AsString:=QrGenerales.FieldByName('Factura').AsString;
       QrVentasPL2.ParamByName('Fecha1').AsDate:=Fecha1.Date;
       QrVentasPL2.ParamByName('Fecha2').AsDate:=Fecha2.date;
       QrVentasPL2.ParamByName('Usuario').AsString:=QrGenerales.FieldByName('Usuario').AsString;
       QrVentasPL2.Open;


       qrventasQ.close;
       QrVentasQ.ParamByName('DevVenta').AsString:=QrGenerales.FieldByName('Devolucion').AsString;
       QrVentasQ.ParamByName('Tienda').AsString:=QrGenerales.FieldByName('Factura').AsString;
       QrVentasQ.ParamByName('Fecha1').AsDate:=Fecha1.Date;
       QrVentasQ.ParamByName('Fecha2').AsDate:=Fecha2.date;
       QrVentasQ.ParamByName('Usuario').AsString:=QrGenerales.FieldByName('Usuario').AsString;
       QrVentasQ.Open;

       qrventasQHD.close;
       QrVentasQHD.ParamByName('DevVenta').AsString:=QrGenerales.FieldByName('Devolucion').AsString;
       QrVentasQHD.ParamByName('Tienda').AsString:=QrGenerales.FieldByName('Factura').AsString;
       QrVentasQHD.ParamByName('Fecha1').AsDate:=Fecha1.Date;
       QrVentasQHD.ParamByName('Fecha2').AsDate:=Fecha2.date;
       QrVentasQHD.ParamByName('Usuario').AsString:=QrGenerales.FieldByName('Usuario').AsString;
       QrVentasQHD.Open;



       if qrVentasO.isempty=true then grid1.ints[1,y]:=0 else
       grid1.ints[1,y]:=qrVentasO.FieldByName('Cantidad').AsInteger;

       if qrVentasE.isempty=true then grid1.ints[5,y]:=0 else
       grid1.ints[5,y]:=qrVentasE.FieldByName('Cantidad').AsInteger;

       if qrVentasHD.isempty=true then grid1.ints[9,y]:=0 else
       grid1.ints[9,y]:=qrVentasHD.FieldByName('Cantidad').AsInteger;

       if qrVentasL.isempty=true then grid1.ints[13,y]:=0 else
       grid1.ints[13,y]:=qrVentasL.FieldByName('Cantidad').AsInteger+qrVentasL2.FieldByName('Cantidad').AsInteger;

       if qrVentasT.isempty=true then grid1.ints[17,y]:=0 else
       grid1.ints[17,y]:=qrVentasT.FieldByName('Cantidad').AsInteger;

       if qrVentasC.isempty=true then grid1.ints[21,y]:=0 else
       grid1.ints[21,y]:=qrVentasC.FieldByName('Cantidad').AsInteger;

       if qrVentasI.isempty=true then grid1.ints[25,y]:=0 else
       grid1.floats[25,y]:=qrVentasI.FieldByName('Cantidad').value;

       if qrVentasS.isempty=true then grid1.ints[29,y]:=0 else
       grid1.ints[29,y]:=qrVentasS.FieldByName('Cantidad').asinteger;

       if qrVentasB.isempty=true then grid1.ints[33,y]:=0 else
       grid1.ints[33,y]:=qrVentasB.FieldByName('Cantidad').asinteger;

       if qrVentasPT.isempty=true then grid1.ints[37,y]:=0 else
       grid1.ints[37,y]:=qrVentasPT.FieldByName('Cantidad').asinteger;

       if qrVentasPL.isempty=true then grid1.ints[41,y]:=0 else
       grid1.ints[41,y]:=qrVentasPL.FieldByName('Cantidad').AsInteger+qrVentasPL2.FieldByName('Cantidad').AsInteger;

       if qrVentasQ.isempty=true then grid1.ints[45,y]:=0 else
       grid1.ints[45,y]:=qrVentasQ.FieldByName('Cantidad').asinteger;

       if qrVentasQHD.isempty=true then grid1.ints[49,y]:=0 else
       grid1.ints[49,y]:=qrVentasQHD.FieldByName('Cantidad').asinteger;


       //suma de piezas totales
       grid1.Ints[53,y]:=grid1.Ints[1,y]+grid1.Ints[5,y]+grid1.Ints[13,y]+grid1.Ints[17,y]+grid1.Ints[21,y]+grid1.Ints[29,y]+grid1.Ints[33,y]+grid1.Ints[37,y]+grid1.Ints[41,y]+grid1.Ints[45,y]+grid1.Ints[49,y];
       //comienzan las operaciones   -siempre es meta - venta   --se cambio venta - meta
       grid1.Ints[3,y]:=grid1.Ints[1,y]-grid1.Ints[2,y];
       grid1.Ints[7,y]:=grid1.Ints[5,y]-grid1.Ints[6,y];
       grid1.Ints[11,y]:=grid1.Ints[9,y]-grid1.Ints[10,y];
       grid1.Ints[15,y]:=grid1.Ints[13,y]-grid1.Ints[14,y];
       grid1.Ints[19,y]:=grid1.Ints[17,y]-grid1.Ints[18,y];
       grid1.Ints[23,y]:=grid1.Ints[21,y]-grid1.Ints[22,y];
       grid1.Ints[27,y]:=grid1.Ints[25,y]-grid1.Ints[26,y];
       grid1.Ints[31,y]:=grid1.Ints[29,y]-grid1.Ints[30,y];
       grid1.Ints[35,y]:=grid1.Ints[33,y]-grid1.Ints[34,y];
       grid1.Ints[39,y]:=grid1.Ints[37,y]-grid1.Ints[38,y];
       grid1.Ints[43,y]:=grid1.Ints[41,y]-grid1.Ints[42,y];
       grid1.Ints[47,y]:=grid1.Ints[45,y]-grid1.Ints[46,y];
       grid1.Ints[51,y]:=grid1.Ints[49,y]-grid1.Ints[50,y];
       grid1.Ints[55,y]:=grid1.Ints[53,y]-grid1.Ints[54,y];
       //porcentajes
       if grid1.Floats[2,y]=0 then  grid1.Floats[4,y]:=0 else
       grid1.Floats[4,y]:=(grid1.Floats[1,y]*100)/grid1.Floats[2,y];

       if grid1.Floats[6,y]=0 then  grid1.Floats[8,y]:=0 else
       grid1.Floats[8,y]:=(grid1.Floats[5,y]*100)/grid1.Floats[6,y];

       if grid1.Floats[10,y]=0 then  grid1.Floats[12,y]:=0 else
       grid1.Floats[12,y]:=(grid1.Floats[9,y]*100)/grid1.Floats[10,y];

       if grid1.Floats[14,y]=0 then  grid1.Floats[16,y]:=0 else
       grid1.Floats[16,y]:=(grid1.Floats[13,y]*100)/grid1.Floats[14,y];

       if grid1.Floats[18,y]=0 then  grid1.Floats[20,y]:=0 else
       grid1.Floats[20,y]:=(grid1.Floats[17,y]*100)/grid1.Floats[18,y];

       if grid1.Floats[22,y]=0 then  grid1.Floats[24,y]:=0 else
       grid1.Floats[24,y]:=(grid1.Floats[21,y]*100)/grid1.Floats[22,y];

       if grid1.Floats[26,y]=0 then  grid1.Floats[28,y]:=0 else
       grid1.Floats[28,y]:=(grid1.Floats[25,y]*100)/grid1.Floats[26,y];

       if grid1.Floats[30,y]=0 then  grid1.Floats[32,y]:=0 else
       grid1.Floats[32,y]:=(grid1.Floats[29,y]*100)/grid1.Floats[30,y];

       if grid1.Floats[34,y]=0 then  grid1.Floats[36,y]:=0 else
       grid1.Floats[36,y]:=(grid1.Floats[33,y]*100)/grid1.Floats[34,y];

       if grid1.Floats[38,y]=0 then  grid1.Floats[40,y]:=0 else
       grid1.Floats[40,y]:=(grid1.Floats[37,y]*100)/grid1.Floats[38,y];

       if grid1.Floats[42,y]=0 then  grid1.Floats[44,y]:=0 else
       grid1.Floats[44,y]:=(grid1.Floats[41,y]*100)/grid1.Floats[42,y];

       if grid1.Floats[46,y]=0 then  grid1.Floats[48,y]:=0 else
       grid1.Floats[48,y]:=(grid1.Floats[45,y]*100)/grid1.Floats[46,y];

       if grid1.Floats[50,y]=0 then  grid1.Floats[52,y]:=0 else
       grid1.Floats[52,y]:=(grid1.Floats[49,y]*100)/grid1.Floats[50,y];

       if grid1.Floats[54,y]=0 then  grid1.Floats[56,y]:=0 else
       grid1.Floats[56,y]:=(grid1.Floats[53,y]*100)/grid1.Floats[54,y];

     end;

     //COMIENZAN DATOS GRAFICOS
      Gridg.ClearAll;
      GridE.ClearAll;
      GridH.ClearAll;
      GridL.ClearAll;
      GridT.ClearAll;
      GridC.ClearAll;
      GridI.ClearAll;
      GridS.ClearAll;
      GridB.ClearAll;
      GridPL.ClearAll;
      GridPT.ClearAll;
      GridTT.ClearAll;
      GridQ.ClearAll;
      GridQHD.ClearAll;

      gridg.RowCount:=1;
      gridE.RowCount:=1;
      gridH.RowCount:=1;
      gridL.RowCount:=1;
      gridT.RowCount:=1;
      gridC.RowCount:=1;
      gridI.RowCount:=1;
      gridS.RowCount:=1;
      gridB.RowCount:=1;
      gridPL.RowCount:=1;
      gridPT.RowCount:=1;
      gridTT.RowCount:=1;

      gridQ.RowCount:=1;
      gridQHD.RowCount:=1;


      //Original B


      for y := 1 to grid1.RowCount-1 do
        begin
          gridG.Cells[0,y-1]:=grid1.Cells[0,y];
          gridg.Floats[1,y-1]:=grid1.Floats[4,y];
          gridG.RowCount:=gridG.RowCount+1;
        end;

        gridG.RowCount:=gridG.RowCount-1;

        gridg.Sort(1,sddescending);

        for x := 1 to 10 do
          gridg.Colors[x,0]:=clskyblue;


       for y := 1 to gridg.RowCount-1 do
        begin
           if gridg.Floats[1,0]=0 then prop:=0 else
           prop:=(gridg.Floats[1,y]*10)/gridg.Floats[1,0];
           if prop<1 then prop:=2;
           for x := 1 to  trunc (prop) do
              gridg.Colors[x,y]:=clskyblue;
            gridg.Floats[x-1,y]:=gridg.Floats[1,y];
            gridg.cells[1,y]:='';
        end;
            gridg.Floats[10,0]:=gridg.Floats[1,0];
            gridg.cells[1,0]:='';

        gridg.InsertRows(0,1);
        gridg.Cells[0,0]:='META VENTA ORIGINAL %';
        gridg.RowCount:=gridG.RowCount+1;
       //original E

       //Evoke B
       prop:=0;

       for y := 1 to grid1.RowCount-1 do
        begin
          gridE.Cells[0,y-1]:=grid1.Cells[0,y];
          gridE.Floats[1,y-1]:=grid1.Floats[8,y];
          gridE.RowCount:=gridE.RowCount+1;
        end;

        gridE.RowCount:=gridE.RowCount-1;

        gridE.Sort(1,sddescending);

        for x := 1 to 10 do
          gridE.Colors[x,0]:=clskyblue;

       for y := 1 to gridE.RowCount-1 do
        begin
           if gridE.Floats[1,0]=0 then prop:=0 else
           prop:=(gridE.Floats[1,y]*10)/gridE.Floats[1,0];
           if prop<1 then prop:=2;

           for x := 1 to  trunc (prop) do
              gridE.Colors[x,y]:=clskyblue;
            gridE.Floats[x-1,y]:=gridE.Floats[1,y];
            gridE.cells[1,y]:='';
        end;
            gridE.Floats[10,0]:=gridE.Floats[1,0];
            gridE.cells[1,0]:='';

            gridE.InsertRows(0,1);
            gridE.Cells[0,0]:='META VENTA EVOKE %';
            gridE.RowCount:=gridE.RowCount+1;

        ///Evoke E

      //HD B
         prop:=0;

       for y := 1 to grid1.RowCount-1 do
        begin
          gridH.Cells[0,y-1]:=grid1.Cells[0,y];
          gridH.Floats[1,y-1]:=grid1.Floats[12,y];
          gridH.RowCount:=gridH.RowCount+1;
        end;

        gridH.RowCount:=gridH.RowCount-1;

        gridH.Sort(1,sddescending);

        for x := 1 to 10 do
          gridH.Colors[x,0]:=clskyblue;

       for y := 1 to gridH.RowCount-1 do
        begin
           if gridH.Floats[1,0]=0 then prop:=0 else
           prop:=(gridH.Floats[1,y]*10)/gridH.Floats[1,0];
           if prop<1 then prop:=2;

           for x := 1 to  trunc (prop) do
              gridH.Colors[x,y]:=clskyblue;
            gridH.Floats[x-1,y]:=gridH.Floats[1,y];
            gridH.cells[1,y]:='';
        end;
            gridH.Floats[10,0]:=gridH.Floats[1,0];
            gridH.cells[1,0]:='';

            gridH.InsertRows(0,1);
            gridH.Cells[0,0]:='META VENTA HD %';
            gridH.RowCount:=gridH.RowCount+1;
     //HD E

     //Llaves B
       prop:=0;

       for y := 1 to grid1.RowCount-1 do
        begin
          gridL.Cells[0,y-1]:=grid1.Cells[0,y];
          gridL.Floats[1,y-1]:=grid1.Floats[16,y];
          gridL.RowCount:=gridL.RowCount+1;
        end;

        gridL.RowCount:=gridL.RowCount-1;

        gridL.Sort(1,sddescending);

        for x := 1 to 10 do
          gridL.Colors[x,0]:=clskyblue;

       for y := 1 to gridL.RowCount-1 do
        begin
           if gridL.Floats[1,0]=0 then prop:=0 else
           prop:=(gridL.Floats[1,y]*10)/gridL.Floats[1,0];
           if prop<1 then prop:=2;

           for x := 1 to  trunc (prop) do
              gridL.Colors[x,y]:=clskyblue;
            gridL.Floats[x-1,y]:=gridL.Floats[1,y];
            gridL.cells[1,y]:='';
        end;
            gridL.Floats[10,0]:=gridL.Floats[1,0];
            gridL.cells[1,0]:='';

            gridL.InsertRows(0,1);
            gridL.Cells[0,0]:='META VENTA LLAVES %';
            gridL.RowCount:=gridL.RowCount+1;

     //Llaves E

     //Tarjas B
       prop:=0;

       for y := 1 to grid1.RowCount-1 do
        begin
          gridT.Cells[0,y-1]:=grid1.Cells[0,y];
          gridT.Floats[1,y-1]:=grid1.Floats[20,y];
          gridT.RowCount:=gridT.RowCount+1;
        end;

        gridT.RowCount:=gridT.RowCount-1;

        gridT.Sort(1,sddescending);

        for x := 1 to 10 do
          gridT.Colors[x,0]:=clskyblue;

       for y := 1 to gridT.RowCount-1 do
        begin
           if gridT.Floats[1,0]=0 then prop:=0 else
           prop:=(gridT.Floats[1,y]*10)/gridT.Floats[1,0];
           if prop<1 then prop:=2;

           for x := 1 to  trunc (prop) do
              gridT.Colors[x,y]:=clskyblue;
            gridT.Floats[x-1,y]:=gridT.Floats[1,y];
            gridT.cells[1,y]:='';
        end;
            gridT.Floats[10,0]:=gridT.Floats[1,0];
            gridT.cells[1,0]:='';

            gridT.InsertRows(0,1);
            gridT.Cells[0,0]:='META VENTA TARJAS %';
            gridT.RowCount:=gridT.RowCount+1;

     //Tarjas E

     //complementos B
       prop:=0;

       for y := 1 to grid1.RowCount-1 do
        begin
          gridC.Cells[0,y-1]:=grid1.Cells[0,y];
          gridC.Floats[1,y-1]:=grid1.Floats[24,y];
          gridC.RowCount:=gridC.RowCount+1;
        end;

        gridC.RowCount:=gridC.RowCount-1;

        gridC.Sort(1,sddescending);

        for x := 1 to 10 do
          gridC.Colors[x,0]:=clskyblue;

       for y := 1 to gridC.RowCount-1 do
        begin
           if gridC.Floats[1,0]=0 then prop:=0 else
           prop:=(gridC.Floats[1,y]*10)/gridC.Floats[1,0];
           if prop<1 then prop:=2;

           for x := 1 to  trunc (prop) do
              gridC.Colors[x,y]:=clskyblue;
            gridC.Floats[x-1,y]:=gridC.Floats[1,y];
            gridC.cells[1,y]:='';
        end;
            gridC.Floats[10,0]:=gridC.Floats[1,0];
            gridC.cells[1,0]:='';

        gridC.InsertRows(0,1);
        gridC.Cells[0,0]:='META VENTA COMPLEMENTOS %';
        gridC.RowCount:=gridC.RowCount+1;


     //complementos E

//Innovika B
       prop:=0;

       for y := 1 to grid1.RowCount-1 do
        begin
          gridI.Cells[0,y-1]:=grid1.Cells[0,y];
          gridI.Floats[1,y-1]:=grid1.Floats[28,y];
          gridI.RowCount:=gridI.RowCount+1;
        end;

        gridI.RowCount:=gridI.RowCount-1;

        gridI.Sort(1,sddescending);

        for x := 1 to 10 do
          gridI.Colors[x,0]:=clskyblue;

       for y := 1 to gridI.RowCount-1 do
        begin
           if gridI.Floats[1,0]=0 then prop:=0 else
           prop:=(gridI.Floats[1,y]*10)/gridI.Floats[1,0];
           if prop<1 then prop:=2;

           for x := 1 to  trunc (prop) do
              gridI.Colors[x,y]:=clskyblue;
            gridI.Floats[x-1,y]:=gridI.Floats[1,y];
            gridI.cells[1,y]:='';
        end;
            gridI.Floats[10,0]:=gridI.Floats[1,0];
            gridI.cells[1,0]:='';

        gridI.InsertRows(0,1);
        gridI.Cells[0,0]:='META VENTA INNOVIKA %';
        gridI.RowCount:=gridI.RowCount+1;

     //Innovika E


//Slim B
       prop:=0;

       for y := 1 to grid1.RowCount-1 do
        begin
          gridS.Cells[0,y-1]:=grid1.Cells[0,y];
          gridS.Floats[1,y-1]:=grid1.Floats[32,y];
          gridS.RowCount:=gridS.RowCount+1;
        end;

        gridS.RowCount:=gridS.RowCount-1;

        gridS.Sort(1,sddescending);

        for x := 1 to 10 do
          gridS.Colors[x,0]:=clskyblue;

       for y := 1 to gridS.RowCount-1 do
        begin
           if gridS.Floats[1,0]=0 then prop:=0 else
           prop:=(gridS.Floats[1,y]*10)/gridS.Floats[1,0];
           if prop<1 then prop:=2;

           for x := 1 to  trunc (prop) do
            gridS.Colors[x,y]:=clskyblue;
            gridS.Floats[x-1,y]:=gridS.Floats[1,y];
            gridS.cells[1,y]:='';
        end;
            gridS.Floats[10,0]:=gridS.Floats[1,0];
            gridS.cells[1,0]:='';

        gridS.InsertRows(0,1);
        gridS.Cells[0,0]:='META VENTA SLIM %';
        gridS.RowCount:=gridS.RowCount+1;

     //Slim E


//BASIK B
       prop:=0;

       for y := 1 to grid1.RowCount-1 do
        begin
          gridB.Cells[0,y-1]:=grid1.Cells[0,y];
          gridB.Floats[1,y-1]:=grid1.Floats[36,y];
          gridB.RowCount:=gridB.RowCount+1;
        end;

        gridB.RowCount:=gridB.RowCount-1;

        gridB.Sort(1,sddescending);

        for x := 1 to 10 do
          gridB.Colors[x,0]:=clskyblue;

       for y := 1 to gridB.RowCount-1 do
        begin
           if gridB.Floats[1,0]=0 then prop:=0 else
           prop:=(gridB.Floats[1,y]*10)/gridB.Floats[1,0];
           if prop<1 then prop:=2;

           for x := 1 to  trunc (prop) do
              gridB.Colors[x,y]:=clskyblue;
            gridB.Floats[x-1,y]:=gridB.Floats[1,y];
            gridB.cells[1,y]:='';
        end;
            gridB.Floats[10,0]:=gridB.Floats[1,0];
            gridB.cells[1,0]:='';

            gridB.InsertRows(0,1);
            gridB.Cells[0,0]:='META VENTA BASI-K %';
            gridB.RowCount:=gridB.RowCount+1;

     //BASIK E

//Plados Tarjas B
       prop:=0;

       for y := 1 to grid1.RowCount-1 do
        begin
          gridPL.Cells[0,y-1]:=grid1.Cells[0,y];
          gridPL.Floats[1,y-1]:=grid1.Floats[40,y];
          gridPL.RowCount:=gridPL.RowCount+1;
        end;

        gridPL.RowCount:=gridPL.RowCount-1;

        gridPL.Sort(1,sddescending);

        for x := 1 to 10 do
          gridPL.Colors[x,0]:=clskyblue;

       for y := 1 to gridPL.RowCount-1 do
        begin
           if gridPL.Floats[1,0]=0 then prop:=0 else
           prop:=(gridPL.Floats[1,y]*10)/gridPL.Floats[1,0];
           if prop<1 then prop:=2;

           for x := 1 to  trunc (prop) do
              gridPL.Colors[x,y]:=clskyblue;
            gridPL.Floats[x-1,y]:=gridPL.Floats[1,y];
            gridPL.cells[1,y]:='';
        end;
            gridPL.Floats[10,0]:=gridPL.Floats[1,0];
            gridPL.cells[1,0]:='';

        gridPL.InsertRows(0,1);
        gridPL.Cells[0,0]:='META VENTA PLADOS TARJAS %';
        gridPL.RowCount:=gridPL.RowCount+1;

     //Plados Tarjas E

//Plados LLaves B
       prop:=0;

       for y := 1 to grid1.RowCount-1 do
        begin
          gridPT.Cells[0,y-1]:=grid1.Cells[0,y];
          gridPT.Floats[1,y-1]:=grid1.Floats[44,y];
          gridPT.RowCount:=gridPT.RowCount+1;
        end;

        gridPT.RowCount:=gridPT.RowCount-1;

        gridPT.Sort(1,sddescending);

        for x := 1 to 10 do
          gridPT.Colors[x,0]:=clskyblue;

       for y := 1 to gridPT.RowCount-1 do
        begin
           if gridPT.Floats[1,0]=0 then prop:=0 else
           prop:=(gridPT.Floats[1,y]*10)/gridPT.Floats[1,0];
           if prop<1 then prop:=2;

           for x := 1 to  trunc (prop) do
              gridPT.Colors[x,y]:=clskyblue;
            gridPT.Floats[x-1,y]:=gridPT.Floats[1,y];
            gridPT.cells[1,y]:='';
        end;
            gridPT.Floats[10,0]:=gridPT.Floats[1,0];
            gridPT.cells[1,0]:='';

        gridPT.InsertRows(0,1);
        gridPT.Cells[0,0]:='META VENTA PLADOS LLAVES %';
        gridPT.RowCount:=gridPT.RowCount+1;

     //Plados Llaves E


      //Q B
       prop:=0;

       for y := 1 to grid1.RowCount-1 do
        begin
          gridQ.Cells[0,y-1]:=grid1.Cells[0,y];
          gridQ.Floats[1,y-1]:=grid1.Floats[48,y];
          gridQ.RowCount:=gridQ.RowCount+1;
        end;

        gridQ.RowCount:=gridQ.RowCount-1;

        gridQ.Sort(1,sddescending);

        for x := 1 to 10 do
          gridQ.Colors[x,0]:=clskyblue;

       for y := 1 to gridQ.RowCount-1 do
        begin
           if gridQ.Floats[1,0]=0 then prop:=0 else
           prop:=(gridQ.Floats[1,y]*10)/gridQ.Floats[1,0];
           if prop<1 then prop:=2;

           for x := 1 to  trunc (prop) do
              gridQ.Colors[x,y]:=clskyblue;
            gridQ.Floats[x-1,y]:=gridQ.Floats[1,y];
            gridQ.cells[1,y]:='';
        end;
            gridQ.Floats[10,0]:=gridQ.Floats[1,0];
            gridQ.cells[1,0]:='';

            gridQ.InsertRows(0,1);
            gridQ.Cells[0,0]:='META VENTA Q %';
            gridQ.RowCount:=gridQ.RowCount+1;

     //Q E


     //QHD B
       prop:=0;

       for y := 1 to grid1.RowCount-1 do
        begin
          gridQHD.Cells[0,y-1]:=grid1.Cells[0,y];
          gridQHD.Floats[1,y-1]:=grid1.Floats[52,y];
          gridQHD.RowCount:=gridQHD.RowCount+1;
        end;

        gridQHD.RowCount:=gridQHD.RowCount-1;

        gridQHD.Sort(1,sddescending);

        for x := 1 to 10 do
          gridQHD.Colors[x,0]:=clskyblue;

       for y := 1 to gridQHD.RowCount-1 do
        begin
           if gridQHD.Floats[1,0]=0 then prop:=0 else
           prop:=(gridQHD.Floats[1,y]*10)/gridQHD.Floats[1,0];
           if prop<1 then prop:=2;

           for x := 1 to  trunc (prop) do
              gridQHD.Colors[x,y]:=clskyblue;
            gridQHD.Floats[x-1,y]:=gridQHD.Floats[1,y];
            gridQHD.cells[1,y]:='';
        end;
            gridQHD.Floats[10,0]:=gridQHD.Floats[1,0];
            gridQHD.cells[1,0]:='';

            gridQHD.InsertRows(0,1);
            gridQHD.Cells[0,0]:='META VENTA Q HD%';
            gridQHD.RowCount:=gridQHD.RowCount+1;

     //QHD E



//Total B
       prop:=0;

       for y := 1 to grid1.RowCount-1 do
        begin
          gridTT.Cells[0,y-1]:=grid1.Cells[0,y];
          gridTT.Floats[1,y-1]:=grid1.Floats[56,y];
          gridTT.RowCount:=gridTT.RowCount+1;
        end;

        gridTT.RowCount:=gridTT.RowCount-1;

        gridTT.Sort(1,sddescending);

        for x := 1 to 10 do
          gridTT.Colors[x,0]:=clskyblue;

       for y := 1 to gridTT.RowCount-1 do
        begin
           if gridTT.Floats[1,0]=0 then prop:=0 else
           prop:=(gridTT.Floats[1,y]*10)/gridTT.Floats[1,0];
           if prop<1 then prop:=2;

           for x := 1 to  trunc (prop) do
              gridTT.Colors[x,y]:=clskyblue;
            gridTT.Floats[x-1,y]:=gridTT.Floats[1,y];
            gridTT.cells[1,y]:='';
        end;
            gridTT.Floats[10,0]:=gridTT.Floats[1,0];
            gridTT.cells[1,0]:='';

            gridTT.InsertRows(0,1);
            gridTT.Cells[0,0]:='META VENTA TOTAL %';
            gridTT.RowCount:=gridTT.RowCount+1;

     //Total E


     //formateo de %
 for y := 1 to grid1.RowCount-1 do
   begin
     grid1.Cells[46,y]:=FormatFloat('#,##0',grid1.Floats[46,y]);

     grid1.Cells[4,y]:=FormatFloat('##0%',grid1.Floats[4,y]);
     grid1.Cells[8,y]:=FormatFloat('##0%',grid1.Floats[8,y]);
     grid1.Cells[12,y]:=FormatFloat('##0%',grid1.Floats[12,y]);
     grid1.Cells[16,y]:=FormatFloat('##0%',grid1.Floats[16,y]);
     grid1.Cells[20,y]:=FormatFloat('##0%',grid1.Floats[20,y]);
     grid1.Cells[24,y]:=FormatFloat('##0%',grid1.Floats[24,y]);
     grid1.Cells[28,y]:=FormatFloat('##0%',grid1.Floats[28,y]);
     grid1.Cells[32,y]:=FormatFloat('##0%',grid1.Floats[32,y]);
     grid1.Cells[36,y]:=FormatFloat('##0%',grid1.Floats[36,y]);
     grid1.Cells[40,y]:=FormatFloat('##0%',grid1.Floats[40,y]);
     grid1.Cells[44,y]:=FormatFloat('##0%',grid1.Floats[44,y]);
     grid1.Cells[48,y]:=FormatFloat('##0%',grid1.Floats[48,y]);
     grid1.Cells[52,y]:=FormatFloat('##0%',grid1.Floats[52,y]);
     grid1.Cells[56,y]:=FormatFloat('##0%',grid1.Floats[56,y]);



   end;

   //comienzan importes de cada categoria
    for y := 1 to grid1.RowCount-1 do
     begin
       SumaImportes:=0;
       qrGenerales.Close;
       qrGenerales.ParamByName('nombre').AsString:=grid1.Cells[0,y];
       qrGenerales.Open;

       QrImportesCABA.close;
       QrImportesCABA.ParamByName('DevVenta').AsString:=QrGenerales.FieldByName('Devolucion').AsString;
       QrImportesCABA.ParamByName('Tienda').AsString:=QrGenerales.FieldByName('Factura').AsString;
       QrImportesCABA.ParamByName('Fecha1').AsDate:=Fecha1.Date;
       QrImportesCABA.ParamByName('Fecha2').AsDate:=Fecha2.date;
       QrImportesCABA.ParamByName('Usuario').AsString:=QrGenerales.FieldByName('Usuario').AsString;
       QrImportesCABA.Open;

       QrImporteEV.close;
       QrImporteEV.ParamByName('DevVenta').AsString:=QrGenerales.FieldByName('Devolucion').AsString;
       QrImporteEV.ParamByName('Tienda').AsString:=QrGenerales.FieldByName('Factura').AsString;
       QrImporteEV.ParamByName('Fecha1').AsDate:=Fecha1.Date;
       QrImporteEV.ParamByName('Fecha2').AsDate:=Fecha2.date;
       QrImporteEV.ParamByName('Usuario').AsString:=QrGenerales.FieldByName('Usuario').AsString;
       QrImporteEV.Open;

       QrImporteHD.close;
       QrImporteHD.ParamByName('DevVenta').AsString:=QrGenerales.FieldByName('Devolucion').AsString;
       QrImporteHD.ParamByName('Tienda').AsString:=QrGenerales.FieldByName('Factura').AsString;
       QrImporteHD.ParamByName('Fecha1').AsDate:=Fecha1.Date;
       QrImporteHD.ParamByName('Fecha2').AsDate:=Fecha2.date;
       QrImporteHD.ParamByName('Usuario').AsString:=QrGenerales.FieldByName('Usuario').AsString;
       QrImporteHD.Open;

       QrImporteLL.close;
       QrImporteLL.ParamByName('DevVenta').AsString:=QrGenerales.FieldByName('Devolucion').AsString;
       QrImporteLL.ParamByName('Tienda').AsString:=QrGenerales.FieldByName('Factura').AsString;
       QrImporteLL.ParamByName('Fecha1').AsDate:=Fecha1.Date;
       QrImporteLL.ParamByName('Fecha2').AsDate:=Fecha2.date;
       QrImporteLL.ParamByName('Usuario').AsString:=QrGenerales.FieldByName('Usuario').AsString;
       QrImporteLL.Open;

       QrImporteLL2.close;
       QrImporteLL2.ParamByName('DevVenta').AsString:=QrGenerales.FieldByName('Devolucion').AsString;
       QrImporteLL2.ParamByName('Tienda').AsString:=QrGenerales.FieldByName('Factura').AsString;
       QrImporteLL2.ParamByName('Fecha1').AsDate:=Fecha1.Date;
       QrImporteLL2.ParamByName('Fecha2').AsDate:=Fecha2.date;
       QrImporteLL2.ParamByName('Usuario').AsString:=QrGenerales.FieldByName('Usuario').AsString;
       QrImporteLL2.Open;

       QrImporteTA.close;
       QrImporteTA.ParamByName('DevVenta').AsString:=QrGenerales.FieldByName('Devolucion').AsString;
       QrImporteTA.ParamByName('Tienda').AsString:=QrGenerales.FieldByName('Factura').AsString;
       QrImporteTA.ParamByName('Fecha1').AsDate:=Fecha1.Date;
       QrImporteTA.ParamByName('Fecha2').AsDate:=Fecha2.date;
       QrImporteTA.ParamByName('Usuario').AsString:=QrGenerales.FieldByName('Usuario').AsString;
       QrImporteTA.Open;


       QrImporteCO.close;
       QrImporteCO.ParamByName('DevVenta').AsString:=QrGenerales.FieldByName('Devolucion').AsString;
       QrImporteCO.ParamByName('Tienda').AsString:=QrGenerales.FieldByName('Factura').AsString;
       QrImporteCO.ParamByName('Fecha1').AsDate:=Fecha1.Date;
       QrImporteCO.ParamByName('Fecha2').AsDate:=Fecha2.date;
       QrImporteCO.ParamByName('Usuario').AsString:=QrGenerales.FieldByName('Usuario').AsString;
       QrImporteCO.Open;

       QrImporteIN.close;
       QrImporteIN.ParamByName('DevVenta').AsString:=QrGenerales.FieldByName('Devolucion').AsString;
       QrImporteIN.ParamByName('Tienda').AsString:=QrGenerales.FieldByName('Factura').AsString;
       QrImporteIN.ParamByName('Fecha1').AsDate:=Fecha1.Date;
       QrImporteIN.ParamByName('Fecha2').AsDate:=Fecha2.date;
       QrImporteIN.ParamByName('Usuario').AsString:=QrGenerales.FieldByName('Usuario').AsString;
       QrImporteIN.Open;

       QrImporteSL.close;
       QrImporteSL.ParamByName('DevVenta').AsString:=QrGenerales.FieldByName('Devolucion').AsString;
       QrImporteSL.ParamByName('Tienda').AsString:=QrGenerales.FieldByName('Factura').AsString;
       QrImporteSL.ParamByName('Fecha1').AsDate:=Fecha1.Date;
       QrImporteSL.ParamByName('Fecha2').AsDate:=Fecha2.date;
       QrImporteSL.ParamByName('Usuario').AsString:=QrGenerales.FieldByName('Usuario').AsString;
       QrImporteSL.Open;

       QrImporteK.close;
       QrImporteK.ParamByName('DevVenta').AsString:=QrGenerales.FieldByName('Devolucion').AsString;
       QrImporteK.ParamByName('Tienda').AsString:=QrGenerales.FieldByName('Factura').AsString;
       QrImporteK.ParamByName('Fecha1').AsDate:=Fecha1.Date;
       QrImporteK.ParamByName('Fecha2').AsDate:=Fecha2.date;
       QrImporteK.ParamByName('Usuario').AsString:=QrGenerales.FieldByName('Usuario').AsString;
       QrImporteK.Open;

       QrImportePT.close;
       QrImportePT.ParamByName('DevVenta').AsString:=QrGenerales.FieldByName('Devolucion').AsString;
       QrImportePT.ParamByName('Tienda').AsString:=QrGenerales.FieldByName('Factura').AsString;
       QrImportePT.ParamByName('Fecha1').AsDate:=Fecha1.Date;
       QrImportePT.ParamByName('Fecha2').AsDate:=Fecha2.date;
       QrImportePT.ParamByName('Usuario').AsString:=QrGenerales.FieldByName('Usuario').AsString;
       QrImportePT.Open;

       QrImportePL.close;
       QrImportePL.ParamByName('DevVenta').AsString:=QrGenerales.FieldByName('Devolucion').AsString;
       QrImportePL.ParamByName('Tienda').AsString:=QrGenerales.FieldByName('Factura').AsString;
       QrImportePL.ParamByName('Fecha1').AsDate:=Fecha1.Date;
       QrImportePL.ParamByName('Fecha2').AsDate:=Fecha2.date;
       QrImportePL.ParamByName('Usuario').AsString:=QrGenerales.FieldByName('Usuario').AsString;
       QrImportePL.Open;

       QrImportePL2.close;
       QrImportePL2.ParamByName('DevVenta').AsString:=QrGenerales.FieldByName('Devolucion').AsString;
       QrImportePL2.ParamByName('Tienda').AsString:=QrGenerales.FieldByName('Factura').AsString;
       QrImportePL2.ParamByName('Fecha1').AsDate:=Fecha1.Date;
       QrImportePL2.ParamByName('Fecha2').AsDate:=Fecha2.date;
       QrImportePL2.ParamByName('Usuario').AsString:=QrGenerales.FieldByName('Usuario').AsString;
       QrImportePL2.Open;

       QrImporteQ.close;
       QrImporteQ.ParamByName('DevVenta').AsString:=QrGenerales.FieldByName('Devolucion').AsString;
       QrImporteQ.ParamByName('Tienda').AsString:=QrGenerales.FieldByName('Factura').AsString;
       QrImporteQ.ParamByName('Fecha1').AsDate:=Fecha1.Date;
       QrImporteQ.ParamByName('Fecha2').AsDate:=Fecha2.date;
       QrImporteQ.ParamByName('Usuario').AsString:=QrGenerales.FieldByName('Usuario').AsString;
       QrImporteQ.Open;

       if QrImportesCABA.isempty=true then grid1.floats[57,y]:=0 else
       grid1.floats[57,y]:=QrImportesCABA.FieldByName('Importe').value;
       SumaImportes:=SumaImportes+grid1.floats[57,y];

       if QrImporteEV.isempty=true then grid1.floats[58,y]:=0 else
       grid1.floats[58,y]:=QrImporteEV.FieldByName('Importe').value;
       SumaImportes:=SumaImportes+grid1.floats[58,y];

       if QrImporteHD.isempty=true then grid1.floats[59,y]:=0 else
       grid1.floats[59,y]:=QrImporteHD.FieldByName('Importe').value;
       //SumaImportes:=SumaImportes+grid1.floats[59,y];

       if QrImporteLL.isempty=true then grid1.floats[60,y]:=0 else
       grid1.floats[60,y]:=QrImporteLL.FieldByName('Importe').value;


       if QrImporteLL2.isempty=true then grid1.floats[60,y]:=grid1.floats[60,y]+0 else
       grid1.floats[60,y]:=grid1.floats[60,y]+QrImporteLL2.FieldByName('Importe').value;
       SumaImportes:=SumaImportes+grid1.floats[60,y];

       if QrImporteTA.isempty=true then grid1.floats[61,y]:=0 else
       grid1.floats[61,y]:=QrImporteTA.FieldByName('Importe').value;
       SumaImportes:=SumaImportes+grid1.floats[61,y];

       if QrImporteCO.isempty=true then grid1.floats[62,y]:=0 else
       grid1.floats[62,y]:=QrImporteCO.FieldByName('Importe').value;
       SumaImportes:=SumaImportes+grid1.floats[62,y];

       if QrImporteIN.isempty=true then grid1.floats[63,y]:=0 else
       grid1.floats[63,y]:=QrImporteIN.FieldByName('Importe').value;
       SumaImportes:=SumaImportes+grid1.floats[63,y];

       if QrImporteSL.isempty=true then grid1.floats[64,y]:=0 else
       grid1.floats[64,y]:=QrImporteSL.FieldByName('Importe').value;
       SumaImportes:=SumaImportes+grid1.floats[64,y];

       if QrImporteK.isempty=true then grid1.floats[65,y]:=0 else
       grid1.floats[65,y]:=QrImporteK.FieldByName('Importe').value;
       SumaImportes:=SumaImportes+grid1.floats[65,y];

       if QrImportePT.isempty=true then grid1.floats[66,y]:=0 else
       grid1.floats[66,y]:=QrImportePT.FieldByName('Importe').value;
       SumaImportes:=SumaImportes+grid1.floats[66,y];

       if QrImportePL.isempty=true then grid1.floats[67,y]:=0 else
       grid1.floats[67,y]:=QrImportePL.FieldByName('Importe').value;


       if QrImportePL2.isempty=true then grid1.floats[67,y]:=grid1.floats[67,y]+0 else
       grid1.floats[67,y]:=grid1.floats[67,y]+QrImportePL2.FieldByName('Importe').value;
       SumaImportes:=SumaImportes+grid1.floats[67,y];

       if QrImporteQ.isempty=true then grid1.floats[68,y]:=0 else
       grid1.floats[68,y]:=QrImporteQ.FieldByName('Importe').value;
       SumaImportes:=SumaImportes+grid1.floats[68,y];

       if QrImporteQHD.isempty=true then grid1.floats[69,y]:=0 else
       grid1.floats[69,y]:=QrImporteQHD.FieldByName('Importe').value;
       SumaImportes:=SumaImportes+grid1.floats[69,y];


       Grid1.Floats[70,y]:=SumaImportes;


     end;

     //mover columnas de importe a cada lugar
     grid1.MoveColumn(57,5); //CABA
     grid1.MoveColumn(58,10); //Evoke
     grid1.MoveColumn(59,15); //HD
     grid1.MoveColumn(60,20); //Llaves
     grid1.MoveColumn(61,25); //Tarjas
     grid1.MoveColumn(62,30); //Complementos
     grid1.MoveColumn(63,35); //Innovika
     grid1.MoveColumn(64,40); //Slim
     grid1.MoveColumn(65,45); //Basi-K
     grid1.MoveColumn(66,50); //Plados T
     grid1.MoveColumn(67,55); //Plados L
     grid1.MoveColumn(68,60); //Q sin HD
     grid1.MoveColumn(69,65); //Q sin HD



end;

procedure TForm2.SpeedButton2Click(Sender: TObject);
begin
Form2.Visible:=false;
form1.Visible:=true;
end;

procedure TForm2.SpeedButton3Click(Sender: TObject);
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

excel1.AdvStringGrid:=gridTT;
excel1.Options.ExportOverwrite := omNever;
excel1.GridStartCol:=0;
Excel1.XLSExport(save1.FileName,'Graficos',-1,1,InsertInSheet_InsertRows);

excel1.AdvStringGrid:=gridQHD;
excel1.Options.ExportOverwrite := omNever;
excel1.GridStartCol:=0;
Excel1.XLSExport(save1.FileName,'Graficos',5,1,InsertInSheet_InsertRows);

excel1.AdvStringGrid:=gridQ;
excel1.Options.ExportOverwrite := omNever;
excel1.GridStartCol:=0;
Excel1.XLSExport(save1.FileName,'Graficos',5,1,InsertInSheet_InsertRows);

excel1.AdvStringGrid:=gridPT;
excel1.Options.ExportOverwrite := omNever;
excel1.GridStartCol:=0;
Excel1.XLSExport(save1.FileName,'Graficos',5,1,InsertInSheet_InsertRows);

excel1.AdvStringGrid:=gridPL;
excel1.Options.ExportOverwrite := omNever;
excel1.GridStartCol:=0;
Excel1.XLSExport(save1.FileName,'Graficos',5,1,InsertInSheet_InsertRows);

excel1.AdvStringGrid:=gridB;
excel1.Options.ExportOverwrite := omNever;
excel1.GridStartCol:=0;
Excel1.XLSExport(save1.FileName,'Graficos',5,1,InsertInSheet_InsertRows);

excel1.AdvStringGrid:=gridS;
excel1.Options.ExportOverwrite := omNever;
excel1.GridStartCol:=0;
Excel1.XLSExport(save1.FileName,'Graficos',5,1,InsertInSheet_InsertRows);

excel1.AdvStringGrid:=gridI;
excel1.Options.ExportOverwrite := omNever;
excel1.GridStartCol:=0;
Excel1.XLSExport(save1.FileName,'Graficos',5,1,InsertInSheet_InsertRows);

excel1.AdvStringGrid:=gridC;
excel1.Options.ExportOverwrite := omNever;
excel1.GridStartCol:=0;
Excel1.XLSExport(save1.FileName,'Graficos',5,1,InsertInSheet_InsertRows);

excel1.AdvStringGrid:=gridT;
excel1.Options.ExportOverwrite := omNever;
excel1.GridStartCol:=0;
Excel1.XLSExport(save1.FileName,'Graficos',5,1,InsertInSheet_InsertRows);

excel1.AdvStringGrid:=gridL;
excel1.Options.ExportOverwrite := omNever;
excel1.GridStartCol:=0;
Excel1.XLSExport(save1.FileName,'Graficos',5,1,InsertInSheet_InsertRows);

excel1.AdvStringGrid:=gridH;
excel1.Options.ExportOverwrite := omNever;
excel1.GridStartCol:=0;
Excel1.XLSExport(save1.FileName,'Graficos',5,1,InsertInSheet_InsertRows);

excel1.AdvStringGrid:=gridE;
excel1.Options.ExportOverwrite := omNever;
excel1.GridStartCol:=0;
Excel1.XLSExport(save1.FileName,'Graficos',5,1,InsertInSheet_InsertRows);

excel1.AdvStringGrid:=gridG;
excel1.Options.ExportOverwrite := omNever;
excel1.GridStartCol:=0;
Excel1.XLSExport(save1.FileName,'Graficos',5,1,InsertInSheet_InsertRows);



showmessage('Reporte Exportado con Exito!');
end;

end.
