unit Unit6;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, Grids, AdvObj, BaseGrid, AdvGrid, ComCtrls, Buttons,
  StdCtrls, tmsAdvGridExcel, DB, MemDS, DBAccess, MSAccess,DateUtils;

type
  TForm6 = class(TForm)
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
  Form6: TForm6;

implementation

{$R *.dfm}

uses unit1,unit2;

procedure TForm6.SpeedButton1Click(Sender: TObject);
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
      if x=1 then begin grid1.Cells[1,y]:=mes1+' '+ano1; grid1.Ints[8,y]:=mes1n; grid1.Cells[9,y]:=ano1 end;
      if x=2 then begin grid1.Cells[1,y]:=mes2+' '+ano2; grid1.Ints[8,y]:=mes2n; grid1.Cells[9,y]:=ano2 end;
      if x=3 then begin grid1.Cells[1,y]:=mes3+' '+ano3; grid1.Ints[8,y]:=mes3n; grid1.Cells[9,y]:=ano3 end;
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

       form2.qrventasCUB.close;
       form2.QrVentasCUB.ParamByName('DevVenta').AsString:=form2.QrGenerales.FieldByName('Devolucion').AsString;
       form2.QrVentasCUB.ParamByName('Tienda').AsString:=form2.QrGenerales.FieldByName('Factura').AsString;
       form2.QrVentasCUB.ParamByName('Fecha1').AsDate:=encodedate(grid1.Ints[9,y],grid1.Ints[8,y],1);
       form2.QrVentasCUB.ParamByName('Fecha2').AsDate:=encodedate(grid1.Ints[9,y],grid1.Ints[8,y],daysinamonth(grid1.Ints[9,y],grid1.Ints[8,y]));
       form2.QrVentasCUB.ParamByName('Usuario').AsString:=form2.QrGenerales.FieldByName('Usuario').AsString;
       form2.QrVentasCUB.Open;

       form2.qrventasT.close;
       form2.QrVentasT.ParamByName('DevVenta').AsString:=form2.QrGenerales.FieldByName('Devolucion').AsString;
       form2.QrVentasT.ParamByName('Tienda').AsString:=form2.QrGenerales.FieldByName('Factura').AsString;
       form2.QrVentasT.ParamByName('Fecha1').AsDate:=encodedate(grid1.Ints[9,y],grid1.Ints[8,y],1);
       form2.QrVentasT.ParamByName('Fecha2').AsDate:=encodedate(grid1.Ints[9,y],grid1.Ints[8,y],daysinamonth(grid1.Ints[9,y],grid1.Ints[8,y]));
       form2.QrVentasT.ParamByName('Usuario').AsString:=form2.QrGenerales.FieldByName('Usuario').AsString;
       form2.QrVentasT.Open;

       form2.qrventasL.close;
       form2.QrVentasL.ParamByName('DevVenta').AsString:=form2.QrGenerales.FieldByName('Devolucion').AsString;
       form2.QrVentasL.ParamByName('Tienda').AsString:=form2.QrGenerales.FieldByName('Factura').AsString;
       form2.QrVentasL.ParamByName('Fecha1').AsDate:=encodedate(grid1.Ints[9,y],grid1.Ints[8,y],1);
       form2.QrVentasL.ParamByName('Fecha2').AsDate:=encodedate(grid1.Ints[9,y],grid1.Ints[8,y],daysinamonth(grid1.Ints[9,y],grid1.Ints[8,y]));
       form2.QrVentasL.ParamByName('Usuario').AsString:=form2.QrGenerales.FieldByName('Usuario').AsString;
       form2.QrVentasL.Open;

       form2.qrventasL2.close;
       form2.QrVentasL2.ParamByName('DevVenta').AsString:=form2.QrGenerales.FieldByName('Devolucion').AsString;
       form2.QrVentasL2.ParamByName('Tienda').AsString:=form2.QrGenerales.FieldByName('Factura').AsString;
       form2.QrVentasL2.ParamByName('Fecha1').AsDate:=encodedate(grid1.Ints[9,y],grid1.Ints[8,y],1);
       form2.QrVentasL2.ParamByName('Fecha2').AsDate:=encodedate(grid1.Ints[9,y],grid1.Ints[8,y],daysinamonth(grid1.Ints[9,y],grid1.Ints[8,y]));
       form2.QrVentasL2.ParamByName('Usuario').AsString:=form2.QrGenerales.FieldByName('Usuario').AsString;
       form2.QrVentasL2.Open;

       if form2.qrVentasCUB.isempty=true then grid1.ints[2,y]:=0 else
       grid1.ints[2,y]:=form2.qrVentasCUB.FieldByName('Cantidad').AsInteger;

       if form2.qrVentasT.isempty=true then grid1.ints[3,y]:=0 else
       grid1.ints[3,y]:=form2.qrVentasT.FieldByName('Cantidad').AsInteger;

       if form2.qrVentasL.isempty=true then grid1.ints[6,y]:=0 else
       grid1.ints[6,y]:=form2.qrVentasL.FieldByName('Cantidad').AsInteger;

       if form2.qrVentasL2.isempty=true then grid1.ints[6,y]:=grid1.ints[6,y]+0 else
       grid1.ints[6,y]:=grid1.ints[6,y]+form2.qrVentasL2.FieldByName('Cantidad').AsInteger;


       //comienzan calculos

       if grid1.Floats[2,y]=0  then grid1.Floats[4,y]:=0 else
       grid1.floats[4,y]:=(grid1.Floats[3,y]/(grid1.Floats[2,y]/2))*100;

       grid1.Ints[5,y]:=grid1.Ints[3,y];

       if grid1.Floats[5,y]=0  then grid1.Floats[7,y]:=0 else
       grid1.floats[7,y]:=(grid1.Floats[6,y]/(grid1.Floats[5,y]/1))*100;
     end;

     grid1.HideColumns(8,9);
     Grid1.SortIndexes.Clear;
     Grid1.SortIndexes.Add(0);
     Grid1.SortIndexes.Add(8);
     Grid1.SortIndexes.Add(9);
     Grid1.QSortIndexed;

    {//comienzan graficos
      GridCT.ClearAll;
      gridCT.RowCount:=1;
      GridTL.ClearAll;
      gridTL.RowCount:=1;

    //Cub vs Tarjas
      for y := 1 to grid1.RowCount-1 do
        begin
          gridCT.Cells[0,y-1]:=grid1.Cells[0,y];
          gridCT.Floats[1,y-1]:=grid1.Floats[3,y];
          gridCT.RowCount:=gridCT.RowCount+1;
        end;

        gridct.RowCount:=gridct.RowCount-1;

        gridct.Sort(1,sddescending);

        for x := 1 to 10 do
          gridct.Colors[x,0]:=clskyblue;

       for y := 1 to gridct.RowCount-1 do
        begin
           if gridct.Floats[1,0]=0 then prop:=0 else
           prop:=(gridct.Floats[1,y]*10)/gridct.Floats[1,0];
           if prop<1 then prop:=2;
           for x := 1 to  trunc (prop) do
              gridct.Colors[x,y]:=clskyblue;
            gridct.Floats[x-1,y]:=gridct.Floats[1,y];
            gridct.cells[1,y]:='';
        end;
            gridct.Floats[10,0]:=gridct.Floats[1,0];
            gridct.cells[1,0]:='';

        gridcT.InsertRows(0,1);
        gridcT.Cells[0,0]:='CUB VS TARJAS';
        gridcT.RowCount:=gridcT.RowCount+1;


       //Tarjas vs LLaves
       prop:=0;

       for y := 1 to grid1.RowCount-1 do
        begin
          gridTl.Cells[0,y-1]:=grid1.Cells[0,y];
          gridTl.Floats[1,y-1]:=grid1.Floats[6,y];
          gridTl.RowCount:=gridTl.RowCount+1;
        end;

        gridtl.RowCount:=gridtl.RowCount-1;

        gridtl.Sort(1,sddescending);

        for x := 1 to 10 do
          gridtl.Colors[x,0]:=clskyblue;

       for y := 1 to gridtl.RowCount-1 do
        begin
           if gridtl.Floats[1,0]=0 then prop:=0 else
           prop:=(gridtl.Floats[1,y]*10)/gridtl.Floats[1,0];
           if prop<1 then prop:=2;
           for x := 1 to  trunc (prop) do
              gridtl.Colors[x,y]:=clskyblue;
            gridtl.Floats[x-1,y]:=gridtl.Floats[1,y];
            gridtl.cells[1,y]:='';
        end;
            gridtl.Floats[10,0]:=gridtl.Floats[1,0];
            gridtl.cells[1,0]:='';

        gridTL.InsertRows(0,1);
        gridTL.Cells[0,0]:='TARJAS VS LLAVES';
        gridTL.RowCount:=gridTL.RowCount+1;

for y := 1 to grid1.RowCount-1 do
   begin
     grid1.Cells[3,y]:=FormatFloat('##0.0%',grid1.Floats[3,y]);
     grid1.Cells[6,y]:=FormatFloat('##0.0%',grid1.Floats[6,y]);

   end;}

end;

procedure TForm6.SpeedButton2Click(Sender: TObject);
begin
Form6.Visible:=false;
form1.Visible:=true;
end;

procedure TForm6.SpeedButton3Click(Sender: TObject);
begin
save1.Execute;
Excel1.XLSExport(save1.FileName);

showmessage('Reporte Exportado con Exito!');

end;

end.
