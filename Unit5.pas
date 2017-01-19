unit Unit5;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, Grids, AdvObj, BaseGrid, AdvGrid, ComCtrls, Buttons,
  StdCtrls, tmsAdvGridExcel;

type
  TForm5 = class(TForm)
    Panel2: TPanel;
    Panel1: TPanel;
    grid1: TAdvStringGrid;
    Label1: TLabel;
    Label2: TLabel;
    SpeedButton1: TSpeedButton;
    Panel3: TPanel;
    Fecha1: TDateTimePicker;
    Fecha2: TDateTimePicker;
    Panel4: TPanel;
    SpeedButton2: TSpeedButton;
    SpeedButton3: TSpeedButton;
    Excel1: TAdvGridExcelIO;
    Save1: TSaveDialog;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    ScrollBox1: TScrollBox;
    Gridct: TAdvStringGrid;
    GridTL: TAdvStringGrid;
    procedure SpeedButton1Click(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
    procedure SpeedButton3Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form5: TForm5;

implementation

{$R *.dfm}

uses unit1,unit2;

procedure TForm5.SpeedButton1Click(Sender: TObject);
var
x,y: integer;
prop : single;
begin
grid1.ClearRows(1,grid1.RowCount-1);
grid1.RowCount:=2;
grid1.ColCount:=7;

grid1.loadfromfile('configCOM.cx2');
y:=1;

form1.qrmetas.Close;
form1.qrmetas.ParamByName('tienda').AsString:=form1.combobox1.Text;
form1.qrmetas.Open;

while not form1.qrmetas.Eof do
  begin
    grid1.cells[0,y]:=form1.qrmetas.FieldByName('nombre').AsString;
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
       form2.qrGenerales.Close;
       form2.qrGenerales.ParamByName('nombre').AsString:=grid1.Cells[0,y];
       form2.qrGenerales.Open;

       form2.qrventasCUB.close;
       form2.QrVentasCUB.ParamByName('DevVenta').AsString:=form2.QrGenerales.FieldByName('Devolucion').AsString;
       form2.QrVentasCUB.ParamByName('Tienda').AsString:=form2.QrGenerales.FieldByName('Factura').AsString;
       form2.QrVentasCUB.ParamByName('Fecha1').AsDate:=Fecha1.Date;
       form2.QrVentasCUB.ParamByName('Fecha2').AsDate:=Fecha2.date;
       form2.QrVentasCUB.ParamByName('Usuario').AsString:=form2.QrGenerales.FieldByName('Usuario').AsString;
       form2.QrVentasCUB.Open;

       form2.qrventasT.close;
       form2.QrVentasT.ParamByName('DevVenta').AsString:=form2.QrGenerales.FieldByName('Devolucion').AsString;
       form2.QrVentasT.ParamByName('Tienda').AsString:=form2.QrGenerales.FieldByName('Factura').AsString;
       form2.QrVentasT.ParamByName('Fecha1').AsDate:=Fecha1.Date;
       form2.QrVentasT.ParamByName('Fecha2').AsDate:=Fecha2.date;
       form2.QrVentasT.ParamByName('Usuario').AsString:=form2.QrGenerales.FieldByName('Usuario').AsString;
       form2.QrVentasT.Open;

       form2.qrventasL.close;
       form2.QrVentasL.ParamByName('DevVenta').AsString:=form2.QrGenerales.FieldByName('Devolucion').AsString;
       form2.QrVentasL.ParamByName('Tienda').AsString:=form2.QrGenerales.FieldByName('Factura').AsString;
       form2.QrVentasL.ParamByName('Fecha1').AsDate:=Fecha1.Date;
       form2.QrVentasL.ParamByName('Fecha2').AsDate:=Fecha2.date;
       form2.QrVentasL.ParamByName('Usuario').AsString:=form2.QrGenerales.FieldByName('Usuario').AsString;
       form2.QrVentasL.Open;

       form2.qrventasL2.close;
       form2.QrVentasL2.ParamByName('DevVenta').AsString:=form2.QrGenerales.FieldByName('Devolucion').AsString;
       form2.QrVentasL2.ParamByName('Tienda').AsString:=form2.QrGenerales.FieldByName('Factura').AsString;
       form2.QrVentasL2.ParamByName('Fecha1').AsDate:=Fecha1.Date;
       form2.QrVentasL2.ParamByName('Fecha2').AsDate:=Fecha2.date;
       form2.QrVentasL2.ParamByName('Usuario').AsString:=form2.QrGenerales.FieldByName('Usuario').AsString;
       form2.QrVentasL2.Open;

       if form2.qrVentasCUB.isempty=true then grid1.ints[1,y]:=0 else
       grid1.ints[1,y]:=form2.qrVentasCUB.FieldByName('Cantidad').AsInteger;

       if form2.qrVentasT.isempty=true then grid1.ints[2,y]:=0 else
       grid1.ints[2,y]:=form2.qrVentasT.FieldByName('Cantidad').AsInteger;

       if form2.qrVentasL.isempty=true then grid1.ints[5,y]:=0 else
       grid1.ints[5,y]:=form2.qrVentasL.FieldByName('Cantidad').AsInteger;

       if form2.qrVentasL2.isempty=true then grid1.ints[5,y]:=grid1.ints[5,y]+0 else
       grid1.ints[5,y]:=grid1.ints[5,y]+form2.qrVentasL2.FieldByName('Cantidad').AsInteger;


       //comienzan calculos

       if grid1.Floats[1,y]=0  then grid1.Floats[3,y]:=0 else
       grid1.floats[3,y]:=(grid1.Floats[2,y]/(grid1.Floats[1,y]/2))*100;

       grid1.Ints[4,y]:=grid1.Ints[2,y];

       if grid1.Floats[4,y]=0  then grid1.Floats[6,y]:=0 else
       grid1.floats[6,y]:=(grid1.Floats[5,y]/(grid1.Floats[4,y]/1))*100;
     end;

    //comienzan graficos
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

   end;

end;

procedure TForm5.SpeedButton2Click(Sender: TObject);
begin
Form5.Visible:=false;
form1.Visible:=true;
end;

procedure TForm5.SpeedButton3Click(Sender: TObject);
begin
save1.Execute;
Excel1.XLSExport(save1.FileName);

excel1.AdvStringGrid:=gridTL;
excel1.Options.ExportOverwrite := omNever;
excel1.GridStartCol:=0;
Excel1.XLSExport(save1.FileName,'Graficos',-1,1,InsertInSheet_InsertRows);

excel1.AdvStringGrid:=gridCT;
excel1.Options.ExportOverwrite := omNever;
excel1.GridStartCol:=0;
Excel1.XLSExport(save1.FileName,'Graficos',5,1,InsertInSheet_InsertRows);


showmessage('Reporte Exportado con Exito!');

end;

end.
