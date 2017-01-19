unit Unit4;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, AdvObj, BaseGrid, AdvGrid, ExtCtrls, StdCtrls, Buttons,
  ComCtrls, tmsAdvGridExcel;

type
  TForm4 = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    Grid1: TAdvStringGrid;
    Fecha1: TDateTimePicker;
    Fecha2: TDateTimePicker;
    SpeedButton1: TSpeedButton;
    Label1: TLabel;
    Label2: TLabel;
    Panel3: TPanel;
    Panel4: TPanel;
    SpeedButton2: TSpeedButton;
    SpeedButton3: TSpeedButton;
    Excel1: TAdvGridExcelIO;
    Save1: TSaveDialog;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    ScrollBox1: TScrollBox;
    GridC: TAdvStringGrid;
    procedure SpeedButton1Click(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
    procedure SpeedButton3Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form4: TForm4;

implementation

{$R *.dfm}

uses unit1,unit2;

procedure TForm4.SpeedButton1Click(Sender: TObject);
var
x,y : integer;
prop : single;
begin

grid1.ClearRows(1,grid1.RowCount-1);
grid1.RowCount:=2;
grid1.ColCount:=4;
grid1.loadfromfile('configCan.cx2');

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

       form2.qrventasOT.close;
       form2.QrVentasOT.ParamByName('DevVenta').AsString:=form2.QrGenerales.FieldByName('Devolucion').AsString;
       form2.QrVentasOT.ParamByName('Tienda').AsString:=form2.QrGenerales.FieldByName('Factura').AsString;
       form2.QrVentasOT.ParamByName('Fecha1').AsDate:=Fecha1.Date;
       form2.QrVentasOT.ParamByName('Fecha2').AsDate:=Fecha2.date;
       form2.QrVentasOT.ParamByName('Usuario').AsString:=form2.QrGenerales.FieldByName('Usuario').AsString;
       form2.QrVentasOT.Open;

       form2.qrventasCBHD.close;
       form2.QrVentasCBHD.ParamByName('DevVenta').AsString:=form2.QrGenerales.FieldByName('Devolucion').AsString;
       form2.QrVentasCBHD.ParamByName('Tienda').AsString:=form2.QrGenerales.FieldByName('Factura').AsString;
       form2.QrVentasCBHD.ParamByName('Fecha1').AsDate:=Fecha1.Date;
       form2.QrVentasCBHD.ParamByName('Fecha2').AsDate:=Fecha2.date;
       form2.QrVentasCBHD.ParamByName('Usuario').AsString:=form2.QrGenerales.FieldByName('Usuario').AsString;
       form2.QrVentasCBHD.Open;

       form2.qrventasE.close;
       form2.QrVentasE.ParamByName('DevVenta').AsString:=form2.QrGenerales.FieldByName('Devolucion').AsString;
       form2.QrVentasE.ParamByName('Tienda').AsString:=form2.QrGenerales.FieldByName('Factura').AsString;
       form2.QrVentasE.ParamByName('Fecha1').AsDate:=Fecha1.Date;
       form2.QrVentasE.ParamByName('Fecha2').AsDate:=Fecha2.date;
       form2.QrVentasE.ParamByName('Usuario').AsString:=form2.QrGenerales.FieldByName('Usuario').AsString;
       form2.QrVentasE.Open;

       if form2.qrVentasOT.isempty=true then grid1.ints[1,y]:=0 else
       grid1.ints[1,y]:=form2.qrVentasOT.FieldByName('Cantidad').AsInteger;

       if form2.qrVentasCBHD.isempty=true then grid1.ints[2,y]:=0 else
       grid1.ints[2,y]:=form2.qrVentasCBHD.FieldByName('Cantidad').AsInteger;

       if form2.qrVentasE.isempty=true then grid1.ints[2,y]:=grid1.ints[2,y]+0 else
       grid1.ints[2,y]:=grid1.ints[2,y]+form2.qrVentasE.FieldByName('Cantidad').AsInteger;

       // comienzan calculos
       if grid1.Ints[1,y]=0 then grid1.Floats[3,y]:=0 else
       grid1.Floats[3,y]:=(grid1.Floats[2,y]/(grid1.Floats[1,y]+grid1.Floats[2,y]))*100;

     end;

     //comienzan graficas


      GridC.ClearAll;
      gridc.RowCount:=1;

    //Canibal B
      for y := 1 to grid1.RowCount-1 do
        begin
          gridC.Cells[0,y-1]:=grid1.Cells[0,y];
          gridC.Floats[1,y-1]:=grid1.Floats[3,y];
          gridC.RowCount:=gridC.RowCount+1;
        end;

        gridc.RowCount:=gridc.RowCount-1;

        gridc.Sort(1,sddescending);

        for x := 1 to 10 do
          gridc.Colors[x,0]:=clskyblue;

       for y := 1 to gridc.RowCount-1 do
        begin
           if gridc.Floats[1,0]=0 then prop:=0 else
           prop:=(gridc.Floats[1,y]*10)/gridc.Floats[1,0];
           if prop<1 then prop:=2;
           for x := 1 to  trunc (prop) do
              gridc.Colors[x,y]:=clskyblue;
            gridc.Floats[x-1,y]:=gridc.Floats[1,y];
            gridc.cells[1,y]:='';
        end;
            gridc.Floats[10,0]:=gridc.Floats[1,0];
            gridc.cells[1,0]:='';

            gridc.InsertRows(0,1);
            gridc.Cells[0,0]:='CANIBALIZACION POR AGENTE';
            gridc.RowCount:=gridc.RowCount+1;

       //Canibal E

     for y := 1 to grid1.RowCount-1 do
   begin
     grid1.Cells[3,y]:=FormatFloat('##0.0%',grid1.Floats[3,y]);
   end;


end;

procedure TForm4.SpeedButton2Click(Sender: TObject);
begin
Form4.Visible:=false;
form1.Visible:=true;
end;

procedure TForm4.SpeedButton3Click(Sender: TObject);
begin
save1.Execute;
Excel1.XLSExport(save1.FileName,'Datos');

excel1.AdvStringGrid:=gridC;
excel1.Options.ExportOverwrite := omNever;
excel1.GridStartCol:=0;
Excel1.XLSExport(save1.FileName,'Graficos',-1,1,InsertInSheet_InsertRows);


showmessage('Reporte Exportado con Exito!');
end;

end.
