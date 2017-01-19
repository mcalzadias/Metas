unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Buttons, StdCtrls, ExtCtrls, DB, MemDS, DBAccess, MSAccess, Grids,
  AdvObj, BaseGrid, AdvGrid;

type
  TForm1 = class(TForm)
    Panel1: TPanel;
    Edit1: TEdit;
    Edit2: TEdit;
    Label1: TLabel;
    Label2: TLabel;
    SpeedButton1: TSpeedButton;
    Panel2: TPanel;
    ComboBox1: TComboBox;
    Label3: TLabel;
    Panel3: TPanel;
    BtGestion: TSpeedButton;
    BtVtasMetas: TSpeedButton;
    Panel4: TPanel;
    SpeedButton4: TSpeedButton;
    Genesis: TMSConnection;
    QrLogin: TMSQuery;
    QrMetas: TMSQuery;
    Intelisis: TMSConnection;
    BtCanibal: TSpeedButton;
    BtCompara: TSpeedButton;
    bt3Metas: TSpeedButton;
    bt3canibal: TSpeedButton;
    bt3comp: TSpeedButton;
    BtCruzada: TSpeedButton;
    Bt3cruzado: TSpeedButton;
    procedure SpeedButton4Click(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure BtGestionClick(Sender: TObject);
    procedure Grid1CanEditCell(Sender: TObject; ARow, ACol: Integer;
      var CanEdit: Boolean);
    procedure ComboBox1Change(Sender: TObject);
    procedure BtVtasMetasClick(Sender: TObject);
    procedure BtCanibalClick(Sender: TObject);
    procedure BtComparaClick(Sender: TObject);
    procedure bt3compClick(Sender: TObject);
    procedure bt3canibalClick(Sender: TObject);
    procedure bt3MetasClick(Sender: TObject);
    procedure BtCruzadaClick(Sender: TObject);
    procedure Bt3cruzadoClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
  Menux: Integer;

implementation

{$R *.dfm}

uses unit2,unit3,unit4,unit5,unit6,unit7,unit8,unit9,unit10;

procedure TForm1.BtComparaClick(Sender: TObject);
begin
form5.visible:=true;
form1.Visible:=false;
form5.Fecha1.Date:=date;
form5.Fecha2.Date:=date;
end;

procedure TForm1.BtGestionClick(Sender: TObject);
var
y,x: integer;
begin


form3.grid1.ClearRows(1,Form3.grid1.RowCount-1);
form3.grid1.RowCount:=2;
y:=1;

qrmetas.Close;
qrmetas.ParamByName('tienda').AsString:=combobox1.Text;
qrmetas.Open;

while not qrmetas.Eof do
  begin
    form3.grid1.cells[0,y]:=qrmetas.FieldByName('nombre').AsString;
    form3.grid1.ints[1,y]:=qrmetas.FieldByName('MetaUO').AsInteger;
    form3.grid1.ints[2,y]:=qrmetas.FieldByName('MetaUE').AsInteger;
    form3.grid1.ints[3,y]:=qrmetas.FieldByName('MetaUEHD').AsInteger;
    form3.grid1.ints[4,y]:=qrmetas.FieldByName('MetaLL').AsInteger;
    form3.grid1.ints[5,y]:=qrmetas.FieldByName('MetaT').AsInteger;
    form3.grid1.ints[6,y]:=qrmetas.FieldByName('MetaC').AsInteger;
    form3.grid1.floats[7,y]:=qrmetas.FieldByName('MetaI').value;
    form3.grid1.ints[8,y]:=qrmetas.FieldByName('MetaS').AsInteger;
    form3.grid1.ints[9,y]:=qrmetas.FieldByName('MetaB').AsInteger;
    form3.grid1.ints[10,y]:=qrmetas.FieldByName('MetaPT').AsInteger;
    form3.grid1.ints[11,y]:=qrmetas.FieldByName('MetaPL').AsInteger;
    form3.grid1.ints[12,y]:=qrmetas.FieldByName('MetaQ').AsInteger;
    form3.grid1.ints[13,y]:=qrmetas.FieldByName('MetaQHD').AsInteger;
    form3.grid1.ints[14,y]:=qrmetas.FieldByName('MetaTotal').AsInteger;
    form3.grid1.RowCount:=form3.grid1.RowCount+1;
    inc(y);
    qrmetas.Next;
  end;
  form3.grid1.RowCount:=form3.grid1.RowCount-1;

  //formato de celdas
  for y := 1 to form3.grid1.RowCount-1 do
    for x := 1 to form3.grid1.ColCount-1 do
     begin
       form3.grid1.Alignments[x,y]:=tarightjustify;
     end;

 // for y := 1 to form3.grid1.RowCount-1 do
    // form3.grid1.Cells[13,y]:=FormatFloat('#,##0',form3.grid1.Floats[13,y]);

  form3.Visible:=true;
  form1.Visible:=false;

end;

procedure TForm1.BtVtasMetasClick(Sender: TObject);
begin
form2.visible:=true;
form1.Visible:=false;
form2.Fecha1.Date:=date;
form2.Fecha2.Date:=date;
end;

procedure TForm1.ComboBox1Change(Sender: TObject);
begin
btVtasMetas.Enabled:=true;
btCanibal.Enabled:=true;
btCompara.Enabled:=true;

bt3Metas.Enabled:=true;
bt3Canibal.Enabled:=true;
bt3Comp.Enabled:=true;
btcruzada.Enabled:=true;
bt3cruzado.Enabled:=true;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
combobox1.Text:='';
edit1.Text:='';
edit2.Text:='';
end;

procedure TForm1.Grid1CanEditCell(Sender: TObject; ARow, ACol: Integer;
  var CanEdit: Boolean);
begin
if acol=13 then canedit:=false;

end;

procedure TForm1.SpeedButton1Click(Sender: TObject);
begin

Menux:=0;
qrLogin.Close;
qrLogin.ParamByName('usuario').AsString:=edit1.Text;
qrLogin.ParamByName('password').AsString:=edit2.Text;
qrLogin.Open;
if qrLogin.IsEmpty=false then
  begin
    if (edit1.text='mcalzadias') or (edit1.text='egarcia') or (edit1.text='cjaime') or (edit1.text='mrivera') or (edit1.text='gtejeda') or (edit1.text='mrojas') then
      begin
        menuX:=1;
        if edit1.text='egarcia' then
           begin
             combobox1.Items.Clear;
             ComboBox1.Items.Add('MK GDL');
             btGestion.Enabled:=true;
           end;
        if edit1.text='cjaime' then
           begin
             combobox1.Items.Clear;
             ComboBox1.Items.Add('MK MTY');
             btGestion.Enabled:=true;
           end;
        if edit1.text='mrivera' then
           begin
             combobox1.Items.Clear;
             ComboBox1.Items.Add('CDMX');
             btGestion.Enabled:=true;
           end;
        if edit1.text='gtejeda' then
           begin
             combobox1.Items.Clear;
             ComboBox1.Items.Add('MK Leon');
             btGestion.Enabled:=true;
           end;
        if (edit1.text='mcalzadias') or (edit1.text='mrojas') then
           begin
             combobox1.Items.Clear;
             ComboBox1.Items.Add('MK GDL');
             ComboBox1.Items.Add('MK MTY');
             ComboBox1.Items.Add('MK Leon');
             ComboBox1.Items.Add('CDMX');
             btGestion.Enabled:=true;
           end;

      end;

    if menux=0 then
     begin
       combobox1.Items.Clear;
       ComboBox1.Items.Add(QrLogin.FieldByName('Tienda').AsString);
       btGestion.Enabled:=false;
     end;

     combobox1.Enabled:=true;



  end;


end;

procedure TForm1.Bt3cruzadoClick(Sender: TObject);
begin
form10.visible:=true;
form1.Visible:=false;
end;

procedure TForm1.BtCruzadaClick(Sender: TObject);
begin
form9.visible:=true;
form1.Visible:=false;
form9.Fecha1.Date:=date;
form9.Fecha2.Date:=date;
end;

procedure TForm1.bt3canibalClick(Sender: TObject);
begin
form7.visible:=true;
form1.Visible:=false;
end;

procedure TForm1.bt3compClick(Sender: TObject);
begin
form6.visible:=true;
form1.Visible:=false;
end;

procedure TForm1.bt3MetasClick(Sender: TObject);
begin
form8.visible:=true;
form1.Visible:=false;
end;

procedure TForm1.BtCanibalClick(Sender: TObject);
begin
form4.visible:=true;
form1.Visible:=false;
form4.Fecha1.Date:=date;
form4.Fecha2.Date:=date;
end;

procedure TForm1.SpeedButton4Click(Sender: TObject);
begin
Application.Terminate;
end;

end.
