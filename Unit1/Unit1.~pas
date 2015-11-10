unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, sSkinProvider, sSkinManager, ExtCtrls, sPanel, jpeg, acImage,
  Menus, Buttons, sSpeedButton, StdCtrls, sGroupBox, Mask, sMaskEdit,
  sCustomComboEdit, sCurrEdit, sCurrencyEdit, sSpinEdit, sEdit, sLabel,
  Grids, DBGrids, sComboBox, sScrollBar, sScrollBox, DB, ADODB;

type
  TForm1 = class(TForm)
    sPanel1: TsPanel;
    sSkinManager1: TsSkinManager;
    sSkinProvider1: TsSkinProvider;
    sGroupBox1: TsGroupBox;
    sGroupBox2: TsGroupBox;
    sSpeedButton1: TsSpeedButton;
    sSpeedButton11: TsSpeedButton;
    sGroupBox3: TsGroupBox;
    sSpeedButton4: TsSpeedButton;
    sGroupBox4: TsGroupBox;
    sSpeedButton6: TsSpeedButton;
    MainMenu1: TMainMenu;
    P1: TMenuItem;
    A2: TMenuItem;
    R1: TMenuItem;
    V1: TMenuItem;
    V2: TMenuItem;
    F1: TMenuItem;
    S1: TMenuItem;
    P2: TMenuItem;
    F2: TMenuItem;
    A1: TMenuItem;
    A3: TMenuItem;
    A4: TMenuItem;
    sImage1: TsImage;
    sPanel2: TsPanel;
    sGroupBox5: TsGroupBox;
    sLabel1: TsLabel;
    sLabel2: TsLabel;
    sLabel3: TsLabel;
    sSpeedButton7: TsSpeedButton;
    sSpeedButton8: TsSpeedButton;
    sLabel14: TsLabel;
    sLabel15: TsLabel;
    sEdit1: TsEdit;
    sDecimalSpinEdit1: TsDecimalSpinEdit;
    sCurrencyEdit1: TsCurrencyEdit;
    sPanel3: TsPanel;
    sGroupBox6: TsGroupBox;
    sSpeedButton2: TsSpeedButton;
    sSpeedButton3: TsSpeedButton;
    sGroupBox7: TsGroupBox;
    sLabel4: TsLabel;
    sLabel5: TsLabel;
    sLabel6: TsLabel;
    sLabel7: TsLabel;
    sSpeedButton9: TsSpeedButton;
    sSpeedButton10: TsSpeedButton;
    sLabel13: TsLabel;
    sDecimalSpinEdit2: TsDecimalSpinEdit;
    sCurrencyEdit2: TsCurrencyEdit;
    sCurrencyEdit3: TsCurrencyEdit;
    sDecimalSpinEdit10: TsDecimalSpinEdit;
    DBGrid1: TDBGrid;
    sPanel4: TsPanel;
    sGroupBox8: TsGroupBox;
    sGroupBox9: TsGroupBox;
    sLabel11: TsLabel;
    sLabel12: TsLabel;
    sSpeedButton12: TsSpeedButton;
    sSpeedButton14: TsSpeedButton;
    sSpeedButton13: TsSpeedButton;
    sGroupBox10: TsGroupBox;
    sLabel16: TsLabel;
    sLabel17: TsLabel;
    sLabel18: TsLabel;
    sLabel19: TsLabel;
    sLabel20: TsLabel;
    sLabel21: TsLabel;
    sLabel22: TsLabel;
    sLabel23: TsLabel;
    sSpeedButton5: TsSpeedButton;
    sLabel24: TsLabel;
    sLabel25: TsLabel;
    sComboBox8: TsComboBox;
    sComboBox9: TsComboBox;
    sComboBox10: TsComboBox;
    sComboBox11: TsComboBox;
    sComboBox12: TsComboBox;
    sGroupBox11: TsGroupBox;
    sLabel26: TsLabel;
    sLabel27: TsLabel;
    sLabel28: TsLabel;
    sLabel29: TsLabel;
    sLabel30: TsLabel;
    sLabel31: TsLabel;
    sLabel32: TsLabel;
    sLabel33: TsLabel;
    sLabel34: TsLabel;
    sLabel35: TsLabel;
    sSpeedButton15: TsSpeedButton;
    sGroupBox12: TsGroupBox;
    sEdit3: TsEdit;
    sScrollBox1: TsScrollBox;
    sLabel9: TsLabel;
    sLabel8: TsLabel;
    sLabel10: TsLabel;
    sDecimalSpinEdit3: TsDecimalSpinEdit;
    sComboBox1: TsComboBox;
    sCurrencyEdit4: TsCurrencyEdit;
    sDecimalSpinEdit4: TsDecimalSpinEdit;
    sComboBox2: TsComboBox;
    sCurrencyEdit5: TsCurrencyEdit;
    sDecimalSpinEdit5: TsDecimalSpinEdit;
    sComboBox3: TsComboBox;
    sCurrencyEdit6: TsCurrencyEdit;
    sDecimalSpinEdit6: TsDecimalSpinEdit;
    sComboBox4: TsComboBox;
    sCurrencyEdit8: TsCurrencyEdit;
    sDecimalSpinEdit7: TsDecimalSpinEdit;
    sComboBox5: TsComboBox;
    sCurrencyEdit9: TsCurrencyEdit;
    sDecimalSpinEdit8: TsDecimalSpinEdit;
    sComboBox6: TsComboBox;
    sCurrencyEdit10: TsCurrencyEdit;
    sDecimalSpinEdit9: TsDecimalSpinEdit;
    sComboBox7: TsComboBox;
    sCurrencyEdit7: TsCurrencyEdit;
    sDecimalSpinEdit11: TsDecimalSpinEdit;
    sComboBox13: TsComboBox;
    sCurrencyEdit11: TsCurrencyEdit;
    sDecimalSpinEdit12: TsDecimalSpinEdit;
    sComboBox14: TsComboBox;
    sCurrencyEdit12: TsCurrencyEdit;
    sDecimalSpinEdit13: TsDecimalSpinEdit;
    sComboBox15: TsComboBox;
    sCurrencyEdit13: TsCurrencyEdit;
    sDecimalSpinEdit14: TsDecimalSpinEdit;
    sComboBox16: TsComboBox;
    sCurrencyEdit14: TsCurrencyEdit;
    sDecimalSpinEdit15: TsDecimalSpinEdit;
    sComboBox17: TsComboBox;
    sCurrencyEdit15: TsCurrencyEdit;
    sDecimalSpinEdit16: TsDecimalSpinEdit;
    sComboBox18: TsComboBox;
    sCurrencyEdit16: TsCurrencyEdit;
    sDecimalSpinEdit17: TsDecimalSpinEdit;
    sComboBox19: TsComboBox;
    sCurrencyEdit17: TsCurrencyEdit;
    sDecimalSpinEdit18: TsDecimalSpinEdit;
    sComboBox20: TsComboBox;
    sCurrencyEdit18: TsCurrencyEdit;
    sDecimalSpinEdit19: TsDecimalSpinEdit;
    sComboBox21: TsComboBox;
    sCurrencyEdit19: TsCurrencyEdit;
    sDecimalSpinEdit20: TsDecimalSpinEdit;
    sCurrencyEdit20: TsCurrencyEdit;
    sDecimalSpinEdit21: TsDecimalSpinEdit;
    sComboBox22: TsComboBox;
    sComboBox23: TsComboBox;
    sCurrencyEdit21: TsCurrencyEdit;
    sDecimalSpinEdit22: TsDecimalSpinEdit;
    sComboBox24: TsComboBox;
    sCurrencyEdit22: TsCurrencyEdit;
    sDecimalSpinEdit23: TsDecimalSpinEdit;
    sComboBox25: TsComboBox;
    sCurrencyEdit23: TsCurrencyEdit;
    sDecimalSpinEdit24: TsDecimalSpinEdit;
    sComboBox26: TsComboBox;
    sCurrencyEdit24: TsCurrencyEdit;
    sDecimalSpinEdit25: TsDecimalSpinEdit;
    sComboBox27: TsComboBox;
    sCurrencyEdit25: TsCurrencyEdit;
    sDecimalSpinEdit26: TsDecimalSpinEdit;
    sComboBox28: TsComboBox;
    sCurrencyEdit26: TsCurrencyEdit;
    sDecimalSpinEdit27: TsDecimalSpinEdit;
    sComboBox29: TsComboBox;
    sCurrencyEdit27: TsCurrencyEdit;
    sDecimalSpinEdit28: TsDecimalSpinEdit;
    sComboBox30: TsComboBox;
    sCurrencyEdit28: TsCurrencyEdit;
    sDecimalSpinEdit29: TsDecimalSpinEdit;
    sComboBox31: TsComboBox;
    sCurrencyEdit29: TsCurrencyEdit;
    sDecimalSpinEdit30: TsDecimalSpinEdit;
    sComboBox32: TsComboBox;
    sCurrencyEdit30: TsCurrencyEdit;
    sDecimalSpinEdit31: TsDecimalSpinEdit;
    sComboBox33: TsComboBox;
    sCurrencyEdit31: TsCurrencyEdit;
    sDecimalSpinEdit32: TsDecimalSpinEdit;
    sComboBox34: TsComboBox;
    sCurrencyEdit32: TsCurrencyEdit;
    sDecimalSpinEdit33: TsDecimalSpinEdit;
    sComboBox35: TsComboBox;
    sCurrencyEdit33: TsCurrencyEdit;
    sDecimalSpinEdit34: TsDecimalSpinEdit;
    sComboBox36: TsComboBox;
    sCurrencyEdit34: TsCurrencyEdit;
    sDecimalSpinEdit35: TsDecimalSpinEdit;
    sComboBox37: TsComboBox;
    sCurrencyEdit35: TsCurrencyEdit;
    sEdit4: TsEdit;
    sEdit5: TsEdit;
    sEdit6: TsEdit;
    sEdit7: TsEdit;
    ADOConnection1: TADOConnection;
    DataSource1: TDataSource;
    ADOTable1: TADOTable;
    sComboBox38: TsComboBox;
    procedure sSpeedButton7Click(Sender: TObject);
    procedure sSpeedButton1Click(Sender: TObject);
    procedure sSpeedButton8Click(Sender: TObject);
    procedure sSpeedButton10Click(Sender: TObject);
    procedure sSpeedButton9Click(Sender: TObject);
    procedure sSpeedButton2Click(Sender: TObject);
    procedure sSpeedButton3Click(Sender: TObject);
    procedure DBGrid1DrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure sSpeedButton11Click(Sender: TObject);
    procedure sSpeedButton4Click(Sender: TObject);
    procedure sSpeedButton12Click(Sender: TObject);
    procedure sSpeedButton13Click(Sender: TObject);
    procedure sSpeedButton14Click(Sender: TObject);
   procedure sLabel11Click(Sender: TObject);
 procedure sLabel12Click(Sender: TObject);
    procedure sComboBox1Click(Sender: TObject);
    procedure sComboBox2Click(Sender: TObject);
    procedure sComboBox3Click(Sender: TObject);
    procedure sComboBox4Click(Sender: TObject);
    procedure sComboBox5Click(Sender: TObject);
    procedure sComboBox6Click(Sender: TObject);
    procedure sComboBox8Click(Sender: TObject);
    procedure sSpeedButton5Click(Sender: TObject);
    procedure P2Click(Sender: TObject);
    procedure sSpeedButton15Click(Sender: TObject);
    procedure F2Click(Sender: TObject);
    procedure sSpeedButton6Click(Sender: TObject);
    function EnLettres(N:integer):string;
    procedure A4Click(Sender: TObject);
    procedure sComboBox13Click(Sender: TObject);
    procedure sComboBox14Click(Sender: TObject);
    procedure sComboBox15Click(Sender: TObject);
    procedure sComboBox16Click(Sender: TObject);
    procedure sComboBox17Click(Sender: TObject);
    procedure sComboBox18Click(Sender: TObject);
    procedure sComboBox19Click(Sender: TObject);
    procedure sComboBox20Click(Sender: TObject);
    procedure sComboBox21Click(Sender: TObject);
    procedure sComboBox22Click(Sender: TObject);
    procedure sComboBox23Click(Sender: TObject);
    procedure sComboBox24Click(Sender: TObject);
    procedure sComboBox25Click(Sender: TObject);
    procedure sComboBox26Click(Sender: TObject);
    procedure sComboBox27Click(Sender: TObject);
    procedure sComboBox28Click(Sender: TObject);
    procedure sComboBox29Click(Sender: TObject);
    procedure sComboBox30Click(Sender: TObject);
    procedure sComboBox31Click(Sender: TObject);
    procedure sComboBox32Click(Sender: TObject);
    procedure sComboBox33Click(Sender: TObject);
    procedure sComboBox34Click(Sender: TObject);
    procedure sComboBox35Click(Sender: TObject);
    procedure sComboBox36Click(Sender: TObject);
    procedure sComboBox37Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormMouseWheelDown(Sender: TObject; Shift: TShiftState;
      MousePos: TPoint; var Handled: Boolean);
    procedure FormMouseWheelUp(Sender: TObject; Shift: TShiftState;
      MousePos: TPoint; var Handled: Boolean);

  private
 p_u1,p_u2,p_u3,p_u4,p_u5,p_u6,p_u7,p_u8,p_u9,p_u10,p_u11,p_u12,p_u13,p_u14,p_u15,p_u16,p_u17,p_u18,p_u19,p_u20,p_u21,p_u22,p_u23,p_u24,p_u25,p_u26,p_u27,p_u28,p_u29,p_u30,p_u31,p_u32:integer;
      OldWindowProc: TWndMethod;
          Procedure DBGridNewWindowProc(var Msg: TMessage);

  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

uses Unit2, DateUtils, ComObj, Unit3;

{$R *.dfm}

procedure TForm1.sSpeedButton7Click(Sender: TObject);
begin

   DataModule2.ADOTable1.close;
   DataModule2.ADOTable1.Open;
   if (sEdit1.Text='Nom') or (sCurrencyEdit1.Text='0') or (sDecimalSpinEdit1.Text='0') then
   ShowMessage('Erreur ! Vous devez remplir les champs nécessaires')
   else
   begin
  With datamodule2.ADOTable1 Do
  begin

// Mettre la table sur nouveau enregistrement
   Append;
// Donner la valeur de chaque champs ici on peux utiliser soit Fields[...] soit FieldsByName(...)



        datamodule2.ADOTable1.FieldByName('nom_produit').AsString:=sEdit1.Text;
        datamodule2.ADOTable1.FieldByName('prix_produit').AsString:=sCurrencyEdit1.Text;
        datamodule2.ADOTable1.FieldByName('quantite_produit').AsString:=sDecimalSpinEdit1.Text;
        datamodule2.ADOTable1.FieldByName('date_ajout').AsDateTime:=date();
        datamodule2.ADOTable1.FieldByName('date_modif').AsDateTime:=date();


// Valider l'enregistrement
   Post;
   DataModule2.ADOTable1.Close;
   sEdit1.Text:='Nom';
   sCurrencyEdit1.Text:='0';
   sDecimalSpinEdit1.Text:='0';
   ShowMessage('Le produit est bien ajouté');

  end;
end;


end;



procedure TForm1.sSpeedButton1Click(Sender: TObject);
begin
spanel4.Visible:=false;
sGroupBox8.Visible:=false;
spanel3.Visible:=false;
sGroupBox6.Visible:=false;
sGroupBox10.Visible:=false;
sPanel2.Visible:=true;
sGroupBox5.Visible:=true;
end;

procedure TForm1.sSpeedButton8Click(Sender: TObject);
begin
if MessageDlg('Etes-vous sûr de vouloir annuler L ajout des produits ?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
begin
sPanel2.Visible:=false;
sGroupBox5.Visible:=false;
end;
end;

procedure TForm1.sSpeedButton9Click(Sender: TObject);
begin
DBGrid1.Visible:=false;
ADOTable1.Filtered:=false;
if (sComboBox38.Text='Produit') or (sComboBox38.Text='') then DBGrid1.Visible:=true
else
begin
ADOTable1.Filter:='nom_produit='+QuotedStr(sComboBox38.Text);
ADOTable1.Filtered:=true;
DBGrid1.Visible:=true;
end;


end;


procedure TForm1.sSpeedButton10Click(Sender: TObject);
begin
if MessageDlg('Etes-vous sûr de vouloir annuler La recherche des produits ?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
begin
sPanel3.Visible:=false;
DBGrid1.Visible:=false;
sGroupBox6.Visible:=false;
end;
end;
procedure TForm1.sSpeedButton2Click(Sender: TObject);
 var Modifier: TForm;
    quantite: TsDecimalSpinEdit;
    prix: TsCurrencyEdit;


begin
if (dbgrid1.Visible=false) then
begin
ShowMessage('Erreur vous devez lancer la recherche et sélectionner un produit pour le modifier');
end
else
begin
 prix:=TsCurrencyEdit.Create(Nil);
 quantite:=TsDecimalSpinEdit.Create(Nil);



 // Créer le message modifier l'enregistrement (#13= Sauter la ligne entrer)

 modifier := CreateMessageDialog('Modifier  Un produit                                   '+#13+#13+
                                 'Quantité                                                         '+#13+#13+#13+
                                'Prix                                              '+#13+#13+#13+#13
                                ,mtInformation,[mbYes, mbCancel]);

 with Modifier do
 try
 width :=0;
 height :=900;
 autoscroll:=false;
 autosize:=true;



 // Modifier la position de la zone de l'adresse
 quantite.Parent:=modifier;
 quantite.Left:=55;
 quantite.Top:=50;
 quantite.Width:=150;
 quantite.Height:=30;

 // Modifier la position de la zone de numero de telephone
 prix.Parent:=modifier;
 prix.Left:=55;
 prix.Top:=100;
 prix.Width:=150;
 prix.Height:=30;


// Remplir les valeurs de chaque zone de texte
quantite.Text:= ADOTable1.Fields[3].AsString;
prix.Text:= ADOTable1.Fields[2].AsString;




 if (ShowModal = ID_YES) then
Begin
With ADOTable1 Do
begin
// Mettre la table mode d'edition = modification
   Edit;
// Donner la valeur de chaque champs ici on peux utiliser soit Fields[...] soit FieldsByName(...)
   Fields[3].Value:=quantite.Text;
   Fields[2].Value:=prix.Text;
   Fields[5].Value:=date();

// Valider l'enregistrement
   Post;

end;
end;
 finally
 // Libérer les compos crées ainsi que la fiche modifier
    quantite.Free;
     prix.Free;
    modifier.Free;


end;
end;
end;

procedure TForm1.sSpeedButton3Click(Sender: TObject);
begin
if (dbgrid1.Visible=false) then
begin
ShowMessage('Erreur vous devez lancer la recherche');
end
else
begin

if MessageDlg('Etes-vous sûr de vouloir supprimer ce produit ', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
  datamodule2.ADOTable1.Delete;
end;
end;

procedure TForm1.DBGrid1DrawColumnCell(Sender: TObject; const Rect: TRect;
  DataCol: Integer; Column: TColumn; State: TGridDrawState);
    begin

        if ADOTable1.FieldByName('nom_produit').Asstring ='' then
        begin
        DBGrid1.Canvas.Brush.Color:= cllime;
        DBGrid1.DefaultDrawColumnCell(Rect, DataCol, Column, State);
        end;
        if (ADOTable1.FieldByName('nom_produit').Asstring=sComboBox38.Text) and  (sDecimalSpinEdit2.Text='0') and (sCurrencyEdit2.Text='0') and (sCurrencyEdit3.Text='0') and (sDecimalSpinEdit10.Text='0') then
        begin

        DBGrid1.Canvas.Brush.Color:= Clred;
        DBGrid1.DefaultDrawColumnCell(Rect, DataCol, Column, State);
        end;

        if (sComboBox38.Text='produit') and (ADOTable1.FieldByName('quantite_produit').AsInteger>=strtoint(sDecimalSpinEdit2.text)) and (ADOTable1.FieldByName('quantite_produit').AsInteger<=strtoint(sDecimalSpinEdit10.text))   then
        begin
        DBGrid1.Canvas.Brush.Color:= Clred;
        DBGrid1.DefaultDrawColumnCell(Rect, DataCol, Column, State);
        end;

        if (sComboBox38.Text='produit') and  (sDecimalSpinEdit2.text='0') and (sDecimalSpinEdit10.text='0')and (ADOTable1.FieldByName('Prix_produit').Asinteger >= strtoint(sCurrencyEdit2.Text)) and (ADOTable1.FieldByName('Prix_produit').Asinteger<=strtoint(scurrencyedit3.Text))   then
        begin
        DBGrid1.Canvas.Brush.Color:= Clred;
        DBGrid1.DefaultDrawColumnCell(Rect, DataCol, Column, State);
        end;

        if (ADOTable1.FieldByName('nom_produit').Asstring=sComboBox38.Text) and  (ADOTable1.FieldByName('quantite_produit').Asstring=sdecimalspinedit2.Text) and (ADOTable1.FieldByName('Prix_produit').Asinteger >= strtoint(sCurrencyEdit2.Text)) and (ADOTable1.FieldByName('Prix_produit').Asinteger<=strtoint(scurrencyedit3.Text))   then
        begin
        DBGrid1.Canvas.Brush.Color:= Clred;
        DBGrid1.DefaultDrawColumnCell(Rect, DataCol, Column, State);
        end;

    end;


procedure TForm1.sSpeedButton11Click(Sender: TObject);
begin
sPanel4.Visible:=false;
sGroupBox8.Visible:=false;
sPanel2.Visible:=false;
sGroupBox5.Visible:=false;

sGroupBox10.Visible:=false;
//sGroupBox11.Visible:=false;


sComboBox38.Items.BeginUpdate;
sComboBox38.Items.Clear;


while not (ADOTable1.Eof) do
begin
sComboBox38.Items.Add(ADOTable1.FieldByName('nom_produit').AsString);
ADOTable1.Next;
end;
sComboBox38.Items.EndUpdate;


sPanel3.Visible:=true;

sGroupBox6.Visible:=true;


end;

procedure TForm1.sSpeedButton4Click(Sender: TObject);
var i,j : integer;
begin
sCurrencyEdit4.Text:='0';
sCurrencyEdit5.Text:='0';
sCurrencyEdit6.Text:='0';
sCurrencyEdit7.Text:='0';
sCurrencyEdit8.Text:='0';
sCurrencyEdit9.Text:='0';
sCurrencyEdit10.Text:='0';

sDecimalSpinEdit3.Text:='1';
sDecimalSpinEdit4.Text:='1';
sDecimalSpinEdit5.Text:='1';
sDecimalSpinEdit6.Text:='1';
sDecimalSpinEdit7.Text:='1';
sDecimalSpinEdit8.Text:='1';
sDecimalSpinEdit9.Text:='1';

sComboBox1.Text:='';
sComboBox2.Text:='';
sComboBox3.Text:='';
sComboBox4.Text:='';
sComboBox5.Text:='';
sComboBox6.Text:='';
sComboBox7.Text:='';

sLabel12.Caption:='                   DA';


sPanel3.Visible:=false;
sGroupBox6.Visible:=false;
sPanel2.Visible:=false;
sGroupBox5.Visible:=false;

sGroupBox10.Visible:=false;
//sGroupBox11.Visible:=false;


sPanel4.Visible:=true;
sGroupBox8.Visible:=true;


DataModule2.ADOTable1.Close;
//datamodule2.DataSource1.DataSet:=datamodule2.ADOTable1;

DataModule2.ADOTable1.Open;
i:=datamodule2.ADOTable1.RecordCount;

sComboBox1.Items.BeginUpdate;
sComboBox1.Items.Clear;

sComboBox2.Items.BeginUpdate;
sComboBox2.Items.Clear;

sComboBox3.Items.BeginUpdate;
sComboBox3.Items.Clear;

sComboBox4.Items.BeginUpdate;
sComboBox4.Items.Clear;


sComboBox5.Items.BeginUpdate;
sComboBox5.Items.Clear;

sComboBox6.Items.BeginUpdate;
sComboBox6.Items.Clear;

sComboBox7.Items.BeginUpdate;
sComboBox7.Items.Clear;

sComboBox13.Items.BeginUpdate;
sComboBox13.Items.Clear;

sComboBox14.Items.BeginUpdate;
sComboBox14.Items.Clear;

sComboBox15.Items.BeginUpdate;
sComboBox15.Items.Clear;

sComboBox16.Items.BeginUpdate;
sComboBox16.Items.Clear;

sComboBox17.Items.BeginUpdate;
sComboBox17.Items.Clear;

sComboBox18.Items.BeginUpdate;
sComboBox18.Items.Clear;

sComboBox19.Items.BeginUpdate;
sComboBox19.Items.Clear;

sComboBox20.Items.BeginUpdate;
sComboBox20.Items.Clear;

sComboBox21.Items.BeginUpdate;
sComboBox21.Items.Clear;

sComboBox22.Items.BeginUpdate;
sComboBox22.Items.Clear;

sComboBox23.Items.BeginUpdate;
sComboBox23.Items.Clear;

sComboBox24.Items.BeginUpdate;
sComboBox24.Items.Clear;

sComboBox25.Items.BeginUpdate;
sComboBox25.Items.Clear;

sComboBox26.Items.BeginUpdate;
sComboBox26.Items.Clear;


sComboBox27.Items.BeginUpdate;
sComboBox27.Items.Clear;

sComboBox28.Items.BeginUpdate;
sComboBox28.Items.Clear;

sComboBox29.Items.BeginUpdate;
sComboBox29.Items.Clear;


sComboBox30.Items.BeginUpdate;
sComboBox30.Items.Clear;

sComboBox31.Items.BeginUpdate;
sComboBox31.Items.Clear;

sComboBox32.Items.BeginUpdate;
sComboBox32.Items.Clear;

sComboBox33.Items.BeginUpdate;
sComboBox33.Items.Clear;

sComboBox34.Items.BeginUpdate;
sComboBox34.Items.Clear;


sComboBox35.Items.BeginUpdate;
sComboBox35.Items.Clear;

sComboBox36.Items.BeginUpdate;
sComboBox36.Items.Clear;

sComboBox37.Items.BeginUpdate;
sComboBox37.Items.Clear;

for j := 1 to i do
begin
with datamodule2.ADOTable1 do
begin
sComboBox1.Items.Add(datamodule2.ADOTable1.FieldByName('nom_produit').AsString);
sComboBox2.Items.Add(datamodule2.ADOTable1.FieldByName('nom_produit').AsString);
sComboBox3.Items.Add(datamodule2.ADOTable1.FieldByName('nom_produit').AsString);
sComboBox4.Items.Add(datamodule2.ADOTable1.FieldByName('nom_produit').AsString);
sComboBox5.Items.Add(datamodule2.ADOTable1.FieldByName('nom_produit').AsString);
sComboBox6.Items.Add(datamodule2.ADOTable1.FieldByName('nom_produit').AsString);
sComboBox7.Items.Add(datamodule2.ADOTable1.FieldByName('nom_produit').AsString);
sComboBox13.Items.Add(datamodule2.ADOTable1.FieldByName('nom_produit').AsString);
sComboBox14.Items.Add(datamodule2.ADOTable1.FieldByName('nom_produit').AsString);
sComboBox15.Items.Add(datamodule2.ADOTable1.FieldByName('nom_produit').AsString);
sComboBox16.Items.Add(datamodule2.ADOTable1.FieldByName('nom_produit').AsString);
sComboBox17.Items.Add(datamodule2.ADOTable1.FieldByName('nom_produit').AsString);
sComboBox18.Items.Add(datamodule2.ADOTable1.FieldByName('nom_produit').AsString);
sComboBox19.Items.Add(datamodule2.ADOTable1.FieldByName('nom_produit').AsString);
sComboBox20.Items.Add(datamodule2.ADOTable1.FieldByName('nom_produit').AsString);
sComboBox21.Items.Add(datamodule2.ADOTable1.FieldByName('nom_produit').AsString);
sComboBox22.Items.Add(datamodule2.ADOTable1.FieldByName('nom_produit').AsString);
sComboBox23.Items.Add(datamodule2.ADOTable1.FieldByName('nom_produit').AsString);
sComboBox24.Items.Add(datamodule2.ADOTable1.FieldByName('nom_produit').AsString);
sComboBox25.Items.Add(datamodule2.ADOTable1.FieldByName('nom_produit').AsString);
sComboBox26.Items.Add(datamodule2.ADOTable1.FieldByName('nom_produit').AsString);
sComboBox27.Items.Add(datamodule2.ADOTable1.FieldByName('nom_produit').AsString);
sComboBox28.Items.Add(datamodule2.ADOTable1.FieldByName('nom_produit').AsString);
sComboBox29.Items.Add(datamodule2.ADOTable1.FieldByName('nom_produit').AsString);
sComboBox30.Items.Add(datamodule2.ADOTable1.FieldByName('nom_produit').AsString);
sComboBox31.Items.Add(datamodule2.ADOTable1.FieldByName('nom_produit').AsString);
sComboBox32.Items.Add(datamodule2.ADOTable1.FieldByName('nom_produit').AsString);
sComboBox33.Items.Add(datamodule2.ADOTable1.FieldByName('nom_produit').AsString);
sComboBox34.Items.Add(datamodule2.ADOTable1.FieldByName('nom_produit').AsString);
sComboBox35.Items.Add(datamodule2.ADOTable1.FieldByName('nom_produit').AsString);
sComboBox36.Items.Add(datamodule2.ADOTable1.FieldByName('nom_produit').AsString);
sComboBox37.Items.Add(datamodule2.ADOTable1.FieldByName('nom_produit').AsString);


end;
datamodule2.ADOTable1.Next;

end;
sComboBox1.Items.EndUpdate;
sComboBox2.Items.EndUpdate;
sComboBox3.Items.EndUpdate;
sComboBox4.Items.EndUpdate;
sComboBox5.Items.EndUpdate;
sComboBox6.Items.EndUpdate;
sComboBox7.Items.EndUpdate;

sComboBox13.Items.EndUpdate;
sComboBox14.Items.EndUpdate;
sComboBox15.Items.EndUpdate;
sComboBox16.Items.EndUpdate;
sComboBox17.Items.EndUpdate;
sComboBox18.Items.EndUpdate;
sComboBox19.Items.EndUpdate;

sComboBox20.Items.EndUpdate;
sComboBox21.Items.EndUpdate;
sComboBox22.Items.EndUpdate;
sComboBox23.Items.EndUpdate;
sComboBox24.Items.EndUpdate;
sComboBox25.Items.EndUpdate;
sComboBox26.Items.EndUpdate;

sComboBox27.Items.EndUpdate;
sComboBox28.Items.EndUpdate;
sComboBox29.Items.EndUpdate;
sComboBox30.Items.EndUpdate;
sComboBox31.Items.EndUpdate;
sComboBox32.Items.EndUpdate;
sComboBox33.Items.EndUpdate;
sComboBox34.Items.EndUpdate;
sComboBox35.Items.EndUpdate;
sComboBox36.Items.EndUpdate;
sComboBox37.Items.EndUpdate;



end;

procedure TForm1.sSpeedButton12Click(Sender: TObject);
var s: integer;
    i,j : integer;
begin
if MessageDlg('Etes-vous sûr de vouloir valider opération de vente ?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
begin
DataModule2.ADOTable2.Close;
datamodule2.ADOTable1.Close;
datamodule2.DataSource1.DataSet:=datamodule2.ADOTable1;
DataModule2.ADOTable1.Open;
DataModule2.ADOTable2.open;
i:=datamodule2.ADOTable1.RecordCount;
//s:=sLabel12.Caption;
//Delete(s,length(s)-1,2);
s:=strtoint(sCurrencyEdit4.Text)+strtoint(sCurrencyEdit5.Text)+strtoint(sCurrencyEdit6.Text)+strtoint(sCurrencyEdit7.Text)+strtoint(sCurrencyEdit8.Text)+strtoint(sCurrencyEdit9.Text)+strtoint(sCurrencyEdit10.Text)+strtoint(sCurrencyEdit11.Text)+strtoint(sCurrencyEdit12.Text)+strtoint(sCurrencyEdit13.Text)+strtoint(sCurrencyEdit14.Text)+strtoint(sCurrencyEdit15.Text)+strtoint(sCurrencyEdit16.Text)+strtoint(sCurrencyEdit17.Text)+strtoint(sCurrencyEdit18.Text)+strtoint(sCurrencyEdit19.Text)+strtoint(sCurrencyEdit20.Text)+strtoint(sCurrencyEdit21.Text)+strtoint(sCurrencyEdit22.Text)+strtoint(sCurrencyEdit23.Text)+strtoint(sCurrencyEdit24.Text)+strtoint(sCurrencyEdit25.Text)+strtoint(sCurrencyEdit26.Text)+strtoint(sCurrencyEdit27.Text)+strtoint(sCurrencyEdit28.Text)+strtoint(sCurrencyEdit29.Text)+strtoint(sCurrencyEdit30.Text)+strtoint(sCurrencyEdit31.Text)+strtoint(sCurrencyEdit32.Text)+strtoint(sCurrencyEdit33.Text);

for j := 1 to i do
begin
with datamodule2.ADOTable1 do
begin
if (sComboBox1.Text=datamodule2.ADOTable1.FieldByName('nom_produit').AsString) then
begin
datamodule2.ADOTable2.Append;

DataModule2.ADOTable2.FieldByName('id_produit').AsInteger:=DataModule2.ADOTable1.FieldByName('id_produit').AsInteger;

DataModule2.ADOTable2.FieldByName('qte_commande').Asinteger:=strtoint(sDecimalSpinEdit4.Text);
DataModule2.ADOTable2.FieldByName('date_facture').AsDateTime:=date();

DataModule2.ADOTable2.FieldByName('prix_facture').asinteger:=strtoint(sCurrencyEdit4.Text);

datamodule2.ADOTable2.Post;

datamodule2.ADOTable1.Edit;
datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger:=datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger-strtoint(sDecimalSpinEdit4.Text);
datamodule2.ADOTable1.Post;

end;

if (sComboBox2.Text=datamodule2.ADOTable1.FieldByName('nom_produit').AsString)  then
begin
datamodule2.ADOTable2.Append;
DataModule2.ADOTable2.FieldByName('id_produit').AsInteger:=DataModule2.ADOTable1.FieldByName('id_produit').AsInteger;
DataModule2.ADOTable2.FieldByName('qte_commande').Asstring:=sDecimalSpinEdit5.Text;
DataModule2.ADOTable2.FieldByName('date_facture').AsDateTime:=date();
DataModule2.ADOTable2.FieldByName('prix_facture').asinteger:=strtoint(sCurrencyEdit5.Text);
datamodule2.ADOTable2.Post;

datamodule2.ADOTable1.Edit;
datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger:=datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger-strtoint(sDecimalSpinEdit5.Text);
datamodule2.ADOTable1.Post;

end;

if (sComboBox3.Text=datamodule2.ADOTable1.FieldByName('nom_produit').AsString) then
begin
datamodule2.ADOTable2.Append;
DataModule2.ADOTable2.FieldByName('id_produit').AsInteger:=DataModule2.ADOTable1.FieldByName('id_produit').AsInteger;
DataModule2.ADOTable2.FieldByName('qte_commande').Asstring:=sDecimalSpinEdit6.Text;
DataModule2.ADOTable2.FieldByName('date_facture').AsDateTime:=date();
DataModule2.ADOTable2.FieldByName('prix_facture').asinteger:=strtoint(sCurrencyEdit6.Text);
datamodule2.ADOTable2.Post;

datamodule2.ADOTable1.Edit;
datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger:=datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger-strtoint(sDecimalSpinEdit6.Text);
datamodule2.ADOTable1.Post;

end;

if (sComboBox4.Text=datamodule2.ADOTable1.FieldByName('nom_produit').AsString) then
begin
datamodule2.ADOTable2.Append;
DataModule2.ADOTable2.FieldByName('id_produit').AsInteger:=DataModule2.ADOTable1.FieldByName('id_produit').AsInteger;
DataModule2.ADOTable2.FieldByName('qte_commande').Asstring:=sDecimalSpinEdit8.Text;
DataModule2.ADOTable2.FieldByName('date_facture').AsDateTime:=date();
DataModule2.ADOTable2.FieldByName('prix_facture').asinteger:=strtoint(sCurrencyEdit7.Text);
datamodule2.ADOTable2.Post;

datamodule2.ADOTable1.Edit;
datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger:=datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger-strtoint(sDecimalSpinEdit8.Text);
datamodule2.ADOTable1.Post;

end;

if (sComboBox5.Text=datamodule2.ADOTable1.FieldByName('nom_produit').AsString) then
begin
datamodule2.ADOTable2.Append;
DataModule2.ADOTable2.FieldByName('id_produit').AsInteger:=DataModule2.ADOTable1.FieldByName('id_produit').AsInteger;
DataModule2.ADOTable2.FieldByName('qte_commande').Asstring:=sDecimalSpinEdit9.Text;
DataModule2.ADOTable2.FieldByName('date_facture').AsDateTime:=date();
DataModule2.ADOTable2.FieldByName('prix_facture').asinteger:=strtoint(sCurrencyEdit8.Text);
datamodule2.ADOTable2.Post;

datamodule2.ADOTable1.Edit;

datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger:=datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger-strtoint(sDecimalSpinEdit9.Text);
datamodule2.ADOTable1.Post;

end;

if (sComboBox6.Text=datamodule2.ADOTable1.FieldByName('nom_produit').AsString)  then
begin
datamodule2.ADOTable2.Append;
DataModule2.ADOTable2.FieldByName('id_produit').AsInteger:=DataModule2.ADOTable1.FieldByName('id_produit').AsInteger;
DataModule2.ADOTable2.FieldByName('qte_commande').Asstring:=sDecimalSpinEdit10.Text;

DataModule2.ADOTable2.FieldByName('date_facture').AsDateTime:=date();
DataModule2.ADOTable2.FieldByName('prix_facture').asinteger:=strtoint(sCurrencyEdit9.Text);
datamodule2.ADOTable2.Post;

datamodule2.ADOTable1.Edit;
datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger:=datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger-strtoint(sDecimalSpinEdit10.Text);
datamodule2.ADOTable1.Post;

end;

if (sComboBox7.Text=datamodule2.ADOTable1.FieldByName('nom_produit').AsString)  then
begin
datamodule2.ADOTable2.Append;
datamodule2.ADOTable1.Edit;
DataModule2.ADOTable2.FieldByName('id_produit').AsInteger:=DataModule2.ADOTable1.FieldByName('id_produit').AsInteger;
DataModule2.ADOTable2.FieldByName('qte_commande').Asstring:=sDecimalSpinEdit7.Text;
DataModule2.ADOTable2.FieldByName('date_facture').AsDateTime:=date();
DataModule2.ADOTable2.FieldByName('prix_facture').asinteger:=strtoint(sCurrencyEdit10.Text);
datamodule2.ADOTable2.Post;
datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger:=datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger-strtoint(sDecimalSpinEdit7.Text);
datamodule2.ADOTable1.Post;
end;

if (sComboBox13.Text=datamodule2.ADOTable1.FieldByName('nom_produit').AsString)  then
begin
datamodule2.ADOTable2.Append;
datamodule2.ADOTable1.Edit;
DataModule2.ADOTable2.FieldByName('id_produit').AsInteger:=DataModule2.ADOTable1.FieldByName('id_produit').AsInteger;
DataModule2.ADOTable2.FieldByName('qte_commande').Asstring:=sDecimalSpinEdit11.Text;
DataModule2.ADOTable2.FieldByName('date_facture').AsDateTime:=date();
DataModule2.ADOTable2.FieldByName('prix_facture').asinteger:=strtoint(sCurrencyEdit11.Text);
datamodule2.ADOTable2.Post;
datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger:=datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger-strtoint(sDecimalSpinEdit11.Text);
datamodule2.ADOTable1.Post;
end;

if (sComboBox14.Text=datamodule2.ADOTable1.FieldByName('nom_produit').AsString)  then
begin
datamodule2.ADOTable2.Append;
datamodule2.ADOTable1.Edit;
DataModule2.ADOTable2.FieldByName('id_produit').AsInteger:=DataModule2.ADOTable1.FieldByName('id_produit').AsInteger;
DataModule2.ADOTable2.FieldByName('qte_commande').Asstring:=sDecimalSpinEdit12.Text;
DataModule2.ADOTable2.FieldByName('date_facture').AsDateTime:=date();
DataModule2.ADOTable2.FieldByName('prix_facture').asinteger:=strtoint(sCurrencyEdit12.Text);
datamodule2.ADOTable2.Post;
datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger:=datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger-strtoint(sDecimalSpinEdit12.Text);
datamodule2.ADOTable1.Post;
end;


if (sComboBox15.Text=datamodule2.ADOTable1.FieldByName('nom_produit').AsString)  then
begin
datamodule2.ADOTable2.Append;
datamodule2.ADOTable1.Edit;
DataModule2.ADOTable2.FieldByName('id_produit').AsInteger:=DataModule2.ADOTable1.FieldByName('id_produit').AsInteger;
DataModule2.ADOTable2.FieldByName('qte_commande').Asstring:=sDecimalSpinEdit13.Text;
DataModule2.ADOTable2.FieldByName('date_facture').AsDateTime:=date();
DataModule2.ADOTable2.FieldByName('prix_facture').asinteger:=strtoint(sCurrencyEdit13.Text);
datamodule2.ADOTable2.Post;
datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger:=datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger-strtoint(sDecimalSpinEdit13.Text);
datamodule2.ADOTable1.Post;
end;

if (sComboBox16.Text=datamodule2.ADOTable1.FieldByName('nom_produit').AsString)  then
begin
datamodule2.ADOTable2.Append;
datamodule2.ADOTable1.Edit;
DataModule2.ADOTable2.FieldByName('id_produit').AsInteger:=DataModule2.ADOTable1.FieldByName('id_produit').AsInteger;
DataModule2.ADOTable2.FieldByName('qte_commande').Asstring:=sDecimalSpinEdit14.Text;
DataModule2.ADOTable2.FieldByName('date_facture').AsDateTime:=date();
DataModule2.ADOTable2.FieldByName('prix_facture').asinteger:=strtoint(sCurrencyEdit14.Text);
datamodule2.ADOTable2.Post;
datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger:=datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger-strtoint(sDecimalSpinEdit14.Text);
datamodule2.ADOTable1.Post;
end;

if (sComboBox17.Text=datamodule2.ADOTable1.FieldByName('nom_produit').AsString)  then
begin
datamodule2.ADOTable2.Append;
datamodule2.ADOTable1.Edit;
DataModule2.ADOTable2.FieldByName('id_produit').AsInteger:=DataModule2.ADOTable1.FieldByName('id_produit').AsInteger;
DataModule2.ADOTable2.FieldByName('qte_commande').Asstring:=sDecimalSpinEdit15.Text;
DataModule2.ADOTable2.FieldByName('date_facture').AsDateTime:=date();
DataModule2.ADOTable2.FieldByName('prix_facture').asinteger:=strtoint(sCurrencyEdit15.Text);
datamodule2.ADOTable2.Post;
datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger:=datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger-strtoint(sDecimalSpinEdit15.Text);
datamodule2.ADOTable1.Post;
end;

if (sComboBox18.Text=datamodule2.ADOTable1.FieldByName('nom_produit').AsString)  then
begin
datamodule2.ADOTable2.Append;
datamodule2.ADOTable1.Edit;
DataModule2.ADOTable2.FieldByName('id_produit').AsInteger:=DataModule2.ADOTable1.FieldByName('id_produit').AsInteger;
DataModule2.ADOTable2.FieldByName('qte_commande').Asstring:=sDecimalSpinEdit16.Text;
DataModule2.ADOTable2.FieldByName('date_facture').AsDateTime:=date();
DataModule2.ADOTable2.FieldByName('prix_facture').asinteger:=strtoint(sCurrencyEdit16.Text);
datamodule2.ADOTable2.Post;
datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger:=datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger-strtoint(sDecimalSpinEdit16.Text);
datamodule2.ADOTable1.Post;
end;

if (sComboBox19.Text=datamodule2.ADOTable1.FieldByName('nom_produit').AsString)  then
begin
datamodule2.ADOTable2.Append;
datamodule2.ADOTable1.Edit;
DataModule2.ADOTable2.FieldByName('id_produit').AsInteger:=DataModule2.ADOTable1.FieldByName('id_produit').AsInteger;
DataModule2.ADOTable2.FieldByName('qte_commande').Asstring:=sDecimalSpinEdit17.Text;
DataModule2.ADOTable2.FieldByName('date_facture').AsDateTime:=date();
DataModule2.ADOTable2.FieldByName('prix_facture').asinteger:=strtoint(sCurrencyEdit17.Text);
datamodule2.ADOTable2.Post;
datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger:=datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger-strtoint(sDecimalSpinEdit17.Text);
datamodule2.ADOTable1.Post;
end;

if (sComboBox20.Text=datamodule2.ADOTable1.FieldByName('nom_produit').AsString)  then
begin
datamodule2.ADOTable2.Append;
datamodule2.ADOTable1.Edit;
DataModule2.ADOTable2.FieldByName('id_produit').AsInteger:=DataModule2.ADOTable1.FieldByName('id_produit').AsInteger;
DataModule2.ADOTable2.FieldByName('qte_commande').Asstring:=sDecimalSpinEdit18.Text;
DataModule2.ADOTable2.FieldByName('date_facture').AsDateTime:=date();
DataModule2.ADOTable2.FieldByName('prix_facture').asinteger:=strtoint(sCurrencyEdit18.Text);
datamodule2.ADOTable2.Post;
datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger:=datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger-strtoint(sDecimalSpinEdit18.Text);
datamodule2.ADOTable1.Post;
end;

if (sComboBox21.Text=datamodule2.ADOTable1.FieldByName('nom_produit').AsString)  then
begin
datamodule2.ADOTable2.Append;
datamodule2.ADOTable1.Edit;
DataModule2.ADOTable2.FieldByName('id_produit').AsInteger:=DataModule2.ADOTable1.FieldByName('id_produit').AsInteger;
DataModule2.ADOTable2.FieldByName('qte_commande').Asstring:=sDecimalSpinEdit19.Text;
DataModule2.ADOTable2.FieldByName('date_facture').AsDateTime:=date();
DataModule2.ADOTable2.FieldByName('prix_facture').asinteger:=strtoint(sCurrencyEdit19.Text);
datamodule2.ADOTable2.Post;
datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger:=datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger-strtoint(sDecimalSpinEdit19.Text);
datamodule2.ADOTable1.Post;
end;


if (sComboBox22.Text=datamodule2.ADOTable1.FieldByName('nom_produit').AsString)  then
begin
datamodule2.ADOTable2.Append;
datamodule2.ADOTable1.Edit;
DataModule2.ADOTable2.FieldByName('id_produit').AsInteger:=DataModule2.ADOTable1.FieldByName('id_produit').AsInteger;
DataModule2.ADOTable2.FieldByName('qte_commande').Asstring:=sDecimalSpinEdit20.Text;
DataModule2.ADOTable2.FieldByName('date_facture').AsDateTime:=date();
DataModule2.ADOTable2.FieldByName('prix_facture').asinteger:=strtoint(sCurrencyEdit20.Text);
datamodule2.ADOTable2.Post;
datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger:=datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger-strtoint(sDecimalSpinEdit20.Text);
datamodule2.ADOTable1.Post;
end;


if (sComboBox23.Text=datamodule2.ADOTable1.FieldByName('nom_produit').AsString)  then
begin
datamodule2.ADOTable2.Append;
datamodule2.ADOTable1.Edit;
DataModule2.ADOTable2.FieldByName('id_produit').AsInteger:=DataModule2.ADOTable1.FieldByName('id_produit').AsInteger;
DataModule2.ADOTable2.FieldByName('qte_commande').Asstring:=sDecimalSpinEdit21.Text;
DataModule2.ADOTable2.FieldByName('date_facture').AsDateTime:=date();
DataModule2.ADOTable2.FieldByName('prix_facture').asinteger:=strtoint(sCurrencyEdit21.Text);
datamodule2.ADOTable2.Post;
datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger:=datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger-strtoint(sDecimalSpinEdit21.Text);
datamodule2.ADOTable1.Post;
end;

if (sComboBox24.Text=datamodule2.ADOTable1.FieldByName('nom_produit').AsString)  then
begin
datamodule2.ADOTable2.Append;
datamodule2.ADOTable1.Edit;
DataModule2.ADOTable2.FieldByName('id_produit').AsInteger:=DataModule2.ADOTable1.FieldByName('id_produit').AsInteger;
DataModule2.ADOTable2.FieldByName('qte_commande').Asstring:=sDecimalSpinEdit22.Text;
DataModule2.ADOTable2.FieldByName('date_facture').AsDateTime:=date();
DataModule2.ADOTable2.FieldByName('prix_facture').asinteger:=strtoint(sCurrencyEdit22.Text);
datamodule2.ADOTable2.Post;
datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger:=datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger-strtoint(sDecimalSpinEdit22.Text);
datamodule2.ADOTable1.Post;
end;


if (sComboBox25.Text=datamodule2.ADOTable1.FieldByName('nom_produit').AsString)  then
begin
datamodule2.ADOTable2.Append;
datamodule2.ADOTable1.Edit;
DataModule2.ADOTable2.FieldByName('id_produit').AsInteger:=DataModule2.ADOTable1.FieldByName('id_produit').AsInteger;
DataModule2.ADOTable2.FieldByName('qte_commande').Asstring:=sDecimalSpinEdit23.Text;
DataModule2.ADOTable2.FieldByName('date_facture').AsDateTime:=date();
DataModule2.ADOTable2.FieldByName('prix_facture').asinteger:=strtoint(sCurrencyEdit23.Text);
datamodule2.ADOTable2.Post;
datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger:=datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger-strtoint(sDecimalSpinEdit23.Text);
datamodule2.ADOTable1.Post;
end;

if (sComboBox26.Text=datamodule2.ADOTable1.FieldByName('nom_produit').AsString)  then
begin
datamodule2.ADOTable2.Append;
datamodule2.ADOTable1.Edit;
DataModule2.ADOTable2.FieldByName('id_produit').AsInteger:=DataModule2.ADOTable1.FieldByName('id_produit').AsInteger;
DataModule2.ADOTable2.FieldByName('qte_commande').Asstring:=sDecimalSpinEdit24.Text;
DataModule2.ADOTable2.FieldByName('date_facture').AsDateTime:=date();
DataModule2.ADOTable2.FieldByName('prix_facture').asinteger:=strtoint(sCurrencyEdit24.Text);
datamodule2.ADOTable2.Post;
datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger:=datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger-strtoint(sDecimalSpinEdit24.Text);
datamodule2.ADOTable1.Post;
end;

if (sComboBox27.Text=datamodule2.ADOTable1.FieldByName('nom_produit').AsString)  then
begin
datamodule2.ADOTable2.Append;
datamodule2.ADOTable1.Edit;
DataModule2.ADOTable2.FieldByName('id_produit').AsInteger:=DataModule2.ADOTable1.FieldByName('id_produit').AsInteger;
DataModule2.ADOTable2.FieldByName('qte_commande').Asstring:=sDecimalSpinEdit25.Text;
DataModule2.ADOTable2.FieldByName('date_facture').AsDateTime:=date();
DataModule2.ADOTable2.FieldByName('prix_facture').asinteger:=strtoint(sCurrencyEdit25.Text);
datamodule2.ADOTable2.Post;
datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger:=datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger-strtoint(sDecimalSpinEdit25.Text);
datamodule2.ADOTable1.Post;
end;


if (sComboBox28.Text=datamodule2.ADOTable1.FieldByName('nom_produit').AsString)  then
begin
datamodule2.ADOTable2.Append;
datamodule2.ADOTable1.Edit;
DataModule2.ADOTable2.FieldByName('id_produit').AsInteger:=DataModule2.ADOTable1.FieldByName('id_produit').AsInteger;
DataModule2.ADOTable2.FieldByName('qte_commande').Asstring:=sDecimalSpinEdit26.Text;
DataModule2.ADOTable2.FieldByName('date_facture').AsDateTime:=date();
DataModule2.ADOTable2.FieldByName('prix_facture').asinteger:=strtoint(sCurrencyEdit26.Text);
datamodule2.ADOTable2.Post;
datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger:=datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger-strtoint(sDecimalSpinEdit26.Text);
datamodule2.ADOTable1.Post;
end;


if (sComboBox29.Text=datamodule2.ADOTable1.FieldByName('nom_produit').AsString)  then
begin
datamodule2.ADOTable2.Append;
datamodule2.ADOTable1.Edit;
DataModule2.ADOTable2.FieldByName('id_produit').AsInteger:=DataModule2.ADOTable1.FieldByName('id_produit').AsInteger;
DataModule2.ADOTable2.FieldByName('qte_commande').Asstring:=sDecimalSpinEdit27.Text;
DataModule2.ADOTable2.FieldByName('date_facture').AsDateTime:=date();
DataModule2.ADOTable2.FieldByName('prix_facture').asinteger:=strtoint(sCurrencyEdit27.Text);
datamodule2.ADOTable2.Post;
datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger:=datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger-strtoint(sDecimalSpinEdit27.Text);
datamodule2.ADOTable1.Post;
end;

if (sComboBox30.Text=datamodule2.ADOTable1.FieldByName('nom_produit').AsString)  then
begin
datamodule2.ADOTable2.Append;
datamodule2.ADOTable1.Edit;
DataModule2.ADOTable2.FieldByName('id_produit').AsInteger:=DataModule2.ADOTable1.FieldByName('id_produit').AsInteger;
DataModule2.ADOTable2.FieldByName('qte_commande').Asstring:=sDecimalSpinEdit28.Text;
DataModule2.ADOTable2.FieldByName('date_facture').AsDateTime:=date();
DataModule2.ADOTable2.FieldByName('prix_facture').asinteger:=strtoint(sCurrencyEdit28.Text);
datamodule2.ADOTable2.Post;
datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger:=datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger-strtoint(sDecimalSpinEdit28.Text);
datamodule2.ADOTable1.Post;
end;

if (sComboBox31.Text=datamodule2.ADOTable1.FieldByName('nom_produit').AsString)  then
begin
datamodule2.ADOTable2.Append;
datamodule2.ADOTable1.Edit;
DataModule2.ADOTable2.FieldByName('id_produit').AsInteger:=DataModule2.ADOTable1.FieldByName('id_produit').AsInteger;
DataModule2.ADOTable2.FieldByName('qte_commande').Asstring:=sDecimalSpinEdit29.Text;
DataModule2.ADOTable2.FieldByName('date_facture').AsDateTime:=date();
DataModule2.ADOTable2.FieldByName('prix_facture').asinteger:=strtoint(sCurrencyEdit29.Text);
datamodule2.ADOTable2.Post;
datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger:=datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger-strtoint(sDecimalSpinEdit29.Text);
datamodule2.ADOTable1.Post;
end;

if (sComboBox32.Text=datamodule2.ADOTable1.FieldByName('nom_produit').AsString)  then
begin
datamodule2.ADOTable2.Append;
datamodule2.ADOTable1.Edit;
DataModule2.ADOTable2.FieldByName('id_produit').AsInteger:=DataModule2.ADOTable1.FieldByName('id_produit').AsInteger;
DataModule2.ADOTable2.FieldByName('qte_commande').Asstring:=sDecimalSpinEdit30.Text;
DataModule2.ADOTable2.FieldByName('date_facture').AsDateTime:=date();
DataModule2.ADOTable2.FieldByName('prix_facture').asinteger:=strtoint(sCurrencyEdit30.Text);
datamodule2.ADOTable2.Post;
datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger:=datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger-strtoint(sDecimalSpinEdit30.Text);
datamodule2.ADOTable1.Post;
end;

if (sComboBox33.Text=datamodule2.ADOTable1.FieldByName('nom_produit').AsString)  then
begin
datamodule2.ADOTable2.Append;
datamodule2.ADOTable1.Edit;
DataModule2.ADOTable2.FieldByName('id_produit').AsInteger:=DataModule2.ADOTable1.FieldByName('id_produit').AsInteger;
DataModule2.ADOTable2.FieldByName('qte_commande').Asstring:=sDecimalSpinEdit31.Text;
DataModule2.ADOTable2.FieldByName('date_facture').AsDateTime:=date();
DataModule2.ADOTable2.FieldByName('prix_facture').asinteger:=strtoint(sCurrencyEdit31.Text);
datamodule2.ADOTable2.Post;
datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger:=datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger-strtoint(sDecimalSpinEdit31.Text);
datamodule2.ADOTable1.Post;
end;

if (sComboBox34.Text=datamodule2.ADOTable1.FieldByName('nom_produit').AsString)  then
begin
datamodule2.ADOTable2.Append;
datamodule2.ADOTable1.Edit;
DataModule2.ADOTable2.FieldByName('id_produit').AsInteger:=DataModule2.ADOTable1.FieldByName('id_produit').AsInteger;
DataModule2.ADOTable2.FieldByName('qte_commande').Asstring:=sDecimalSpinEdit32.Text;
DataModule2.ADOTable2.FieldByName('date_facture').AsDateTime:=date();
DataModule2.ADOTable2.FieldByName('prix_facture').asinteger:=strtoint(sCurrencyEdit32.Text);
datamodule2.ADOTable2.Post;
datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger:=datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger-strtoint(sDecimalSpinEdit32.Text);
datamodule2.ADOTable1.Post;
end;

if (sComboBox35.Text=datamodule2.ADOTable1.FieldByName('nom_produit').AsString)  then
begin
datamodule2.ADOTable2.Append;
datamodule2.ADOTable1.Edit;
DataModule2.ADOTable2.FieldByName('id_produit').AsInteger:=DataModule2.ADOTable1.FieldByName('id_produit').AsInteger;
DataModule2.ADOTable2.FieldByName('qte_commande').Asstring:=sDecimalSpinEdit33.Text;
DataModule2.ADOTable2.FieldByName('date_facture').AsDateTime:=date();
DataModule2.ADOTable2.FieldByName('prix_facture').asinteger:=strtoint(sCurrencyEdit33.Text);
datamodule2.ADOTable2.Post;
datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger:=datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger-strtoint(sDecimalSpinEdit33.Text);
datamodule2.ADOTable1.Post;
end;

if (sComboBox36.Text=datamodule2.ADOTable1.FieldByName('nom_produit').AsString)  then
begin
datamodule2.ADOTable2.Append;
datamodule2.ADOTable1.Edit;
DataModule2.ADOTable2.FieldByName('id_produit').AsInteger:=DataModule2.ADOTable1.FieldByName('id_produit').AsInteger;
DataModule2.ADOTable2.FieldByName('qte_commande').Asstring:=sDecimalSpinEdit34.Text;
DataModule2.ADOTable2.FieldByName('date_facture').AsDateTime:=date();
DataModule2.ADOTable2.FieldByName('prix_facture').asinteger:=strtoint(sCurrencyEdit34.Text);
datamodule2.ADOTable2.Post;
datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger:=datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger-strtoint(sDecimalSpinEdit34.Text);
datamodule2.ADOTable1.Post;
end;

if (sComboBox37.Text=datamodule2.ADOTable1.FieldByName('nom_produit').AsString)  then
begin
datamodule2.ADOTable2.Append;
datamodule2.ADOTable1.Edit;
DataModule2.ADOTable2.FieldByName('id_produit').AsInteger:=DataModule2.ADOTable1.FieldByName('id_produit').AsInteger;
DataModule2.ADOTable2.FieldByName('qte_commande').Asstring:=sDecimalSpinEdit35.Text;
DataModule2.ADOTable2.FieldByName('date_facture').AsDateTime:=date();
DataModule2.ADOTable2.FieldByName('prix_facture').asinteger:=strtoint(sCurrencyEdit35.Text);
datamodule2.ADOTable2.Post;
datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger:=datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger-strtoint(sDecimalSpinEdit35.Text);
datamodule2.ADOTable1.Post;
end;

datamodule2.ADOTable1.next;
end;
end;
end;

ShowMessage('Le client doit payer '+inttostr(s)+' DA' );
sPanel4.Visible:=false;
sGroupBox8.Visible:=false;
sPanel2.Visible:=false;
sGroupBox5.Visible:=false;

sPanel3.Visible:=false;
sGroupBox6.Visible:=false;


sCurrencyEdit4.Text:='0';
sCurrencyEdit5.Text:='0';
sCurrencyEdit6.Text:='0';
sCurrencyEdit7.Text:='0';
sCurrencyEdit8.Text:='0';
sCurrencyEdit9.Text:='0';
sCurrencyEdit10.Text:='0';
sCurrencyEdit11.Text:='0';
sCurrencyEdit12.Text:='0';
sCurrencyEdit13.Text:='0';
sCurrencyEdit14.Text:='0';
sCurrencyEdit15.Text:='0';
sCurrencyEdit16.Text:='0';
sCurrencyEdit17.Text:='0';
sCurrencyEdit18.Text:='0';
sCurrencyEdit19.Text:='0';
sCurrencyEdit20.Text:='0';
sCurrencyEdit21.Text:='0';
sCurrencyEdit22.Text:='0';
sCurrencyEdit23.Text:='0';
sCurrencyEdit24.Text:='0';
sCurrencyEdit25.Text:='0';
sCurrencyEdit26.Text:='0';
sCurrencyEdit27.Text:='0';
sCurrencyEdit28.Text:='0';
sCurrencyEdit29.Text:='0';
sCurrencyEdit30.Text:='0';
sCurrencyEdit31.Text:='0';
sCurrencyEdit32.Text:='0';
sCurrencyEdit33.Text:='0';






sDecimalSpinEdit3.Text:='1';
sDecimalSpinEdit4.Text:='1';
sDecimalSpinEdit5.Text:='1';
sDecimalSpinEdit6.Text:='1';
sDecimalSpinEdit7.Text:='1';
sDecimalSpinEdit8.Text:='1';
sDecimalSpinEdit9.Text:='1';
sDecimalSpinEdit10.Text:='1';
sDecimalSpinEdit11.Text:='1';
sDecimalSpinEdit12.Text:='1';
sDecimalSpinEdit13.Text:='1';
sDecimalSpinEdit14.Text:='1';
sDecimalSpinEdit15.Text:='1';
sDecimalSpinEdit16.Text:='1';
sDecimalSpinEdit17.Text:='1';
sDecimalSpinEdit18.Text:='1';
sDecimalSpinEdit19.Text:='1';
sDecimalSpinEdit20.Text:='1';
sDecimalSpinEdit21.Text:='1';
sDecimalSpinEdit22.Text:='1';
sDecimalSpinEdit23.Text:='1';
sDecimalSpinEdit24.Text:='1';
sDecimalSpinEdit25.Text:='1';
sDecimalSpinEdit26.Text:='1';
sDecimalSpinEdit27.Text:='1';
sDecimalSpinEdit28.Text:='1';
sDecimalSpinEdit29.Text:='1';
sDecimalSpinEdit30.Text:='1';
sDecimalSpinEdit31.Text:='1';
sDecimalSpinEdit32.Text:='1';
sDecimalSpinEdit33.Text:='1';
sDecimalSpinEdit34.Text:='1';
sDecimalSpinEdit35.Text:='1';




sComboBox1.Text:='';
sComboBox2.Text:='';
sComboBox3.Text:='';
sComboBox4.Text:='';
sComboBox5.Text:='';
sComboBox6.Text:='';
sComboBox7.Text:='';

sComboBox13.Text:='';
sComboBox14.Text:='';
sComboBox15.Text:='';
sComboBox16.Text:='';
sComboBox17.Text:='';
sComboBox18.Text:='';
sComboBox19.Text:='';

sComboBox20.Text:='';
sComboBox21.Text:='';
sComboBox22.Text:='';
sComboBox23.Text:='';
sComboBox24.Text:='';
sComboBox25.Text:='';
sComboBox26.Text:='';

sComboBox27.Text:='';
sComboBox28.Text:='';
sComboBox29.Text:='';
sComboBox30.Text:='';
sComboBox31.Text:='';
sComboBox32.Text:='';
sComboBox33.Text:='';
sComboBox34.Text:='';
sComboBox35.Text:='';
sComboBox36.Text:='';
sComboBox37.Text:='';


sLabel12.Caption:='                   DA';


end;

procedure TForm1.sSpeedButton13Click(Sender: TObject);
begin
if MessageDlg('Etes-vous sûr de vouloir annuler opération de vente ?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
begin
sPanel4.Visible:=false;
sGroupBox8.Visible:=false;
sPanel2.Visible:=false;
sGroupBox5.Visible:=false;

sPanel3.Visible:=false;
sGroupBox6.Visible:=false;


end;
end;

procedure TForm1.sSpeedButton14Click(Sender: TObject);
var k: integer;
    i,j : integer;
    H : string;
    s : integer;
word,document: variant;
//i:Integer;
aTable,bTable : OLEVariant;

chiffre_lettre : string;

C, D, centimes, dinars : string;
car : array [1..3] of string;
X, Y, Z : integer;
num_facture : integer;
jour,mois,annee,jour_nom:string;

o,m :integer;

begin
if MessageDlg('Etes-vous sûr de vouloir vendre et faire une facture ?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
begin
if (sedit3.Text='Nom/Entreprise') or (sedit3.Text='') then
begin
ShowMessage('Erreur !! Vous devez Saisir le nom du client !!');
end
else
begin
DataModule2.ADOTable2.Close;
datamodule2.ADOTable1.Close;
datamodule2.DataSource1.DataSet:=datamodule2.ADOTable1;
DataModule2.ADOTable1.Open;
DataModule2.ADOTable2.open;
m:=0;
i:=datamodule2.ADOTable1.RecordCount;
s:=strtoint(sCurrencyEdit4.Text)+strtoint(sCurrencyEdit5.Text)+strtoint(sCurrencyEdit6.Text)+strtoint(sCurrencyEdit7.Text)+strtoint(sCurrencyEdit8.Text)+strtoint(sCurrencyEdit9.Text)+strtoint(sCurrencyEdit10.Text)+strtoint(sCurrencyEdit11.Text)+strtoint(sCurrencyEdit12.Text)+strtoint(sCurrencyEdit13.Text)+strtoint(sCurrencyEdit14.Text)+strtoint(sCurrencyEdit15.Text)+strtoint(sCurrencyEdit16.Text)+strtoint(sCurrencyEdit17.Text)+strtoint(sCurrencyEdit18.Text)+strtoint(sCurrencyEdit19.Text)+strtoint(sCurrencyEdit20.Text)+strtoint(sCurrencyEdit21.Text)+strtoint(sCurrencyEdit22.Text)+strtoint(sCurrencyEdit23.Text)+strtoint(sCurrencyEdit24.Text)+strtoint(sCurrencyEdit25.Text)+strtoint(sCurrencyEdit26.Text)+strtoint(sCurrencyEdit27.Text)+strtoint(sCurrencyEdit28.Text)+strtoint(sCurrencyEdit29.Text)+strtoint(sCurrencyEdit30.Text)+strtoint(sCurrencyEdit31.Text)+strtoint(sCurrencyEdit32.Text)+strtoint(sCurrencyEdit33.Text);

for j := 1 to i do
begin
with datamodule2.ADOTable1 do
begin
if (sComboBox1.Text=datamodule2.ADOTable1.FieldByName('nom_produit').AsString) then
begin
datamodule2.ADOTable2.Append;

DataModule2.ADOTable2.FieldByName('id_produit').AsInteger:=DataModule2.ADOTable1.FieldByName('id_produit').AsInteger;

DataModule2.ADOTable2.FieldByName('qte_commande').Asinteger:=strtoint(sDecimalSpinEdit4.Text);
DataModule2.ADOTable2.FieldByName('date_facture').AsDateTime:=date();

DataModule2.ADOTable2.FieldByName('prix_facture').asinteger:=strtoint(sCurrencyEdit4.Text);

datamodule2.ADOTable2.Post;

datamodule2.ADOTable1.Edit;
datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger:=datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger-strtoint(sDecimalSpinEdit4.Text);
datamodule2.ADOTable1.Post;

end;

if (sComboBox2.Text=datamodule2.ADOTable1.FieldByName('nom_produit').AsString)  then
begin
datamodule2.ADOTable2.Append;
DataModule2.ADOTable2.FieldByName('id_produit').AsInteger:=DataModule2.ADOTable1.FieldByName('id_produit').AsInteger;
DataModule2.ADOTable2.FieldByName('qte_commande').Asstring:=sDecimalSpinEdit5.Text;
DataModule2.ADOTable2.FieldByName('date_facture').AsDateTime:=date();
DataModule2.ADOTable2.FieldByName('prix_facture').asinteger:=strtoint(sCurrencyEdit5.Text);
datamodule2.ADOTable2.Post;

datamodule2.ADOTable1.Edit;
datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger:=datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger-strtoint(sDecimalSpinEdit5.Text);
datamodule2.ADOTable1.Post;

end;

if (sComboBox3.Text=datamodule2.ADOTable1.FieldByName('nom_produit').AsString) then
begin
datamodule2.ADOTable2.Append;
DataModule2.ADOTable2.FieldByName('id_produit').AsInteger:=DataModule2.ADOTable1.FieldByName('id_produit').AsInteger;
DataModule2.ADOTable2.FieldByName('qte_commande').Asstring:=sDecimalSpinEdit6.Text;
DataModule2.ADOTable2.FieldByName('date_facture').AsDateTime:=date();
DataModule2.ADOTable2.FieldByName('prix_facture').asinteger:=strtoint(sCurrencyEdit6.Text);
datamodule2.ADOTable2.Post;

datamodule2.ADOTable1.Edit;
datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger:=datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger-strtoint(sDecimalSpinEdit6.Text);
datamodule2.ADOTable1.Post;

end;

if (sComboBox4.Text=datamodule2.ADOTable1.FieldByName('nom_produit').AsString) then
begin
datamodule2.ADOTable2.Append;
DataModule2.ADOTable2.FieldByName('id_produit').AsInteger:=DataModule2.ADOTable1.FieldByName('id_produit').AsInteger;
DataModule2.ADOTable2.FieldByName('qte_commande').Asstring:=sDecimalSpinEdit8.Text;
DataModule2.ADOTable2.FieldByName('date_facture').AsDateTime:=date();
DataModule2.ADOTable2.FieldByName('prix_facture').asinteger:=strtoint(sCurrencyEdit7.Text);
datamodule2.ADOTable2.Post;

datamodule2.ADOTable1.Edit;
datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger:=datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger-strtoint(sDecimalSpinEdit8.Text);
datamodule2.ADOTable1.Post;

end;

if (sComboBox5.Text=datamodule2.ADOTable1.FieldByName('nom_produit').AsString) then
begin
datamodule2.ADOTable2.Append;
DataModule2.ADOTable2.FieldByName('id_produit').AsInteger:=DataModule2.ADOTable1.FieldByName('id_produit').AsInteger;
DataModule2.ADOTable2.FieldByName('qte_commande').Asstring:=sDecimalSpinEdit9.Text;
DataModule2.ADOTable2.FieldByName('date_facture').AsDateTime:=date();
DataModule2.ADOTable2.FieldByName('prix_facture').asinteger:=strtoint(sCurrencyEdit8.Text);
datamodule2.ADOTable2.Post;

datamodule2.ADOTable1.Edit;

datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger:=datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger-strtoint(sDecimalSpinEdit9.Text);
datamodule2.ADOTable1.Post;

end;

if (sComboBox6.Text=datamodule2.ADOTable1.FieldByName('nom_produit').AsString)  then
begin
datamodule2.ADOTable2.Append;
DataModule2.ADOTable2.FieldByName('id_produit').AsInteger:=DataModule2.ADOTable1.FieldByName('id_produit').AsInteger;
DataModule2.ADOTable2.FieldByName('qte_commande').Asstring:=sDecimalSpinEdit10.Text;

DataModule2.ADOTable2.FieldByName('date_facture').AsDateTime:=date();
DataModule2.ADOTable2.FieldByName('prix_facture').asinteger:=strtoint(sCurrencyEdit9.Text);
datamodule2.ADOTable2.Post;

datamodule2.ADOTable1.Edit;

datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger:=datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger-strtoint(sDecimalSpinEdit10.Text);
datamodule2.ADOTable1.Post;

end;

if (sComboBox7.Text=datamodule2.ADOTable1.FieldByName('nom_produit').AsString)  then
begin
datamodule2.ADOTable2.Append;
DataModule2.ADOTable2.FieldByName('id_produit').AsInteger:=DataModule2.ADOTable1.FieldByName('id_produit').AsInteger;
DataModule2.ADOTable2.FieldByName('qte_commande').Asstring:=sDecimalSpinEdit7.Text;
DataModule2.ADOTable2.FieldByName('date_facture').AsDateTime:=date();
DataModule2.ADOTable2.FieldByName('prix_facture').asinteger:=strtoint(sCurrencyEdit10.Text);
datamodule2.ADOTable2.Post;
datamodule2.ADOTable1.Edit;

datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger:=datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger-strtoint(sDecimalSpinEdit7.Text);
datamodule2.ADOTable1.Post;

end;



if (sComboBox13.Text=datamodule2.ADOTable1.FieldByName('nom_produit').AsString)  then
begin
datamodule2.ADOTable2.Append;
DataModule2.ADOTable2.FieldByName('id_produit').AsInteger:=DataModule2.ADOTable1.FieldByName('id_produit').AsInteger;
DataModule2.ADOTable2.FieldByName('qte_commande').Asstring:=sDecimalSpinEdit11.Text;
DataModule2.ADOTable2.FieldByName('date_facture').AsDateTime:=date();
DataModule2.ADOTable2.FieldByName('prix_facture').asinteger:=strtoint(sCurrencyEdit11.Text);
datamodule2.ADOTable2.Post;
datamodule2.ADOTable1.Edit;

datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger:=datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger-strtoint(sDecimalSpinEdit11.Text);
datamodule2.ADOTable1.Post;
end;

if (sComboBox14.Text=datamodule2.ADOTable1.FieldByName('nom_produit').AsString)  then
begin
datamodule2.ADOTable2.Append;
datamodule2.ADOTable1.Edit;
DataModule2.ADOTable2.FieldByName('id_produit').AsInteger:=DataModule2.ADOTable1.FieldByName('id_produit').AsInteger;
DataModule2.ADOTable2.FieldByName('qte_commande').Asstring:=sDecimalSpinEdit12.Text;
DataModule2.ADOTable2.FieldByName('date_facture').AsDateTime:=date();
DataModule2.ADOTable2.FieldByName('prix_facture').asinteger:=strtoint(sCurrencyEdit12.Text);
datamodule2.ADOTable2.Post;
datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger:=datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger-strtoint(sDecimalSpinEdit12.Text);
datamodule2.ADOTable1.Post;
end;


if (sComboBox15.Text=datamodule2.ADOTable1.FieldByName('nom_produit').AsString)  then
begin
datamodule2.ADOTable2.Append;
datamodule2.ADOTable1.Edit;
DataModule2.ADOTable2.FieldByName('id_produit').AsInteger:=DataModule2.ADOTable1.FieldByName('id_produit').AsInteger;
DataModule2.ADOTable2.FieldByName('qte_commande').Asstring:=sDecimalSpinEdit13.Text;
DataModule2.ADOTable2.FieldByName('date_facture').AsDateTime:=date();
DataModule2.ADOTable2.FieldByName('prix_facture').asinteger:=strtoint(sCurrencyEdit13.Text);
datamodule2.ADOTable2.Post;
datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger:=datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger-strtoint(sDecimalSpinEdit13.Text);
datamodule2.ADOTable1.Post;
end;

if (sComboBox16.Text=datamodule2.ADOTable1.FieldByName('nom_produit').AsString)  then
begin
datamodule2.ADOTable2.Append;
datamodule2.ADOTable1.Edit;
DataModule2.ADOTable2.FieldByName('id_produit').AsInteger:=DataModule2.ADOTable1.FieldByName('id_produit').AsInteger;
DataModule2.ADOTable2.FieldByName('qte_commande').Asstring:=sDecimalSpinEdit14.Text;
DataModule2.ADOTable2.FieldByName('date_facture').AsDateTime:=date();
DataModule2.ADOTable2.FieldByName('prix_facture').asinteger:=strtoint(sCurrencyEdit14.Text);
datamodule2.ADOTable2.Post;
datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger:=datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger-strtoint(sDecimalSpinEdit14.Text);
datamodule2.ADOTable1.Post;
end;

if (sComboBox17.Text=datamodule2.ADOTable1.FieldByName('nom_produit').AsString)  then
begin
datamodule2.ADOTable2.Append;
datamodule2.ADOTable1.Edit;
DataModule2.ADOTable2.FieldByName('id_produit').AsInteger:=DataModule2.ADOTable1.FieldByName('id_produit').AsInteger;
DataModule2.ADOTable2.FieldByName('qte_commande').Asstring:=sDecimalSpinEdit15.Text;
DataModule2.ADOTable2.FieldByName('date_facture').AsDateTime:=date();
DataModule2.ADOTable2.FieldByName('prix_facture').asinteger:=strtoint(sCurrencyEdit15.Text);
datamodule2.ADOTable2.Post;
datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger:=datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger-strtoint(sDecimalSpinEdit15.Text);
datamodule2.ADOTable1.Post;
end;

if (sComboBox18.Text=datamodule2.ADOTable1.FieldByName('nom_produit').AsString)  then
begin
datamodule2.ADOTable2.Append;
datamodule2.ADOTable1.Edit;
DataModule2.ADOTable2.FieldByName('id_produit').AsInteger:=DataModule2.ADOTable1.FieldByName('id_produit').AsInteger;
DataModule2.ADOTable2.FieldByName('qte_commande').Asstring:=sDecimalSpinEdit16.Text;
DataModule2.ADOTable2.FieldByName('date_facture').AsDateTime:=date();
DataModule2.ADOTable2.FieldByName('prix_facture').asinteger:=strtoint(sCurrencyEdit16.Text);
datamodule2.ADOTable2.Post;
datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger:=datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger-strtoint(sDecimalSpinEdit16.Text);
datamodule2.ADOTable1.Post;
end;

if (sComboBox19.Text=datamodule2.ADOTable1.FieldByName('nom_produit').AsString)  then
begin
datamodule2.ADOTable2.Append;
datamodule2.ADOTable1.Edit;
DataModule2.ADOTable2.FieldByName('id_produit').AsInteger:=DataModule2.ADOTable1.FieldByName('id_produit').AsInteger;
DataModule2.ADOTable2.FieldByName('qte_commande').Asstring:=sDecimalSpinEdit17.Text;
DataModule2.ADOTable2.FieldByName('date_facture').AsDateTime:=date();
DataModule2.ADOTable2.FieldByName('prix_facture').asinteger:=strtoint(sCurrencyEdit17.Text);
datamodule2.ADOTable2.Post;
datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger:=datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger-strtoint(sDecimalSpinEdit17.Text);
datamodule2.ADOTable1.Post;
end;

if (sComboBox20.Text=datamodule2.ADOTable1.FieldByName('nom_produit').AsString)  then
begin
datamodule2.ADOTable2.Append;
datamodule2.ADOTable1.Edit;
DataModule2.ADOTable2.FieldByName('id_produit').AsInteger:=DataModule2.ADOTable1.FieldByName('id_produit').AsInteger;
DataModule2.ADOTable2.FieldByName('qte_commande').Asstring:=sDecimalSpinEdit18.Text;
DataModule2.ADOTable2.FieldByName('date_facture').AsDateTime:=date();
DataModule2.ADOTable2.FieldByName('prix_facture').asinteger:=strtoint(sCurrencyEdit18.Text);
datamodule2.ADOTable2.Post;
datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger:=datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger-strtoint(sDecimalSpinEdit18.Text);
datamodule2.ADOTable1.Post;
end;

if (sComboBox21.Text=datamodule2.ADOTable1.FieldByName('nom_produit').AsString)  then
begin
datamodule2.ADOTable2.Append;
datamodule2.ADOTable1.Edit;
DataModule2.ADOTable2.FieldByName('id_produit').AsInteger:=DataModule2.ADOTable1.FieldByName('id_produit').AsInteger;
DataModule2.ADOTable2.FieldByName('qte_commande').Asstring:=sDecimalSpinEdit19.Text;
DataModule2.ADOTable2.FieldByName('date_facture').AsDateTime:=date();
DataModule2.ADOTable2.FieldByName('prix_facture').asinteger:=strtoint(sCurrencyEdit19.Text);
datamodule2.ADOTable2.Post;
datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger:=datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger-strtoint(sDecimalSpinEdit19.Text);
datamodule2.ADOTable1.Post;
end;


if (sComboBox22.Text=datamodule2.ADOTable1.FieldByName('nom_produit').AsString)  then
begin
datamodule2.ADOTable2.Append;
datamodule2.ADOTable1.Edit;
DataModule2.ADOTable2.FieldByName('id_produit').AsInteger:=DataModule2.ADOTable1.FieldByName('id_produit').AsInteger;
DataModule2.ADOTable2.FieldByName('qte_commande').Asstring:=sDecimalSpinEdit20.Text;
DataModule2.ADOTable2.FieldByName('date_facture').AsDateTime:=date();
DataModule2.ADOTable2.FieldByName('prix_facture').asinteger:=strtoint(sCurrencyEdit20.Text);
datamodule2.ADOTable2.Post;
datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger:=datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger-strtoint(sDecimalSpinEdit20.Text);
datamodule2.ADOTable1.Post;
end;


if (sComboBox23.Text=datamodule2.ADOTable1.FieldByName('nom_produit').AsString)  then
begin
datamodule2.ADOTable2.Append;
datamodule2.ADOTable1.Edit;
DataModule2.ADOTable2.FieldByName('id_produit').AsInteger:=DataModule2.ADOTable1.FieldByName('id_produit').AsInteger;
DataModule2.ADOTable2.FieldByName('qte_commande').Asstring:=sDecimalSpinEdit21.Text;
DataModule2.ADOTable2.FieldByName('date_facture').AsDateTime:=date();
DataModule2.ADOTable2.FieldByName('prix_facture').asinteger:=strtoint(sCurrencyEdit21.Text);
datamodule2.ADOTable2.Post;
datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger:=datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger-strtoint(sDecimalSpinEdit21.Text);
datamodule2.ADOTable1.Post;
end;

if (sComboBox24.Text=datamodule2.ADOTable1.FieldByName('nom_produit').AsString)  then
begin
datamodule2.ADOTable2.Append;
datamodule2.ADOTable1.Edit;
DataModule2.ADOTable2.FieldByName('id_produit').AsInteger:=DataModule2.ADOTable1.FieldByName('id_produit').AsInteger;
DataModule2.ADOTable2.FieldByName('qte_commande').Asstring:=sDecimalSpinEdit22.Text;
DataModule2.ADOTable2.FieldByName('date_facture').AsDateTime:=date();
DataModule2.ADOTable2.FieldByName('prix_facture').asinteger:=strtoint(sCurrencyEdit22.Text);
datamodule2.ADOTable2.Post;
datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger:=datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger-strtoint(sDecimalSpinEdit22.Text);
datamodule2.ADOTable1.Post;
end;


if (sComboBox25.Text=datamodule2.ADOTable1.FieldByName('nom_produit').AsString)  then
begin
datamodule2.ADOTable2.Append;
datamodule2.ADOTable1.Edit;
DataModule2.ADOTable2.FieldByName('id_produit').AsInteger:=DataModule2.ADOTable1.FieldByName('id_produit').AsInteger;
DataModule2.ADOTable2.FieldByName('qte_commande').Asstring:=sDecimalSpinEdit23.Text;
DataModule2.ADOTable2.FieldByName('date_facture').AsDateTime:=date();
DataModule2.ADOTable2.FieldByName('prix_facture').asinteger:=strtoint(sCurrencyEdit23.Text);
datamodule2.ADOTable2.Post;
datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger:=datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger-strtoint(sDecimalSpinEdit23.Text);
datamodule2.ADOTable1.Post;
end;

if (sComboBox26.Text=datamodule2.ADOTable1.FieldByName('nom_produit').AsString)  then
begin
datamodule2.ADOTable2.Append;
datamodule2.ADOTable1.Edit;
DataModule2.ADOTable2.FieldByName('id_produit').AsInteger:=DataModule2.ADOTable1.FieldByName('id_produit').AsInteger;
DataModule2.ADOTable2.FieldByName('qte_commande').Asstring:=sDecimalSpinEdit24.Text;
DataModule2.ADOTable2.FieldByName('date_facture').AsDateTime:=date();
DataModule2.ADOTable2.FieldByName('prix_facture').asinteger:=strtoint(sCurrencyEdit24.Text);
datamodule2.ADOTable2.Post;
datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger:=datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger-strtoint(sDecimalSpinEdit24.Text);
datamodule2.ADOTable1.Post;
end;

if (sComboBox27.Text=datamodule2.ADOTable1.FieldByName('nom_produit').AsString)  then
begin
datamodule2.ADOTable2.Append;
datamodule2.ADOTable1.Edit;
DataModule2.ADOTable2.FieldByName('id_produit').AsInteger:=DataModule2.ADOTable1.FieldByName('id_produit').AsInteger;
DataModule2.ADOTable2.FieldByName('qte_commande').Asstring:=sDecimalSpinEdit25.Text;
DataModule2.ADOTable2.FieldByName('date_facture').AsDateTime:=date();
DataModule2.ADOTable2.FieldByName('prix_facture').asinteger:=strtoint(sCurrencyEdit25.Text);
datamodule2.ADOTable2.Post;
datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger:=datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger-strtoint(sDecimalSpinEdit25.Text);
datamodule2.ADOTable1.Post;
end;


if (sComboBox28.Text=datamodule2.ADOTable1.FieldByName('nom_produit').AsString)  then
begin
datamodule2.ADOTable2.Append;
datamodule2.ADOTable1.Edit;
DataModule2.ADOTable2.FieldByName('id_produit').AsInteger:=DataModule2.ADOTable1.FieldByName('id_produit').AsInteger;
DataModule2.ADOTable2.FieldByName('qte_commande').Asstring:=sDecimalSpinEdit26.Text;
DataModule2.ADOTable2.FieldByName('date_facture').AsDateTime:=date();
DataModule2.ADOTable2.FieldByName('prix_facture').asinteger:=strtoint(sCurrencyEdit26.Text);
datamodule2.ADOTable2.Post;
datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger:=datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger-strtoint(sDecimalSpinEdit26.Text);
datamodule2.ADOTable1.Post;
end;


if (sComboBox29.Text=datamodule2.ADOTable1.FieldByName('nom_produit').AsString)  then
begin
datamodule2.ADOTable2.Append;
datamodule2.ADOTable1.Edit;
DataModule2.ADOTable2.FieldByName('id_produit').AsInteger:=DataModule2.ADOTable1.FieldByName('id_produit').AsInteger;
DataModule2.ADOTable2.FieldByName('qte_commande').Asstring:=sDecimalSpinEdit27.Text;
DataModule2.ADOTable2.FieldByName('date_facture').AsDateTime:=date();
DataModule2.ADOTable2.FieldByName('prix_facture').asinteger:=strtoint(sCurrencyEdit27.Text);
datamodule2.ADOTable2.Post;
datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger:=datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger-strtoint(sDecimalSpinEdit27.Text);
datamodule2.ADOTable1.Post;
end;

if (sComboBox30.Text=datamodule2.ADOTable1.FieldByName('nom_produit').AsString)  then
begin
datamodule2.ADOTable2.Append;
datamodule2.ADOTable1.Edit;
DataModule2.ADOTable2.FieldByName('id_produit').AsInteger:=DataModule2.ADOTable1.FieldByName('id_produit').AsInteger;
DataModule2.ADOTable2.FieldByName('qte_commande').Asstring:=sDecimalSpinEdit28.Text;
DataModule2.ADOTable2.FieldByName('date_facture').AsDateTime:=date();
DataModule2.ADOTable2.FieldByName('prix_facture').asinteger:=strtoint(sCurrencyEdit28.Text);
datamodule2.ADOTable2.Post;
datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger:=datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger-strtoint(sDecimalSpinEdit28.Text);
datamodule2.ADOTable1.Post;
end;

if (sComboBox31.Text=datamodule2.ADOTable1.FieldByName('nom_produit').AsString)  then
begin
datamodule2.ADOTable2.Append;
datamodule2.ADOTable1.Edit;
DataModule2.ADOTable2.FieldByName('id_produit').AsInteger:=DataModule2.ADOTable1.FieldByName('id_produit').AsInteger;
DataModule2.ADOTable2.FieldByName('qte_commande').Asstring:=sDecimalSpinEdit29.Text;
DataModule2.ADOTable2.FieldByName('date_facture').AsDateTime:=date();
DataModule2.ADOTable2.FieldByName('prix_facture').asinteger:=strtoint(sCurrencyEdit29.Text);
datamodule2.ADOTable2.Post;
datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger:=datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger-strtoint(sDecimalSpinEdit29.Text);
datamodule2.ADOTable1.Post;
end;

if (sComboBox32.Text=datamodule2.ADOTable1.FieldByName('nom_produit').AsString)  then
begin
datamodule2.ADOTable2.Append;
datamodule2.ADOTable1.Edit;
DataModule2.ADOTable2.FieldByName('id_produit').AsInteger:=DataModule2.ADOTable1.FieldByName('id_produit').AsInteger;
DataModule2.ADOTable2.FieldByName('qte_commande').Asstring:=sDecimalSpinEdit30.Text;
DataModule2.ADOTable2.FieldByName('date_facture').AsDateTime:=date();
DataModule2.ADOTable2.FieldByName('prix_facture').asinteger:=strtoint(sCurrencyEdit30.Text);
datamodule2.ADOTable2.Post;
datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger:=datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger-strtoint(sDecimalSpinEdit30.Text);
datamodule2.ADOTable1.Post;
end;

if (sComboBox33.Text=datamodule2.ADOTable1.FieldByName('nom_produit').AsString)  then
begin
datamodule2.ADOTable2.Append;
datamodule2.ADOTable1.Edit;
DataModule2.ADOTable2.FieldByName('id_produit').AsInteger:=DataModule2.ADOTable1.FieldByName('id_produit').AsInteger;
DataModule2.ADOTable2.FieldByName('qte_commande').Asstring:=sDecimalSpinEdit31.Text;
DataModule2.ADOTable2.FieldByName('date_facture').AsDateTime:=date();
DataModule2.ADOTable2.FieldByName('prix_facture').asinteger:=strtoint(sCurrencyEdit31.Text);
datamodule2.ADOTable2.Post;
datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger:=datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger-strtoint(sDecimalSpinEdit31.Text);
datamodule2.ADOTable1.Post;
end;

if (sComboBox34.Text=datamodule2.ADOTable1.FieldByName('nom_produit').AsString)  then
begin
datamodule2.ADOTable2.Append;
datamodule2.ADOTable1.Edit;
DataModule2.ADOTable2.FieldByName('id_produit').AsInteger:=DataModule2.ADOTable1.FieldByName('id_produit').AsInteger;
DataModule2.ADOTable2.FieldByName('qte_commande').Asstring:=sDecimalSpinEdit32.Text;
DataModule2.ADOTable2.FieldByName('date_facture').AsDateTime:=date();
DataModule2.ADOTable2.FieldByName('prix_facture').asinteger:=strtoint(sCurrencyEdit32.Text);
datamodule2.ADOTable2.Post;
datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger:=datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger-strtoint(sDecimalSpinEdit32.Text);
datamodule2.ADOTable1.Post;
end;

if (sComboBox35.Text=datamodule2.ADOTable1.FieldByName('nom_produit').AsString)  then
begin
datamodule2.ADOTable2.Append;
datamodule2.ADOTable1.Edit;
DataModule2.ADOTable2.FieldByName('id_produit').AsInteger:=DataModule2.ADOTable1.FieldByName('id_produit').AsInteger;
DataModule2.ADOTable2.FieldByName('qte_commande').Asstring:=sDecimalSpinEdit33.Text;
DataModule2.ADOTable2.FieldByName('date_facture').AsDateTime:=date();
DataModule2.ADOTable2.FieldByName('prix_facture').asinteger:=strtoint(sCurrencyEdit33.Text);
datamodule2.ADOTable2.Post;
datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger:=datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger-strtoint(sDecimalSpinEdit33.Text);
datamodule2.ADOTable1.Post;
end;

if (sComboBox36.Text=datamodule2.ADOTable1.FieldByName('nom_produit').AsString)  then
begin
datamodule2.ADOTable2.Append;
datamodule2.ADOTable1.Edit;
DataModule2.ADOTable2.FieldByName('id_produit').AsInteger:=DataModule2.ADOTable1.FieldByName('id_produit').AsInteger;
DataModule2.ADOTable2.FieldByName('qte_commande').Asstring:=sDecimalSpinEdit34.Text;
DataModule2.ADOTable2.FieldByName('date_facture').AsDateTime:=date();
DataModule2.ADOTable2.FieldByName('prix_facture').asinteger:=strtoint(sCurrencyEdit34.Text);
datamodule2.ADOTable2.Post;
datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger:=datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger-strtoint(sDecimalSpinEdit34.Text);
datamodule2.ADOTable1.Post;
end;

if (sComboBox37.Text=datamodule2.ADOTable1.FieldByName('nom_produit').AsString)  then
begin
datamodule2.ADOTable2.Append;
datamodule2.ADOTable1.Edit;
DataModule2.ADOTable2.FieldByName('id_produit').AsInteger:=DataModule2.ADOTable1.FieldByName('id_produit').AsInteger;
DataModule2.ADOTable2.FieldByName('qte_commande').Asstring:=sDecimalSpinEdit35.Text;
DataModule2.ADOTable2.FieldByName('date_facture').AsDateTime:=date();
DataModule2.ADOTable2.FieldByName('prix_facture').asinteger:=strtoint(sCurrencyEdit35.Text);
datamodule2.ADOTable2.Post;
datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger:=datamodule2.ADOTable1.FieldByName('quantite_produit').AsInteger-strtoint(sDecimalSpinEdit35.Text);
datamodule2.ADOTable1.Post;
end;





datamodule2.ADOTable1.next;
end;
end;

datamodule2.ADOTable2.Close;
datamodule2.ADOTable2.open;
while not(datamodule2.ADOTable2.Eof) do
begin
datamodule2.ADOTable2.Next;
end;
annee:=copy(datetostr(date),9,2);
datamodule2.ADOTable2.Close;
ShowMessage('Le client doit payer '+intToStr(s)+' DA');

datamodule2.ADOTable3.Close;
datamodule2.ADOTable3.open;
datamodule2.ADOTable3.Last;

num_facture:=datamodule2.ADOTable3.FieldByName('id_num_facture').AsInteger+1;

datamodule2.ADOTable3.Append;
datamodule2.ADOTable3.FieldByName('id_num_facture').AsInteger:=num_facture;
datamodule2.ADOTable3.FieldByName('prix_facture').Value:=s;

datamodule2.ADOTable3.Post;
datamodule2.ADOTable3.Close;

if sComboBox1.Text<>'' then m:=m+1;
if sComboBox2.Text<>'' then m:=m+1;
if sComboBox3.Text<>'' then m:=m+1;
if sComboBox4.Text<>'' then m:=m+1;
if sComboBox5.Text<>'' then m:=m+1;
if sComboBox6.Text<>'' then m:=m+1;
if sComboBox7.Text<>'' then m:=m+1;
if sComboBox13.Text<>'' then m:=m+1;
if sComboBox14.Text<>'' then m:=m+1;
if sComboBox15.Text<>'' then m:=m+1;
if sComboBox16.Text<>'' then m:=m+1;
if sComboBox17.Text<>'' then m:=m+1;
if sComboBox18.Text<>'' then m:=m+1;
if sComboBox19.Text<>'' then m:=m+1;
if sComboBox20.Text<>'' then m:=m+1;
if sComboBox21.Text<>'' then m:=m+1;
if sComboBox22.Text<>'' then m:=m+1;
if sComboBox23.Text<>'' then m:=m+1;
if sComboBox24.Text<>'' then m:=m+1;
if sComboBox25.Text<>'' then m:=m+1;
if sComboBox26.Text<>'' then m:=m+1;
if sComboBox27.Text<>'' then m:=m+1;
if sComboBox28.Text<>'' then m:=m+1;
if sComboBox29.Text<>'' then m:=m+1;
if sComboBox30.Text<>'' then m:=m+1;
if sComboBox31.Text<>'' then m:=m+1;
if sComboBox32.Text<>'' then m:=m+1;
if sComboBox33.Text<>'' then m:=m+1;
if sComboBox34.Text<>'' then m:=m+1;
if sComboBox35.Text<>'' then m:=m+1;
if sComboBox36.Text<>'' then m:=m+1;
if sComboBox37.Text<>'' then m:=m+1;





try
word:=CreateOleObject('Word.Application');
except
ShowMessage('Word n a pas pu etre lancé!');
Exit;
end;
while word.Templates.Count = 0 do
  Sleep(200);

word.Visible := True;
document:=word.Documents.Add;
If word.ActiveWindow.View.SplitSpecial <> 0 Then
word.ActiveWindow.Panes[2].Close;
If (word.ActiveWindow.ActivePane.View.Type = 1) Or
(word.ActiveWindow.ActivePane.View.Type = 2) Or
(word.ActiveWindow.ActivePane.View.Type = 5) Then
word.ActiveWindow.ActivePane.View.Type := 3;
word.ActiveWindow.ActivePane.View.SeekView := 9;
word.Selection.Font.Name := 'Times New Roman';
word.Selection.Font.Size := 14;
word.Selection.Font.Bold := True;
//word.Selection.TypeText(Text:='Kohlenhandel Brikett-GmbH & Co.-KG. - Holzweg 16 - 54633 Steinhausen');
If word.Selection.HeaderFooter.IsHeader = True Then
word.ActiveWindow.ActivePane.View.SeekView := 10
Else
word.ActiveWindow.ActivePane.View.SeekView :=10;


word.Selection.TypeText(Text:= '1');

word.ActiveWindow.ActivePane.View.SeekView :=0;
Word.Selection.ParagraphFormat.Alignment :=2;
word.Selection.Font.Name := 'Comic Sans MS';
word.Selection.Font.Size := 14;
//word.Selection.Font.Bold := True;

//word.Selection.TypeParagraph;
//word.Selection.TypeParagraph;
word.Selection.TypeText(Text:='Le : '+DateToStr(date())+' .');

word.Selection.TypeParagraph;
word.Selection.TypeText(Text:='Doit : '+sEdit3.Text+' .');
word.Selection.TypeParagraph;
if (sEdit4.Text<>'Adresse') then
begin
 word.Selection.Font.Size := 10;
 word.Selection.TypeText(Text:='Adresse: '+sEdit4.Text+' .');
 word.Selection.TypeParagraph;

end;
if (sEdit5.Text<>'IF') then
begin
 word.Selection.Font.Size := 10;
 word.Selection.TypeText(Text:='IF: '+sEdit5.Text+' .');
 word.Selection.TypeParagraph;
end;
if (sEdit6.Text<>'RC') then
begin
word.Selection.Font.Size := 10;
word.Selection.TypeText(Text:='RC: '+sEdit6.Text+' .');
word.Selection.TypeParagraph;

end;
 if (sEdit7.Text<>'ART') then
 begin
 word.Selection.Font.Size := 10;
 word.Selection.TypeText(Text:='ART: '+sEdit7.Text+' .');
  end;
word.Selection.TypeParagraph;
//word.Selection.TypeParagraph;
word.Selection.Font.Size := 18;
word.Selection.Font.Bold := True;
word.Selection.font.underline:=true;
Word.Selection.ParagraphFormat.Alignment :=1;

word.Selection.TypeText(Text:='Facture N°= '+inttostr(num_facture)+'/'+annee+'.');
word.Selection.TypeParagraph;
word.Selection.TypeParagraph;

aTable :=Document.Tables.Add(Word.Selection.Range,m+1,6);
Document.Tables.Item(1).Rows.Alignment := 1;

Document.Tables.Item(1).Rows.Item(1).Range.Font.Size:=10;
Document.Tables.Item(1).Rows.Item(1).Range.Font.bold:=true;
Document.Tables.Item(1).Rows.Item(1).Range.Font.underline:=false;

Document.Tables.Item(1).Rows.Item(1).Cells.Shading.BackgroundPatternColorIndex :=16;
Document.Tables.Item(1).Rows.Item(1).Cells.Shading.Texture := 75;

Document.Tables.Item(1).Columns.Item(1).SetWidth(30,0);
Document.Tables.Item(1).Columns.Item(3).SetWidth(20,0);
Document.Tables.Item(1).Columns.Item(2).SetWidth(200,0);
Document.Tables.Item(1).Columns.Item(4).SetWidth(40,0);
Document.Tables.Item(1).Columns.Item(5).SetWidth(60,0);
Document.Tables.Item(1).Columns.Item(6).SetWidth(100,0);

//Document.Tables.Item(1).Cell(1,1).borders.DefaultBorderLineWidth:=300 ;

aTable:=document.Tables.Item(1).Cell(1,1).Range.InsertAfter('N°=');
aTable:=document.Tables.Item(1).Cell(1,2).Range.InsertAfter('Désignation');
aTable:=document.Tables.Item(1).Cell(1,3).Range.InsertAfter('U');
aTable:=document.Tables.Item(1).Cell(1,4).Range.InsertAfter('Qte');
aTable:=document.Tables.Item(1).Cell(1,5).Range.InsertAfter('P.Unitaire');
aTable:=document.Tables.Item(1).Cell(1,6).Range.InsertAfter('Montant');

for i:=1 to 6 do
begin

Document.Tables.Item(1).Cell(1,i).borders.item(-2).LineStyle := 22;
Document.Tables.Item(1).Cell(1,i).borders.item(-4).LineStyle := 22;
Document.Tables.Item(1).Cell(1,i).borders.item(-1).LineStyle := 22;
Document.Tables.Item(1).Cell(1,i).borders.item(-3).LineStyle := 22;



end;
for i:=2 to m+1 do
begin
Document.Tables.Item(1).Rows.Item(i).Range.Font.Size:=12;
Document.Tables.Item(1).Rows.Item(i).Range.Font.bold:=false;
Document.Tables.Item(1).Rows.Item(i).Range.Font.underline:=false;

end;
k:=0;
if sComboBox1.Text<>'' then
begin
k:=k+1;
aTable:=document.Tables.Item(1).Cell(k+1,1).Range.InsertAfter(k);
aTable:=document.Tables.Item(1).Cell(k+1,2).Range.InsertAfter(sComboBox1.Text);
aTable:=document.Tables.Item(1).Cell(k+1,4).Range.InsertAfter(sDecimalSpinEdit3.Text);
aTable:=document.Tables.Item(1).Cell(k+1,5).Range.InsertAfter(p_u1);
aTable:=document.Tables.Item(1).Cell(k+1,6).Range.InsertAfter(sCurrencyEdit4.Text);
end;

if sComboBox2.Text<>'' then
begin
k:=k+1;
aTable:=document.Tables.Item(1).Cell(k+1,1).Range.InsertAfter(k);
aTable:=document.Tables.Item(1).Cell(k+1,2).Range.InsertAfter(sComboBox2.Text);
aTable:=document.Tables.Item(1).Cell(k+1,4).Range.InsertAfter(sDecimalSpinEdit4.Text);
aTable:=document.Tables.Item(1).Cell(k+1,5).Range.InsertAfter(inttostr(p_u2));
aTable:=document.Tables.Item(1).Cell(k+1,6).Range.InsertAfter(sCurrencyEdit5.Text);
end;

if sComboBox3.Text<>'' then
begin
k:=k+1;
aTable:=document.Tables.Item(1).Cell(k+1,1).Range.InsertAfter(k);
aTable:=document.Tables.Item(1).Cell(k+1,2).Range.InsertAfter(sComboBox3.Text);
aTable:=document.Tables.Item(1).Cell(k+1,4).Range.InsertAfter(sDecimalSpinEdit5.Text);
aTable:=document.Tables.Item(1).Cell(k+1,5).Range.InsertAfter(inttostr(p_u3));
aTable:=document.Tables.Item(1).Cell(k+1,6).Range.InsertAfter(sCurrencyEdit6.Text);
end;

if sComboBox4.Text<>'' then
begin
k:=k+1;
aTable:=document.Tables.Item(1).Cell(k+1,1).Range.InsertAfter(k);
aTable:=document.Tables.Item(1).Cell(k+1,2).Range.InsertAfter(sComboBox4.Text);
aTable:=document.Tables.Item(1).Cell(k+1,4).Range.InsertAfter(sDecimalSpinEdit6.Text);
aTable:=document.Tables.Item(1).Cell(k+1,5).Range.InsertAfter(inttostr(p_u4));
aTable:=document.Tables.Item(1).Cell(k+1,6).Range.InsertAfter(sCurrencyEdit7.Text);
end;


if sComboBox5.Text<>'' then
begin
k:=k+1;
aTable:=document.Tables.Item(1).Cell(k+1,1).Range.InsertAfter(k);
aTable:=document.Tables.Item(1).Cell(k+1,2).Range.InsertAfter(sComboBox5.Text);
aTable:=document.Tables.Item(1).Cell(k+1,4).Range.InsertAfter(sDecimalSpinEdit7.Text);
aTable:=document.Tables.Item(1).Cell(k+1,5).Range.InsertAfter(inttostr(p_u5));
aTable:=document.Tables.Item(1).Cell(k+1,6).Range.InsertAfter(sCurrencyEdit9.Text);
end;

if sComboBox6.Text<>'' then
begin
k:=k+1;
aTable:=document.Tables.Item(1).Cell(k+1,1).Range.InsertAfter(k);
aTable:=document.Tables.Item(1).Cell(k+1,2).Range.InsertAfter(sComboBox6.Text);
aTable:=document.Tables.Item(1).Cell(k+1,4).Range.InsertAfter(sDecimalSpinEdit8.Text);
aTable:=document.Tables.Item(1).Cell(k+1,5).Range.InsertAfter(inttostr(p_u6));
aTable:=document.Tables.Item(1).Cell(k+1,6).Range.InsertAfter(sCurrencyEdit10.Text);
end;

if sComboBox7.Text<>'' then
begin
k:=k+1;
aTable:=document.Tables.Item(1).Cell(k+1,1).Range.InsertAfter(k);
aTable:=document.Tables.Item(1).Cell(k+1,2).Range.InsertAfter(sComboBox7.Text);
aTable:=document.Tables.Item(1).Cell(k+1,4).Range.InsertAfter(sDecimalSpinEdit9.Text);
aTable:=document.Tables.Item(1).Cell(k+1,5).Range.InsertAfter(inttostr(p_u7));
aTable:=document.Tables.Item(1).Cell(k+1,6).Range.InsertAfter(sCurrencyEdit8.Text);
end;

if sComboBox13.Text<>'' then
begin
k:=k+1;
aTable:=document.Tables.Item(1).Cell(k+1,1).Range.InsertAfter(k);
aTable:=document.Tables.Item(1).Cell(k+1,2).Range.InsertAfter(sComboBox13.Text);
aTable:=document.Tables.Item(1).Cell(k+1,4).Range.InsertAfter(sDecimalSpinEdit11.Text);
aTable:=document.Tables.Item(1).Cell(k+1,5).Range.InsertAfter(inttostr(p_u8));
aTable:=document.Tables.Item(1).Cell(k+1,6).Range.InsertAfter(sCurrencyEdit11.Text);
end;

if sComboBox14.Text<>'' then
begin
k:=k+1;
aTable:=document.Tables.Item(1).Cell(k+1,1).Range.InsertAfter(k);
aTable:=document.Tables.Item(1).Cell(k+1,2).Range.InsertAfter(sComboBox14.Text);
aTable:=document.Tables.Item(1).Cell(k+1,4).Range.InsertAfter(sDecimalSpinEdit12.Text);
aTable:=document.Tables.Item(1).Cell(k+1,5).Range.InsertAfter(inttostr(p_u9));
aTable:=document.Tables.Item(1).Cell(k+1,6).Range.InsertAfter(sCurrencyEdit12.Text);
end;

if sComboBox15.Text<>'' then
begin
k:=k+1;
aTable:=document.Tables.Item(1).Cell(k+1,1).Range.InsertAfter(k);
aTable:=document.Tables.Item(1).Cell(k+1,2).Range.InsertAfter(sComboBox15.Text);
aTable:=document.Tables.Item(1).Cell(k+1,4).Range.InsertAfter(sDecimalSpinEdit13.Text);
aTable:=document.Tables.Item(1).Cell(k+1,5).Range.InsertAfter(inttostr(p_u10));
aTable:=document.Tables.Item(1).Cell(k+1,6).Range.InsertAfter(sCurrencyEdit13.Text);
end;

if sComboBox16.Text<>'' then
begin
k:=k+1;
aTable:=document.Tables.Item(1).Cell(k+1,1).Range.InsertAfter(k);
aTable:=document.Tables.Item(1).Cell(k+1,2).Range.InsertAfter(sComboBox16.Text);
aTable:=document.Tables.Item(1).Cell(k+1,4).Range.InsertAfter(sDecimalSpinEdit14.Text);
aTable:=document.Tables.Item(1).Cell(k+1,5).Range.InsertAfter(inttostr(p_u11));
aTable:=document.Tables.Item(1).Cell(k+1,6).Range.InsertAfter(sCurrencyEdit14.Text);
end;

if sComboBox17.Text<>'' then
begin
k:=k+1;
aTable:=document.Tables.Item(1).Cell(k+1,1).Range.InsertAfter(k);
aTable:=document.Tables.Item(1).Cell(k+1,2).Range.InsertAfter(sComboBox17.Text);
aTable:=document.Tables.Item(1).Cell(k+1,4).Range.InsertAfter(sDecimalSpinEdit15.Text);
aTable:=document.Tables.Item(1).Cell(k+1,5).Range.InsertAfter(inttostr(p_u12));
aTable:=document.Tables.Item(1).Cell(k+1,6).Range.InsertAfter(sCurrencyEdit15.Text);
end;

if sComboBox18.Text<>'' then
begin
k:=k+1;
aTable:=document.Tables.Item(1).Cell(k+1,1).Range.InsertAfter(k);
aTable:=document.Tables.Item(1).Cell(k+1,2).Range.InsertAfter(sComboBox18.Text);
aTable:=document.Tables.Item(1).Cell(k+1,4).Range.InsertAfter(sDecimalSpinEdit16.Text);
aTable:=document.Tables.Item(1).Cell(k+1,5).Range.InsertAfter(inttostr(p_u13));
aTable:=document.Tables.Item(1).Cell(k+1,6).Range.InsertAfter(sCurrencyEdit16.Text);
end;

if sComboBox19.Text<>'' then
begin
k:=k+1;
aTable:=document.Tables.Item(1).Cell(k+1,1).Range.InsertAfter(k);
aTable:=document.Tables.Item(1).Cell(k+1,2).Range.InsertAfter(sComboBox19.Text);
aTable:=document.Tables.Item(1).Cell(k+1,4).Range.InsertAfter(sDecimalSpinEdit17.Text);
aTable:=document.Tables.Item(1).Cell(k+1,5).Range.InsertAfter(inttostr(p_u14));
aTable:=document.Tables.Item(1).Cell(k+1,6).Range.InsertAfter(sCurrencyEdit17.Text);
end;

if sComboBox20.Text<>'' then
begin
k:=k+1;
aTable:=document.Tables.Item(1).Cell(k+1,1).Range.InsertAfter(k);
aTable:=document.Tables.Item(1).Cell(k+1,2).Range.InsertAfter(sComboBox20.Text);
aTable:=document.Tables.Item(1).Cell(k+1,4).Range.InsertAfter(sDecimalSpinEdit18.Text);
aTable:=document.Tables.Item(1).Cell(k+1,5).Range.InsertAfter(inttostr(p_u15));
aTable:=document.Tables.Item(1).Cell(k+1,6).Range.InsertAfter(sCurrencyEdit18.Text);
end;

if sComboBox21.Text<>'' then
begin
k:=k+1;
aTable:=document.Tables.Item(1).Cell(k+1,1).Range.InsertAfter(k);
aTable:=document.Tables.Item(1).Cell(k+1,2).Range.InsertAfter(sComboBox21.Text);
aTable:=document.Tables.Item(1).Cell(k+1,4).Range.InsertAfter(sDecimalSpinEdit19.Text);
aTable:=document.Tables.Item(1).Cell(k+1,5).Range.InsertAfter(inttostr(p_u16));
aTable:=document.Tables.Item(1).Cell(k+1,6).Range.InsertAfter(sCurrencyEdit19.Text);
end;

if sComboBox22.Text<>'' then
begin
k:=k+1;
aTable:=document.Tables.Item(1).Cell(k+1,1).Range.InsertAfter(k);
aTable:=document.Tables.Item(1).Cell(k+1,2).Range.InsertAfter(sComboBox22.Text);
aTable:=document.Tables.Item(1).Cell(k+1,4).Range.InsertAfter(sDecimalSpinEdit20.Text);
aTable:=document.Tables.Item(1).Cell(k+1,5).Range.InsertAfter(inttostr(p_u17));
aTable:=document.Tables.Item(1).Cell(k+1,6).Range.InsertAfter(sCurrencyEdit20.Text);
end;

if sComboBox23.Text<>'' then
begin
k:=k+1;
aTable:=document.Tables.Item(1).Cell(k+1,1).Range.InsertAfter(k);
aTable:=document.Tables.Item(1).Cell(k+1,2).Range.InsertAfter(sComboBox23.Text);
aTable:=document.Tables.Item(1).Cell(k+1,4).Range.InsertAfter(sDecimalSpinEdit21.Text);
aTable:=document.Tables.Item(1).Cell(k+1,5).Range.InsertAfter(inttostr(p_u18));
aTable:=document.Tables.Item(1).Cell(k+1,6).Range.InsertAfter(sCurrencyEdit21.Text);
end;


if sComboBox24.Text<>'' then
begin
k:=k+1;
aTable:=document.Tables.Item(1).Cell(k+1,1).Range.InsertAfter(k);
aTable:=document.Tables.Item(1).Cell(k+1,2).Range.InsertAfter(sComboBox24.Text);
aTable:=document.Tables.Item(1).Cell(k+1,4).Range.InsertAfter(sDecimalSpinEdit22.Text);
aTable:=document.Tables.Item(1).Cell(k+1,5).Range.InsertAfter(inttostr(p_u19));
aTable:=document.Tables.Item(1).Cell(k+1,6).Range.InsertAfter(sCurrencyEdit22.Text);
end;

if sComboBox25.Text<>'' then
begin
k:=k+1;
aTable:=document.Tables.Item(1).Cell(k+1,1).Range.InsertAfter(k);
aTable:=document.Tables.Item(1).Cell(k+1,2).Range.InsertAfter(sComboBox25.Text);
aTable:=document.Tables.Item(1).Cell(k+1,4).Range.InsertAfter(sDecimalSpinEdit23.Text);
aTable:=document.Tables.Item(1).Cell(k+1,5).Range.InsertAfter(inttostr(p_u20));
aTable:=document.Tables.Item(1).Cell(k+1,6).Range.InsertAfter(sCurrencyEdit23.Text);
end;


if sComboBox26.Text<>'' then
begin
k:=k+1;
aTable:=document.Tables.Item(1).Cell(k+1,1).Range.InsertAfter(k);
aTable:=document.Tables.Item(1).Cell(k+1,2).Range.InsertAfter(sComboBox26.Text);
aTable:=document.Tables.Item(1).Cell(k+1,4).Range.InsertAfter(sDecimalSpinEdit24.Text);
aTable:=document.Tables.Item(1).Cell(k+1,5).Range.InsertAfter(inttostr(p_u21));
aTable:=document.Tables.Item(1).Cell(k+1,6).Range.InsertAfter(sCurrencyEdit24.Text);
end;


if sComboBox27.Text<>'' then
begin
k:=k+1;
aTable:=document.Tables.Item(1).Cell(k+1,1).Range.InsertAfter(k);
aTable:=document.Tables.Item(1).Cell(k+1,2).Range.InsertAfter(sComboBox27.Text);
aTable:=document.Tables.Item(1).Cell(k+1,4).Range.InsertAfter(sDecimalSpinEdit25.Text);
aTable:=document.Tables.Item(1).Cell(k+1,5).Range.InsertAfter(inttostr(p_u22));
aTable:=document.Tables.Item(1).Cell(k+1,6).Range.InsertAfter(sCurrencyEdit25.Text);
end;

if sComboBox28.Text<>'' then
begin
k:=k+1;
aTable:=document.Tables.Item(1).Cell(k+1,1).Range.InsertAfter(k);
aTable:=document.Tables.Item(1).Cell(k+1,2).Range.InsertAfter(sComboBox28.Text);
aTable:=document.Tables.Item(1).Cell(k+1,4).Range.InsertAfter(sDecimalSpinEdit26.Text);
aTable:=document.Tables.Item(1).Cell(k+1,5).Range.InsertAfter(inttostr(p_u23));
aTable:=document.Tables.Item(1).Cell(k+1,6).Range.InsertAfter(sCurrencyEdit26.Text);
end;

if sComboBox29.Text<>'' then
begin
k:=k+1;
aTable:=document.Tables.Item(1).Cell(k+1,1).Range.InsertAfter(k);
aTable:=document.Tables.Item(1).Cell(k+1,2).Range.InsertAfter(sComboBox29.Text);
aTable:=document.Tables.Item(1).Cell(k+1,4).Range.InsertAfter(sDecimalSpinEdit27.Text);
aTable:=document.Tables.Item(1).Cell(k+1,5).Range.InsertAfter(inttostr(p_u24));
aTable:=document.Tables.Item(1).Cell(k+1,6).Range.InsertAfter(sCurrencyEdit27.Text);
end;

if sComboBox30.Text<>'' then
begin
k:=k+1;
aTable:=document.Tables.Item(1).Cell(k+1,1).Range.InsertAfter(k);
aTable:=document.Tables.Item(1).Cell(k+1,2).Range.InsertAfter(sComboBox30.Text);
aTable:=document.Tables.Item(1).Cell(k+1,4).Range.InsertAfter(sDecimalSpinEdit28.Text);
aTable:=document.Tables.Item(1).Cell(k+1,5).Range.InsertAfter(inttostr(p_u25));
aTable:=document.Tables.Item(1).Cell(k+1,6).Range.InsertAfter(sCurrencyEdit28.Text);
end;


if sComboBox31.Text<>'' then
begin
k:=k+1;
aTable:=document.Tables.Item(1).Cell(k+1,1).Range.InsertAfter(k);
aTable:=document.Tables.Item(1).Cell(k+1,2).Range.InsertAfter(sComboBox31.Text);
aTable:=document.Tables.Item(1).Cell(k+1,4).Range.InsertAfter(sDecimalSpinEdit29.Text);
aTable:=document.Tables.Item(1).Cell(k+1,5).Range.InsertAfter(inttostr(p_u26));
aTable:=document.Tables.Item(1).Cell(k+1,6).Range.InsertAfter(sCurrencyEdit29.Text);
end;


if sComboBox32.Text<>'' then
begin
k:=k+1;
aTable:=document.Tables.Item(1).Cell(k+1,1).Range.InsertAfter(k);
aTable:=document.Tables.Item(1).Cell(k+1,2).Range.InsertAfter(sComboBox32.Text);
aTable:=document.Tables.Item(1).Cell(k+1,4).Range.InsertAfter(sDecimalSpinEdit30.Text);
aTable:=document.Tables.Item(1).Cell(k+1,5).Range.InsertAfter(inttostr(p_u27));
aTable:=document.Tables.Item(1).Cell(k+1,6).Range.InsertAfter(sCurrencyEdit30.Text);
end;

if sComboBox33.Text<>'' then
begin
k:=k+1;
aTable:=document.Tables.Item(1).Cell(k+1,1).Range.InsertAfter(k);
aTable:=document.Tables.Item(1).Cell(k+1,2).Range.InsertAfter(sComboBox33.Text);
aTable:=document.Tables.Item(1).Cell(k+1,4).Range.InsertAfter(sDecimalSpinEdit31.Text);
aTable:=document.Tables.Item(1).Cell(k+1,5).Range.InsertAfter(inttostr(p_u28));
aTable:=document.Tables.Item(1).Cell(k+1,6).Range.InsertAfter(sCurrencyEdit31.Text);
end;

if sComboBox34.Text<>'' then
begin
k:=k+1;
aTable:=document.Tables.Item(1).Cell(k+1,1).Range.InsertAfter(k);
aTable:=document.Tables.Item(1).Cell(k+1,2).Range.InsertAfter(sComboBox34.Text);
aTable:=document.Tables.Item(1).Cell(k+1,4).Range.InsertAfter(sDecimalSpinEdit32.Text);
aTable:=document.Tables.Item(1).Cell(k+1,5).Range.InsertAfter(inttostr(p_u29));
aTable:=document.Tables.Item(1).Cell(k+1,6).Range.InsertAfter(sCurrencyEdit32.Text);
end;

if sComboBox35.Text<>'' then
begin
k:=k+1;
aTable:=document.Tables.Item(1).Cell(k+1,1).Range.InsertAfter(k);
aTable:=document.Tables.Item(1).Cell(k+1,2).Range.InsertAfter(sComboBox35.Text);
aTable:=document.Tables.Item(1).Cell(k+1,4).Range.InsertAfter(sDecimalSpinEdit33.Text);
aTable:=document.Tables.Item(1).Cell(k+1,5).Range.InsertAfter(inttostr(p_u30));
aTable:=document.Tables.Item(1).Cell(k+1,6).Range.InsertAfter(sCurrencyEdit33.Text);
end;

if sComboBox36.Text<>'' then
begin
k:=k+1;
aTable:=document.Tables.Item(1).Cell(k+1,1).Range.InsertAfter(k);
aTable:=document.Tables.Item(1).Cell(k+1,2).Range.InsertAfter(sComboBox36.Text);
aTable:=document.Tables.Item(1).Cell(k+1,4).Range.InsertAfter(sDecimalSpinEdit34.Text);
aTable:=document.Tables.Item(1).Cell(k+1,5).Range.InsertAfter(inttostr(p_u31));
aTable:=document.Tables.Item(1).Cell(k+1,6).Range.InsertAfter(sCurrencyEdit34.Text);
end;

if sComboBox37.Text<>'' then
begin
k:=k+1;
aTable:=document.Tables.Item(1).Cell(k+1,1).Range.InsertAfter(k);
aTable:=document.Tables.Item(1).Cell(k+1,2).Range.InsertAfter(sComboBox37.Text);
aTable:=document.Tables.Item(1).Cell(k+1,4).Range.InsertAfter(sDecimalSpinEdit35.Text);
aTable:=document.Tables.Item(1).Cell(k+1,5).Range.InsertAfter(inttostr(p_u32));
aTable:=document.Tables.Item(1).Cell(k+1,6).Range.InsertAfter(sCurrencyEdit35.Text);
end;




Document.Tables.item(1).borders.item(-2).LineStyle := 22;
Document.Tables.item(1).borders.item(-4).LineStyle := 22;
Document.Tables.item(1).borders.item(-3).LineStyle := 22;




Word.Selection.GoTo(3,-1);
word.Selection.TypeParagraph;

bTable :=Document.Tables.Add(Word.Selection.Range, 3,2);

for i :=1 to 3 do
begin
Document.Tables.Item(2).Cell(i,1).Range.Font.Size:=14;
Document.Tables.Item(2).Cell(i,1).Range.Font.bold:=true;
Document.Tables.Item(2).Cell(i,1).Range.Font.underline:=false;

Document.Tables.Item(2).Cell(i,2).Range.Font.Size:=14;
Document.Tables.Item(2).Cell(i,2).Range.Font.bold:=true;
Document.Tables.Item(2).Cell(i,2).Range.Font.underline:=false;

end;





bTable:=document.Tables.Item(2).Cell(1,1).Range.InsertAfter('Montant HT ');
bTable:=document.Tables.Item(2).Cell(2,1).Range.InsertAfter('TVA 17%');
bTable:=document.Tables.Item(2).Cell(3,1).Range.InsertAfter('Montant TTC');


bTable:=document.Tables.Item(2).Cell(1,2).Range.InsertAfter(floattostr(s));
bTable:=document.Tables.Item(2).Cell(2,2).Range.InsertAfter(floattostr(s*0.17));
bTable:=document.Tables.Item(2).Cell(3,2).Range.InsertAfter(floattostr(s+(s*0.17)));

for i:=1 to 3 do
begin

Document.Tables.Item(2).Cell(i,1).borders.item(-2).LineStyle := 22;
Document.Tables.Item(2).Cell(i,1).borders.item(-4).LineStyle := 22;
Document.Tables.Item(2).Cell(i,1).borders.item(-1).LineStyle := 22;
Document.Tables.Item(2).Cell(i,1).borders.item(-3).LineStyle := 22;

Document.Tables.Item(2).Cell(i,2).borders.item(-2).LineStyle := 22;
Document.Tables.Item(2).Cell(i,2).borders.item(-4).LineStyle := 22;
Document.Tables.Item(2).Cell(i,2).borders.item(-1).LineStyle := 22;
Document.Tables.Item(2).Cell(i,2).borders.item(-3).LineStyle := 22;

end;


Document.Tables.Item(2).Columns.Item(1).SetWidth(350,0);
Document.Tables.Item(2).Columns.Item(2).SetWidth(100,0);


Word.Selection.GoTo(3,-1);
h:=floattostr(s+(s*0.17));

x:=0;// connaitre la position du separateur decimale
while not((h[x]=',') or(x > length(h))) do
   begin
     x:=x+1;
    end;
 C:=copy(h,x,1);
 if C=',' then // s'il y a un nombre decimale
    begin
     D:=copy(h,1,(x-1));
     dinars:=enLettres(StrToInt(d));
     centimes:=copy(h, x+1, (Length(h)-x));
     // remplir un tableau car par des chaines vides
     for y:=1 to 3 do     // Le nombre de Zero permit et de 3 maxi !
      begin
       car[y]:='';
      end;
     y:=0; // Y represente le nombre de Zero apres le separateur decimale
     z:=x+1;
     while (h[z]='0') or(z > length(h)) do
      begin
       y:=y+1;
       z:=z+1;
       car[y]:='Zero ';
      end;
     //s'il y a des Zero apres le separateur decimale
     if y>0 then
      // il faut l'ecrire -- maxi 3 nombres ont la valeur egal à 0, soit permit
      centimes:=' et '+car[1]+car[2]+car[3]+enLettres(StrToInt(centimes))+' Centimes'
      else // sinon il n'y a pas de zero à ecrire
        centimes:=' et'+enLettres(StrToInt(centimes))+' Centimes';
      chiffre_lettre:=dinars+' Dinars Algerien'+centimes;
    end
   else // sinon lire la partie entiere
    begin
    D:=copy(h,1,(length(h)));
    dinars:=enLettres(StrToInt(D));
    chiffre_lettre:=dinars+' Dinars Algérien';
    end; // else



word.Selection.Font.Size := 12;
word.Selection.Font.Bold := true;
word.Selection.font.underline:=false;
Word.Selection.ParagraphFormat.Alignment :=3;
word.Selection.TypeText(Text:='Arrêté la présente facture à la somme de : ');
word.Selection.TypeParagraph;
word.Selection.Font.Bold := false;
word.Selection.TypeText(Text:=chiffre_lettre);

word.Selection.TypeParagraph;

word.Selection.font.underline:=true;
word.Selection.Font.Bold := false;

Word.Selection.ParagraphFormat.Alignment :=2;
word.Selection.TypeText(Text:='Cachet & Signature : ');




end;
//end;

sPanel4.Visible:=true;
sGroupBox8.Visible:=true;
sPanel2.Visible:=false;
sGroupBox5.Visible:=false;

sPanel3.Visible:=false;
sGroupBox6.Visible:=false;

sGroupBox9.Visible:=true;

sCurrencyEdit4.Text:='0';
sCurrencyEdit5.Text:='0';
sCurrencyEdit6.Text:='0';
sCurrencyEdit7.Text:='0';
sCurrencyEdit8.Text:='0';
sCurrencyEdit9.Text:='0';
sCurrencyEdit10.Text:='0';

sCurrencyEdit11.Text:='0';
sCurrencyEdit12.Text:='0';
sCurrencyEdit13.Text:='0';
sCurrencyEdit14.Text:='0';
sCurrencyEdit15.Text:='0';
sCurrencyEdit16.Text:='0';
sCurrencyEdit17.Text:='0';

sCurrencyEdit18.Text:='0';
sCurrencyEdit19.Text:='0';
sCurrencyEdit20.Text:='0';
sCurrencyEdit21.Text:='0';
sCurrencyEdit22.Text:='0';
sCurrencyEdit23.Text:='0';
sCurrencyEdit24.Text:='0';
sCurrencyEdit25.Text:='0';
sCurrencyEdit26.Text:='0';
sCurrencyEdit27.Text:='0';
sCurrencyEdit28.Text:='0';
sCurrencyEdit29.Text:='0';
sCurrencyEdit30.Text:='0';
sCurrencyEdit31.Text:='0';
sCurrencyEdit32.Text:='0';
sCurrencyEdit33.Text:='0';




sDecimalSpinEdit3.Text:='1';
sDecimalSpinEdit4.Text:='1';
sDecimalSpinEdit5.Text:='1';
sDecimalSpinEdit6.Text:='1';
sDecimalSpinEdit7.Text:='1';
sDecimalSpinEdit8.Text:='1';
sDecimalSpinEdit9.Text:='1';
sDecimalSpinEdit10.Text:='1';
sDecimalSpinEdit11.Text:='1';
sDecimalSpinEdit12.Text:='1';
sDecimalSpinEdit13.Text:='1';
sDecimalSpinEdit14.Text:='1';
sDecimalSpinEdit15.Text:='1';
sDecimalSpinEdit16.Text:='1';
sDecimalSpinEdit17.Text:='1';
sDecimalSpinEdit18.Text:='1';
sDecimalSpinEdit19.Text:='1';
sDecimalSpinEdit20.Text:='1';
sDecimalSpinEdit21.Text:='1';
sDecimalSpinEdit22.Text:='1';
sDecimalSpinEdit23.Text:='1';
sDecimalSpinEdit24.Text:='1';
sDecimalSpinEdit25.Text:='1';
sDecimalSpinEdit26.Text:='1';
sDecimalSpinEdit27.Text:='1';
sDecimalSpinEdit28.Text:='1';
sDecimalSpinEdit29.Text:='1';
sDecimalSpinEdit30.Text:='1';
sDecimalSpinEdit31.Text:='1';
sDecimalSpinEdit32.Text:='1';
sDecimalSpinEdit33.Text:='1';




sComboBox1.Text:='';
sComboBox2.Text:='';
sComboBox3.Text:='';
sComboBox4.Text:='';
sComboBox5.Text:='';
sComboBox6.Text:='';
sComboBox7.Text:='';

sComboBox13.Text:='';
sComboBox14.Text:='';
sComboBox15.Text:='';
sComboBox16.Text:='';
sComboBox17.Text:='';
sComboBox18.Text:='';
sComboBox19.Text:='';
sComboBox20.Text:='';
sComboBox21.Text:='';
sComboBox22.Text:='';
sComboBox23.Text:='';
sComboBox24.Text:='';
sComboBox25.Text:='';
sComboBox26.Text:='';
sComboBox27.Text:='';
sComboBox28.Text:='';
sComboBox29.Text:='';
sComboBox30.Text:='';
sComboBox31.Text:='';
sComboBox32.Text:='';
sComboBox33.Text:='';
sComboBox34.Text:='';
sComboBox35.Text:='';


sLabel12.Caption:='                   DA';
sEdit3.Text:='Nom';








end;

end;



procedure TForm1.sComboBox1Click(Sender: TObject);
var i,j:integer;
begin
DataModule2.ADOTable1.Close;
datamodule2.DataSource1.DataSet:=datamodule2.ADOTable1;

DataModule2.ADOTable1.Open;
i:=datamodule2.ADOTable1.RecordCount;
//showmessage (inttostr(i));
for j := 1 to i do
begin
with DataModule2.ADOTable1 do
begin
if (sComboBox1.Items[scombobox1.itemIndex]=datamodule2.ADOTable1.FieldByName('nom_produit').AsString) then
begin

if (strtoint(sDecimalSpinEdit3.Text)>strtoint(datamodule2.ADOTable1.Fields[3].AsString)) then
begin
ShowMessage('Erreur !! La quantité voulue dépasse la quantité stockée, Le max est :'+datamodule2.ADOTable1.Fields[3].AsString);
sDecimalSpinEdit3.Text:=datamodule2.ADOTable1.Fields[3].AsString;
end;
sCurrencyEdit4.Text:=IntToStr(strtoint(datamodule2.ADOTable1.Fields[2].AsString)*strtoint(sDecimalSpinEdit3.Text));


p_u1 :=strtoint(datamodule2.ADOTable1.Fields[2].AsString);
//showmessage (inttostr(p_u1));
end;
end;
datamodule2.ADOTable1.Next;
end;
end;

procedure TForm1.sComboBox2Click(Sender: TObject);
var i,j:integer;
begin

DataModule2.ADOTable1.Close;
datamodule2.DataSource1.DataSet:=datamodule2.ADOTable1;
DataModule2.ADOTable1.Open;
i:=datamodule2.ADOTable1.RecordCount;

for j := 1 to i do
begin
with DataModule2.ADOTable1 do
begin
if (sComboBox2.Items[scombobox2.itemIndex]=datamodule2.ADOTable1.FieldByName('nom_produit').AsString) then
begin
if (strtoint(sDecimalSpinEdit4.Text)>strtoint(datamodule2.ADOTable1.Fields[3].AsString)) then
begin
ShowMessage('Erreur !! La quantité voulue dépasse la quantité stockée, Le max est :'+datamodule2.ADOTable1.Fields[3].AsString);
sDecimalSpinEdit4.Text:=datamodule2.ADOTable1.Fields[3].AsString;
end;
sCurrencyEdit5.Text:=IntToStr(strtoint(datamodule2.ADOTable1.Fields[2].AsString)*strtoint(sDecimalSpinEdit4.Text))

end;
p_u2 :=strtoint(datamodule2.ADOTable1.Fields[2].AsString);
//showmessage('pu 2'+inttostr(p_u2));


end;
datamodule2.ADOTable1.Next;
end;
end;


procedure TForm1.sComboBox3Click(Sender: TObject);
var i,j:integer;
begin

DataModule2.ADOTable1.Close;
datamodule2.DataSource1.DataSet:=datamodule2.ADOTable1;
DataModule2.ADOTable1.Open;
i:=datamodule2.ADOTable1.RecordCount;

for j := 1 to i do
begin
with DataModule2.ADOTable1 do
begin
if (sComboBox3.Items[scombobox3.itemIndex]=datamodule2.ADOTable1.FieldByName('nom_produit').AsString) then
begin
if (strtoint(sDecimalSpinEdit5.Text)>strtoint(datamodule2.ADOTable1.Fields[3].AsString)) then
begin
ShowMessage('Erreur !! La quantité voulue dépasse la quantité stockée, Le max est :'+datamodule2.ADOTable1.Fields[3].AsString);
sDecimalSpinEdit5.Text:=datamodule2.ADOTable1.Fields[3].AsString;

end;
sCurrencyEdit6.Text:=IntToStr(strtoint(datamodule2.ADOTable1.Fields[2].AsString)*strtoint(sDecimalSpinEdit5.Text))

end;
p_u3 :=strtoint(datamodule2.ADOTable1.Fields[2].AsString);

end;
datamodule2.ADOTable1.Next;
end;

end;

procedure TForm1.sComboBox4Click(Sender: TObject);
var i,j:integer;
begin

DataModule2.ADOTable1.Close;
datamodule2.DataSource1.DataSet:=datamodule2.ADOTable1;
DataModule2.ADOTable1.Open;
i:=datamodule2.ADOTable1.RecordCount;

for j := 1 to i do
begin
with DataModule2.ADOTable1 do
begin
if (sComboBox4.Items[scombobox4.itemIndex]=datamodule2.ADOTable1.FieldByName('nom_produit').AsString) then
begin

if (strtoint(sDecimalSpinEdit6.Text)>strtoint(datamodule2.ADOTable1.Fields[3].AsString)) then
begin
ShowMessage('Erreur !! La quantité voulue dépasse la quantité stockée, Le max est :'+datamodule2.ADOTable1.Fields[3].AsString);
sDecimalSpinEdit6.Text:=datamodule2.ADOTable1.Fields[3].AsString;

end;
sCurrencyEdit8.Text:=IntToStr(strtoint(datamodule2.ADOTable1.Fields[2].AsString)*strtoint(sDecimalSpinEdit6.Text))
end;
p_u4 :=strtoint(datamodule2.ADOTable1.Fields[2].AsString);

end;
datamodule2.ADOTable1.Next;
end;
end;

procedure TForm1.sComboBox5Click(Sender: TObject);
var i,j:integer;
begin

DataModule2.ADOTable1.Close;
datamodule2.DataSource1.DataSet:=datamodule2.ADOTable1;
DataModule2.ADOTable1.Open;
i:=datamodule2.ADOTable1.RecordCount;

for j := 1 to i do
begin
with DataModule2.ADOTable1 do
begin
if (sComboBox5.Items[scombobox5.itemIndex]=datamodule2.ADOTable1.FieldByName('nom_produit').AsString) then
begin
if (strtoint(sDecimalSpinEdit7.Text)>strtoint(datamodule2.ADOTable1.Fields[3].AsString)) then
begin
ShowMessage('Erreur !! La quantité voulue dépasse la quantité stockée, Le max est :'+datamodule2.ADOTable1.Fields[3].AsString);
sDecimalSpinEdit7.Text:=datamodule2.ADOTable1.Fields[3].AsString;

end;

sCurrencyEdit9.Text:=IntToStr(strtoint(datamodule2.ADOTable1.Fields[2].AsString)*strtoint(sDecimalSpinEdit7.Text))
end;
p_u5 :=strtoint(datamodule2.ADOTable1.Fields[2].AsString);

end;
datamodule2.ADOTable1.Next;
end;
end;









procedure TForm1.sComboBox6Click(Sender: TObject);
var i,j:integer;
begin

DataModule2.ADOTable1.Close;
datamodule2.DataSource1.DataSet:=datamodule2.ADOTable1;

DataModule2.ADOTable1.Open;
i:=datamodule2.ADOTable1.RecordCount;

for j := 1 to i do
begin
with DataModule2.ADOTable1 do
begin
if (sComboBox6.Items[scombobox6.itemIndex]=datamodule2.ADOTable1.FieldByName('nom_produit').AsString) then
begin
if (strtoint(sDecimalSpinEdit8.Text)>strtoint(datamodule2.ADOTable1.Fields[3].AsString)) then
begin
ShowMessage('Erreur !! La quantité voulue dépasse la quantité stockée, Le max est :'+datamodule2.ADOTable1.Fields[3].AsString);
sDecimalSpinEdit8.Text:=datamodule2.ADOTable1.Fields[3].AsString;

end;
sCurrencyEdit10.Text:=IntToStr(strtoint(datamodule2.ADOTable1.Fields[2].AsString)*strtoint(sDecimalSpinEdit8.Text))
end;
p_u6 :=strtoint(datamodule2.ADOTable1.Fields[2].AsString);

end;
datamodule2.ADOTable1.Next;
end;
end;

procedure TForm1.sComboBox8Click(Sender: TObject);
var i,j:integer;
begin

DataModule2.ADOTable1.Close;
datamodule2.DataSource1.DataSet:=datamodule2.ADOTable1;

DataModule2.ADOTable1.Open;
i:=datamodule2.ADOTable1.RecordCount;
for j := 1 to i do
begin
with DataModule2.ADOTable1 do
begin

if (sComboBox7.Items[scombobox7.itemIndex]=datamodule2.ADOTable1.FieldByName('nom_produit').AsString) then
begin
if (strtoint(sDecimalSpinEdit9.Text)>strtoint(datamodule2.ADOTable1.Fields[3].AsString)) then
begin
ShowMessage('Erreur !! La quantité voulue dépasse la quantité stockée, Le max est :'+datamodule2.ADOTable1.Fields[3].AsString);
sDecimalSpinEdit9.Text:=datamodule2.ADOTable1.Fields[3].AsString;

end;
sCurrencyEdit7.Text:=IntToStr(strtoint(datamodule2.ADOTable1.Fields[2].AsString)*strtoint(sDecimalSpinEdit9.Text))
end;
p_u7 :=strtoint(datamodule2.ADOTable1.Fields[2].AsString);

end;
datamodule2.ADOTable1.Next;
end;
end;

procedure TForm1.sLabel11Click(Sender: TObject);
begin
sLabel12.Caption:=inttostr(strtoint(sCurrencyEdit4.Text)+strtoint(sCurrencyEdit5.Text)+strtoint(sCurrencyEdit6.Text)+strtoint(sCurrencyEdit7.Text)+strtoint(sCurrencyEdit8.Text)+strtoint(sCurrencyEdit9.Text)+strtoint(sCurrencyEdit10.Text)+strtoint(sCurrencyEdit11.Text)+strtoint(sCurrencyEdit12.Text)+strtoint(sCurrencyEdit13.Text)+strtoint(sCurrencyEdit14.Text)+strtoint(sCurrencyEdit15.Text)+strtoint(sCurrencyEdit16.Text)+strtoint(sCurrencyEdit17.Text)+strtoint(sCurrencyEdit18.Text)+strtoint(sCurrencyEdit19.Text)+strtoint(sCurrencyEdit20.Text)+strtoint(sCurrencyEdit21.Text)+strtoint(sCurrencyEdit22.Text)+strtoint(sCurrencyEdit23.Text)+strtoint(sCurrencyEdit24.Text)+strtoint(sCurrencyEdit25.Text)+strtoint(sCurrencyEdit26.Text)+strtoint(sCurrencyEdit27.Text)+strtoint(sCurrencyEdit28.Text)+strtoint(sCurrencyEdit29.Text)+strtoint(sCurrencyEdit30.Text)+strtoint(sCurrencyEdit31.Text)+strtoint(sCurrencyEdit32.Text)+strtoint(sCurrencyEdit33.Text))+'DA';
end;

procedure TForm1.sLabel12Click(Sender: TObject);
begin
sLabel12.Caption:=inttostr(strtoint(sCurrencyEdit4.Text)+strtoint(sCurrencyEdit5.Text)+strtoint(sCurrencyEdit6.Text)+strtoint(sCurrencyEdit7.Text)+strtoint(sCurrencyEdit8.Text)+strtoint(sCurrencyEdit9.Text)+strtoint(sCurrencyEdit10.Text)+strtoint(sCurrencyEdit11.Text)+strtoint(sCurrencyEdit12.Text)+strtoint(sCurrencyEdit13.Text)+strtoint(sCurrencyEdit14.Text)+strtoint(sCurrencyEdit15.Text)+strtoint(sCurrencyEdit16.Text)+strtoint(sCurrencyEdit17.Text)+strtoint(sCurrencyEdit18.Text)+strtoint(sCurrencyEdit19.Text)+strtoint(sCurrencyEdit20.Text)+strtoint(sCurrencyEdit21.Text)+strtoint(sCurrencyEdit22.Text)+strtoint(sCurrencyEdit23.Text)+strtoint(sCurrencyEdit24.Text)+strtoint(sCurrencyEdit25.Text)+strtoint(sCurrencyEdit26.Text)+strtoint(sCurrencyEdit27.Text)+strtoint(sCurrencyEdit28.Text)+strtoint(sCurrencyEdit29.Text)+strtoint(sCurrencyEdit30.Text)+strtoint(sCurrencyEdit31.Text)+strtoint(sCurrencyEdit32.Text)+strtoint(sCurrencyEdit33.Text))+'DA';

end;





procedure TForm1.sSpeedButton5Click(Sender: TObject);
begin
sPanel4.Visible:=false;
sGroupBox8.Visible:=false;
sPanel2.Visible:=false;
sGroupBox5.Visible:=false;

sPanel3.Visible:=false;
sGroupBox6.Visible:=false;
sGroupBox10.Visible:=false;
sComboBox8.Text:='';
sComboBox9.Text:='';
sComboBox10.Text:='';
sComboBox11.Text:='';
sComboBox12.Text:='';

end;

procedure TForm1.P2Click(Sender: TObject);
var d : TDateTime;
    jour,mois,annee,jour_nom:string;

begin
sPanel4.Visible:=false;
sGroupBox8.Visible:=false;
sPanel2.Visible:=false;
sGroupBox5.Visible:=false;
spanel3.Visible:=false;
sGroupBox6.Visible:=false;
sPanel3.Visible:=false;
sGroupBox6.Visible:=false;
sGroupBox10.Visible:=true;

datamodule2.ADOTable1.close;
datamodule2.ADOQuery_rech_produit_date.close;
//datamodule2.DataSource1.DataSet:=datamodule2.ADOQuery_rech_produit_date;
datamodule2.ADOQuery_rech_produit_date.Parameters.ParamByName('date').Value:=date();
//datamodule2.DataSource_compte_admin.DataSet:=datamodule10.ADOQuery_authentification;
DataModule2.adoquery_rech_produit_date.Open;
if (datamodule2.adoquery_rech_produit_date.RecordCount>=1) then
begin
while not (datamodule2.ADOQuery_rech_produit_date.Eof) do
begin
sComboBox8.Items.Add(datamodule2.ADOQuery_rech_produit_date.Fields[1].Value);
datamodule2.ADOQuery_rech_produit_date.next;
end;

end
else
sComboBox8.Items.Add('Aucun Produit ajouté');


DataModule2.adoquery_rech_produit_date.Close;
DataModule2.ADOQuery_rech_modif_date.Close;
//datamodule2.DataSource1.DataSet:=datamodule2.ADOQuery_rech_modif_date;

datamodule2.ADOQuery_rech_modif_date.Parameters.ParamByName('date').Value:=date();

DataModule2.ADOQuery_rech_modif_date.Open;

if (datamodule2.ADOQuery_rech_modif_date.RecordCount>=1) then
begin
while not (datamodule2.ADOQuery_rech_modif_date.Eof) do
begin
sComboBox9.Items.Add(datamodule2.ADOQuery_rech_modif_date.Fields[1].Value);
datamodule2.ADOQuery_rech_modif_date.next;
end;

end
else
sComboBox9.Items.Add('Aucun Produit modifié');

DataModule2.ADOQuery_rech_modif_date.Close;

DataModule2.ADOQuery_date_semaine.Close;
//datamodule2.DataSource1.DataSet:=datamodule2.ADOQuery_date_semaine;


jour :=copy(datetostr(date),1,2);
mois:=copy(datetostr(date),4,2);
annee:=copy(datetostr(date),7,4);
d:=EncodeDate(StrToInt(annee),StrToInt(mois),StrToInt(jour));
//showmessage(datetostr(d));

jour_nom :=longdaynames[dayofweek(d)];


if (jour_nom='dimanche') then
begin
//ShowMessage('yassine');
//datamodule2.ADOQuery_date_semaine.Parameters.ParamByName('date1').DataType:=TDate;
datamodule2.ADOQuery_date_semaine.Parameters.ParamByName('date1').Value:=date;

//d:=EncodeDate(StrToInt(annee),StrToInt(mois),StrToInt(jour)+6);
//datamodule2.ADOQuery_date_semaine.Parameters.ParamByName('date2').DataType:=TDateField;

datamodule2.ADOQuery_date_semaine.Parameters.ParamByName('date2').Value:=date;
//showmessage(datamodule2.ADOQuery_date_semaine.Parameters.ParamByName('date2').Value);

end;
if (jour_nom='lundi') then
begin
datamodule2.ADOQuery_date_semaine.Parameters.ParamByName('date1').Value:=Yesterday;

datamodule2.ADOQuery_date_semaine.Parameters.ParamByName('date2').Value:=date;
end;

if (jour_nom='mardi') then
begin

d:=EncodeDate(StrToInt(annee),StrToInt(mois),StrToInt(jour)-2);

datamodule2.ADOQuery_date_semaine.Parameters.ParamByName('date1').Value:=d;

datamodule2.ADOQuery_date_semaine.Parameters.ParamByName('date2').Value:=date;
end;

if (jour_nom='mercredi') then
begin

d:=EncodeDate(StrToInt(annee),StrToInt(mois),StrToInt(jour)-3);
datamodule2.ADOQuery_date_semaine.Parameters.ParamByName('date1').Value:=d;

datamodule2.ADOQuery_date_semaine.Parameters.ParamByName('date2').Value:=date;
end;
if (jour_nom='jeudi') then
begin
d:=EncodeDate(StrToInt(annee),StrToInt(mois),StrToInt(jour)-4);
datamodule2.ADOQuery_date_semaine.Parameters.ParamByName('date1').Value:=d;

datamodule2.ADOQuery_date_semaine.Parameters.ParamByName('date2').Value:=date;
end;
if (jour_nom='vendredi') then
begin
d:=EncodeDate(StrToInt(annee),StrToInt(mois),StrToInt(jour)-5);
datamodule2.ADOQuery_date_semaine.Parameters.ParamByName('date1').Value:=d;

datamodule2.ADOQuery_date_semaine.Parameters.ParamByName('date2').Value:=date;
end;
if (jour_nom='samedi') then
begin
d:=EncodeDate(StrToInt(annee),StrToInt(mois),StrToInt(jour)-6);
datamodule2.ADOQuery_date_semaine.Parameters.ParamByName('date1').Value:=d;

datamodule2.ADOQuery_date_semaine.Parameters.ParamByName('date2').Value:=date;
end;

DataModule2.ADOQuery_date_semaine.open;
//showmessage (inttostr(datamodule2.ADOQuery_date_semaine.RecordCount));
if (datamodule2.ADOQuery_date_semaine.RecordCount>=1) then
begin
while not (datamodule2.ADOQuery_date_semaine.Eof) do
begin
sComboBox10.Items.Add(datamodule2.ADOQuery_date_semaine.Fields[1].Value);
datamodule2.ADOQuery_date_semaine.next;
end;
end
else
sComboBox10.Items.Add('Aucun Produit ajouté');


jour :=copy(datetostr(date),1,2);
mois:=copy(datetostr(date),4,2);
annee:=copy(datetostr(date),7,4);
d:=EncodeDate(StrToInt(annee),StrToInt(mois),1);
DataModule2.ADOQuery_date_semaine.Close;
DataModule2.ADOQuery_date_semaine.open;

DataModule2.ADOQuery_date_semaine.Parameters.ParamByName('date1').Value:=d;
d:=EncodeDate(StrToInt(annee),StrToInt(mois)+1,1);
DataModule2.ADOQuery_date_semaine.Parameters.ParamByName('date2').Value:=d;

if (datamodule2.ADOQuery_date_semaine.RecordCount>=1) then
begin
while not (datamodule2.ADOQuery_date_semaine.Eof) do
begin
sComboBox11.Items.Add(datamodule2.ADOQuery_date_semaine.Fields[1].Value);
datamodule2.ADOQuery_date_semaine.next;
end;
end
else
sComboBox11.Items.Add('Aucun Produit ajouté');


DataModule2.ADOQuery_qte_critique.Close;
DataModule2.ADOQuery_qte_critique.open;


if (datamodule2.ADOQuery_qte_critique.RecordCount>=1) then
begin
while not (datamodule2.ADOQuery_qte_critique.Eof) do
begin
sComboBox12.Items.Add(datamodule2.ADOQuery_qte_critique.Fields[1].Value+' ,'+'qte :'+inttostr(datamodule2.ADOQuery_qte_critique.Fields[3].Value) );
datamodule2.ADOQuery_qte_critique.next;
end;
end
else
sComboBox12.Items.Add('Aucun Produit est dans un état critique ');



end;


procedure TForm1.sSpeedButton15Click(Sender: TObject);
begin
sPanel4.Visible:=false;
sGroupBox8.Visible:=false;
sPanel2.Visible:=false;
sGroupBox5.Visible:=false;

sPanel3.Visible:=false;
sGroupBox6.Visible:=false;
sGroupBox10.Visible:=false;
sGroupBox11.Visible:=false;

end;

procedure TForm1.F2Click(Sender: TObject);
var somme,qte: integer;
    jour,mois,annee,jour_nom:string;
    d : TDateTime;
begin
somme:=0;
sPanel4.Visible:=false;
sGroupBox8.Visible:=false;
sPanel2.Visible:=false;
sGroupBox5.Visible:=false;

sPanel3.Visible:=false;
sGroupBox6.Visible:=false;
sGroupBox10.Visible:=false;
sGroupBox11.Visible:=true;

DataModule2.ADOTable3.Close;
DataModule2.ADOTable3.Open;
DataModule2.ADOTable3.last;

sLabel27.Caption :=inttostr(DataModule2.ADOTable3.Fields[1].Value)+' DA';

DataModule2.ADOQuery_rech_vente.close;
DataModule2.ADOQuery_rech_vente.Parameters.ParamByName('date').Value:=date;

DataModule2.ADoquery_rech_vente.Open;

if (datamodule2.ADOQuery_rech_vente.RecordCount>=1) then
begin
while not (datamodule2.ADoquery_rech_vente.Eof) do
begin
somme:=somme+strtoint(DataModule2.ADOQuery_rech_vente.Fields[4].value);
qte:=qte+strtoint(DataModule2.ADOQuery_rech_vente.Fields[2].value);
DataModule2.ADOQuery_rech_vente.next;
end;
slabel29.caption:=inttostr(somme)+' DA';
sLabel31.Caption:=inttostr(qte)+' pièce(s)';
end
else
begin
slabel29.caption:='0 DA';
slabel31.caption:='0 pièce(s)';
end;


somme:=0;
jour:=copy(datetostr(date),1,2);
mois:=copy(datetostr(date),4,2);
annee:=copy(datetostr(date),7,4);
d:=EncodeDate(StrToInt(annee),StrToInt(mois),StrToInt(jour));

jour_nom :=longdaynames[dayofweek(d)];

DataModule2.ADOQuery_vente_date.close;

if (jour_nom='dimanche') then
begin
datamodule2.ADOQuery_vente_date.Parameters.ParamByName('date1').Value:=date;


datamodule2.ADOQuery_vente_date.Parameters.ParamByName('date2').Value:=date;
end;
if (jour_nom='lundi') then
begin
datamodule2.ADOQuery_vente_date.Parameters.ParamByName('date1').Value:=Yesterday;

datamodule2.ADOQuery_vente_date.Parameters.ParamByName('date2').Value:=date;
end;

if (jour_nom='mardi') then
begin

d:=EncodeDate(StrToInt(annee),StrToInt(mois),StrToInt(jour)-2);

datamodule2.ADOQuery_vente_date.Parameters.ParamByName('date1').Value:=d;

datamodule2.ADOQuery_vente_date.Parameters.ParamByName('date2').Value:=date;
end;

if (jour_nom='mercredi') then
begin
d:=EncodeDate(StrToInt(annee),StrToInt(mois),StrToInt(jour)-3);
datamodule2.ADOQuery_vente_date.Parameters.ParamByName('date1').Value:=d;

datamodule2.ADOQuery_vente_date.Parameters.ParamByName('date2').Value:=date;
end;
if (jour_nom='jeudi') then
begin
d:=EncodeDate(StrToInt(annee),StrToInt(mois),StrToInt(jour)-4);
datamodule2.ADOQuery_vente_date.Parameters.ParamByName('date1').Value:=d;

datamodule2.ADOQuery_vente_date.Parameters.ParamByName('date2').Value:=date;
end;
if (jour_nom='vendredi') then
begin
d:=EncodeDate(StrToInt(annee),StrToInt(mois),StrToInt(jour)-5);
datamodule2.ADOQuery_vente_date.Parameters.ParamByName('date1').Value:=d;

datamodule2.ADOQuery_vente_date.Parameters.ParamByName('date2').Value:=date;
end;
if (jour_nom='samedi') then
begin
d:=EncodeDate(StrToInt(annee),StrToInt(mois),StrToInt(jour)-6);
datamodule2.ADOQuery_vente_date.Parameters.ParamByName('date1').Value:=d;

datamodule2.ADOQuery_date_semaine.Parameters.ParamByName('date2').Value:=date;
end;

DataModule2.ADOQuery_vente_date.open;
//showmessage(inttostr(datamodule2.ADOQuery_vente_date.RecordCount));

if (datamodule2.ADOQuery_vente_date.RecordCount>=1) then
begin
while not (datamodule2.ADOQuery_vente_date.Eof) do
begin
//sComboBox10.Items.Add(datamodule2.ADOQuery_date_semaine.Fields[1].Value);
somme:=somme+strtoint(DataModule2.ADOQuery_vente_date.Fields[4].value);
//showmessage(inttostr(somme));
datamodule2.ADOQuery_vente_date.next;
end;
slabel33.Caption:=inttostr(somme)+' DA';
end
else
slabel33.Caption:=' 0 DA';


jour :=copy(datetostr(date),1,2);
mois:=copy(datetostr(date),4,2);
annee:=copy(datetostr(date),7,4);
d:=EncodeDate(StrToInt(annee),StrToInt(mois),1);
DataModule2.ADOQuery_vente_date.Close;
somme:=0;
DataModule2.ADOQuery_vente_date.Parameters.ParamByName('date1').Value:=d;
d:=EncodeDate(StrToInt(annee),StrToInt(mois)+1,1);
DataModule2.ADOQuery_vente_date.Parameters.ParamByName('date2').Value:=d;

DataModule2.ADOQuery_vente_date.open;
//showmessage('le nombre pour mois'+inttostr(datamodule2.ADOQuery_vente_date.RecordCount));
if (datamodule2.ADOQuery_vente_date.RecordCount>=1) then
begin
while not (datamodule2.ADOQuery_vente_date.Eof) do
begin

somme:=somme+strtoint(DataModule2.ADOQuery_vente_date.Fields[4].value);
datamodule2.ADOQuery_vente_date.next;
end;
slabel35.Caption:=inttostr(somme)+' DA';

end
else
slabel35.Caption:=' 0 DA';

end;

procedure TForm1.sSpeedButton6Click(Sender: TObject);
begin
if MessageDlg('Etes-vous sûr de vouloir quitter lapplication ?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
begin
close();
end;

end;

function TForm1.EnLettres(N:integer):string;
const
  Unite: Array[1..16] of string=('un','deux','trois','quatre','cinq','six',
                                'sept','huit','neuf','dix','onze','douze',
                                'treize','quatorze','quinze','seize');
  Dizaine: Array[2..8] of string=('vingt','trente','quarante','cinquante',
                                 'soixante','','quatre-vingt');
  Coefs:Array[0..3] of string=('cent','mille','million','milliard');
var
  Temp: string;
  C,D,U: Byte;
  Coef: Byte;
  I: Word;
  Neg: boolean;
begin
  if N = 0 then
  begin
    Result := ' Zéro';
    Exit;
  end;
  Result := '';
  Neg := N <0;
  if Neg then N := -N;
  Coef := 0;
  Repeat
    U := N mod 10; N := N div 10; {Récupère unité et dizaine}
    D := N mod 10; N := N div 10; {Récupère dizaine}
    if D in [1,7,9] then
    begin
      Dec(D);
      Inc(U, 10);
    end;
    Temp := '';
    if D > 1 then
    begin
      Temp := ' ' + Dizaine[D];
      if (D < 8) and ((U = 1) or (U = 11)) then
        Temp := Temp + ' et';
    end;
    if U > 16 then
    begin
      Temp := Temp + ' ' + Unite[10];
      Dec(U,10);
    end;
    if U > 0 then Temp := Temp + ' ' + Unite[U];
    if (Result = '') and (D = 8) and (U = 0) then Result := 's';
    Result := Temp + Result;
    C := N mod 10; N := N div 10; {Récupère centaine}
    if C > 0 then
    begin
      Temp := '';
      if C > 1 then Temp := ' ' + Unite[C] + Temp;
      Temp := Temp + ' ' + Coefs[0];
      if (Result = '') and (C > 1) then Result := 's';
      Result := Temp + Result;
    end;
    if N > 0 then
    begin
      Inc(Coef);
      I := N mod 1000;
      if (I > 1) and (Coef > 1) then Result := 's' + Result;
      if I > 0 then Result := ' ' + Coefs[Coef] + Result;
      if (I= 1) and (Coef = 1) then Dec(N);
    end;
  until N = 0;
  if Neg then Result := 'Moins' + Result
  else
  Result[2] := UpCase (Result[2]);
end;



procedure TForm1.A4Click(Sender: TObject);
begin
with form3 do
showmodal;
end;

procedure TForm1.sComboBox13Click(Sender: TObject);
var i,j:integer;
begin

DataModule2.ADOTable1.Close;
datamodule2.DataSource1.DataSet:=datamodule2.ADOTable1;

DataModule2.ADOTable1.Open;
i:=datamodule2.ADOTable1.RecordCount;
for j := 1 to i do
begin
with DataModule2.ADOTable1 do
begin

if (sComboBox13.Items[scombobox13.itemIndex]=datamodule2.ADOTable1.FieldByName('nom_produit').AsString) then
begin
if (strtoint(sDecimalSpinEdit11.Text)>strtoint(datamodule2.ADOTable1.Fields[3].AsString)) then
begin
ShowMessage('Erreur !! La quantité voulue dépasse la quantité stockée, Le max est :'+datamodule2.ADOTable1.Fields[3].AsString);
sDecimalSpinEdit11.Text:=datamodule2.ADOTable1.Fields[3].AsString;

end;
sCurrencyEdit11.Text:=IntToStr(strtoint(datamodule2.ADOTable1.Fields[2].AsString)*strtoint(sDecimalSpinEdit11.Text))

end;
p_u8 :=strtoint(datamodule2.ADOTable1.Fields[2].AsString);

end;
datamodule2.ADOTable1.Next;
end;
end;

procedure TForm1.sComboBox14Click(Sender: TObject);
var i,j:integer;
begin

DataModule2.ADOTable1.Close;
datamodule2.DataSource1.DataSet:=datamodule2.ADOTable1;

DataModule2.ADOTable1.Open;
i:=datamodule2.ADOTable1.RecordCount;
for j := 1 to i do
begin
with DataModule2.ADOTable1 do
begin

if (sComboBox14.Items[scombobox14.itemIndex]=datamodule2.ADOTable1.FieldByName('nom_produit').AsString) then
begin
if (strtoint(sDecimalSpinEdit12.Text)>strtoint(datamodule2.ADOTable1.Fields[3].AsString)) then
begin
ShowMessage('Erreur !! La quantité voulue dépasse la quantité stockée, Le max est :'+datamodule2.ADOTable1.Fields[3].AsString);
sDecimalSpinEdit12.Text:=datamodule2.ADOTable1.Fields[3].AsString;

end;
sCurrencyEdit12.Text:=IntToStr(strtoint(datamodule2.ADOTable1.Fields[2].AsString)*strtoint(sDecimalSpinEdit12.Text))

end;
p_u9 :=strtoint(datamodule2.ADOTable1.Fields[2].AsString);

end;
datamodule2.ADOTable1.Next;
end;
end;

procedure TForm1.sComboBox15Click(Sender: TObject);
var i,j:integer;
begin

DataModule2.ADOTable1.Close;
datamodule2.DataSource1.DataSet:=datamodule2.ADOTable1;

DataModule2.ADOTable1.Open;
i:=datamodule2.ADOTable1.RecordCount;
for j := 1 to i do
begin
with DataModule2.ADOTable1 do
begin

if (sComboBox15.Items[scombobox15.itemIndex]=datamodule2.ADOTable1.FieldByName('nom_produit').AsString) then
begin
if (strtoint(sDecimalSpinEdit13.Text)>strtoint(datamodule2.ADOTable1.Fields[3].AsString)) then
begin
ShowMessage('Erreur !! La quantité voulue dépasse la quantité stockée, Le max est :'+datamodule2.ADOTable1.Fields[3].AsString);
sDecimalSpinEdit13.Text:=datamodule2.ADOTable1.Fields[3].AsString;

end;
sCurrencyEdit13.Text:=IntToStr(strtoint(datamodule2.ADOTable1.Fields[2].AsString)*strtoint(sDecimalSpinEdit13.Text))

end;
p_u10 :=strtoint(datamodule2.ADOTable1.Fields[2].AsString);

end;
datamodule2.ADOTable1.Next;
end;
end;

procedure TForm1.sComboBox16Click(Sender: TObject);
var i,j:integer;
begin

DataModule2.ADOTable1.Close;
datamodule2.DataSource1.DataSet:=datamodule2.ADOTable1;

DataModule2.ADOTable1.Open;
i:=datamodule2.ADOTable1.RecordCount;
for j := 1 to i do
begin
with DataModule2.ADOTable1 do
begin

if (sComboBox16.Items[scombobox16.itemIndex]=datamodule2.ADOTable1.FieldByName('nom_produit').AsString) then
begin
if (strtoint(sDecimalSpinEdit14.Text)>strtoint(datamodule2.ADOTable1.Fields[3].AsString)) then
begin
ShowMessage('Erreur !! La quantité voulue dépasse la quantité stockée, Le max est :'+datamodule2.ADOTable1.Fields[3].AsString);
sDecimalSpinEdit14.Text:=datamodule2.ADOTable1.Fields[3].AsString;

end;
sCurrencyEdit14.Text:=IntToStr(strtoint(datamodule2.ADOTable1.Fields[2].AsString)*strtoint(sDecimalSpinEdit14.Text))

end;
p_u11 :=strtoint(datamodule2.ADOTable1.Fields[2].AsString);

end;
datamodule2.ADOTable1.Next;
end;
end;

procedure TForm1.sComboBox17Click(Sender: TObject);
var i,j:integer;
begin

DataModule2.ADOTable1.Close;
datamodule2.DataSource1.DataSet:=datamodule2.ADOTable1;

DataModule2.ADOTable1.Open;
i:=datamodule2.ADOTable1.RecordCount;
for j := 1 to i do
begin
with DataModule2.ADOTable1 do
begin

if (sComboBox17.Items[scombobox17.itemIndex]=datamodule2.ADOTable1.FieldByName('nom_produit').AsString) then
begin
if (strtoint(sDecimalSpinEdit15.Text)>strtoint(datamodule2.ADOTable1.Fields[3].AsString)) then
begin
ShowMessage('Erreur !! La quantité voulue dépasse la quantité stockée, Le max est :'+datamodule2.ADOTable1.Fields[3].AsString);
sDecimalSpinEdit15.Text:=datamodule2.ADOTable1.Fields[3].AsString;

end;
sCurrencyEdit15.Text:=IntToStr(strtoint(datamodule2.ADOTable1.Fields[2].AsString)*strtoint(sDecimalSpinEdit15.Text))

end;
p_u12 :=strtoint(datamodule2.ADOTable1.Fields[2].AsString);

end;
datamodule2.ADOTable1.Next;
end;
end;

procedure TForm1.sComboBox18Click(Sender: TObject);
var i,j:integer;
begin

DataModule2.ADOTable1.Close;
datamodule2.DataSource1.DataSet:=datamodule2.ADOTable1;

DataModule2.ADOTable1.Open;
i:=datamodule2.ADOTable1.RecordCount;
for j := 1 to i do
begin
with DataModule2.ADOTable1 do
begin

if (sComboBox18.Items[scombobox18.itemIndex]=datamodule2.ADOTable1.FieldByName('nom_produit').AsString) then
begin
if (strtoint(sDecimalSpinEdit16.Text)>strtoint(datamodule2.ADOTable1.Fields[3].AsString)) then
begin
ShowMessage('Erreur !! La quantité voulue dépasse la quantité stockée, Le max est :'+datamodule2.ADOTable1.Fields[3].AsString);
sDecimalSpinEdit16.Text:=datamodule2.ADOTable1.Fields[3].AsString;

end;
sCurrencyEdit16.Text:=IntToStr(strtoint(datamodule2.ADOTable1.Fields[2].AsString)*strtoint(sDecimalSpinEdit16.Text))

end;
p_u13 :=strtoint(datamodule2.ADOTable1.Fields[2].AsString);

end;
datamodule2.ADOTable1.Next;
end;
end;

procedure TForm1.sComboBox19Click(Sender: TObject);
var i,j:integer;
begin

DataModule2.ADOTable1.Close;
datamodule2.DataSource1.DataSet:=datamodule2.ADOTable1;

DataModule2.ADOTable1.Open;
i:=datamodule2.ADOTable1.RecordCount;
for j := 1 to i do
begin
with DataModule2.ADOTable1 do
begin

if (sComboBox19.Items[scombobox19.itemIndex]=datamodule2.ADOTable1.FieldByName('nom_produit').AsString) then
begin
if (strtoint(sDecimalSpinEdit17.Text)>strtoint(datamodule2.ADOTable1.Fields[3].AsString)) then
begin
ShowMessage('Erreur !! La quantité voulue dépasse la quantité stockée, Le max est :'+datamodule2.ADOTable1.Fields[3].AsString);
sDecimalSpinEdit17.Text:=datamodule2.ADOTable1.Fields[3].AsString;

end;
sCurrencyEdit17.Text:=IntToStr(strtoint(datamodule2.ADOTable1.Fields[2].AsString)*strtoint(sDecimalSpinEdit17.Text))

end;
p_u14 :=strtoint(datamodule2.ADOTable1.Fields[2].AsString);

end;
datamodule2.ADOTable1.Next;
end;
end;

procedure TForm1.sComboBox20Click(Sender: TObject);
var i,j:integer;
begin

DataModule2.ADOTable1.Close;
datamodule2.DataSource1.DataSet:=datamodule2.ADOTable1;

DataModule2.ADOTable1.Open;
i:=datamodule2.ADOTable1.RecordCount;
for j := 1 to i do
begin
with DataModule2.ADOTable1 do
begin

if (sComboBox17.Items[scombobox20.itemIndex]=datamodule2.ADOTable1.FieldByName('nom_produit').AsString) then
begin
if (strtoint(sDecimalSpinEdit18.Text)>strtoint(datamodule2.ADOTable1.Fields[3].AsString)) then
begin
ShowMessage('Erreur !! La quantité voulue dépasse la quantité stockée, Le max est :'+datamodule2.ADOTable1.Fields[3].AsString);
sDecimalSpinEdit18.Text:=datamodule2.ADOTable1.Fields[3].AsString;

end;
sCurrencyEdit18.Text:=IntToStr(strtoint(datamodule2.ADOTable1.Fields[2].AsString)*strtoint(sDecimalSpinEdit18.Text))

end;
p_u15 :=strtoint(datamodule2.ADOTable1.Fields[2].AsString);

end;
datamodule2.ADOTable1.Next;
end;
end;

procedure TForm1.sComboBox21Click(Sender: TObject);
var i,j:integer;
begin

DataModule2.ADOTable1.Close;
datamodule2.DataSource1.DataSet:=datamodule2.ADOTable1;

DataModule2.ADOTable1.Open;
i:=datamodule2.ADOTable1.RecordCount;
for j := 1 to i do
begin
with DataModule2.ADOTable1 do
begin

if (sComboBox21.Items[scombobox21.itemIndex]=datamodule2.ADOTable1.FieldByName('nom_produit').AsString) then
begin
if (strtoint(sDecimalSpinEdit19.Text)>strtoint(datamodule2.ADOTable1.Fields[3].AsString)) then
begin
ShowMessage('Erreur !! La quantité voulue dépasse la quantité stockée, Le max est :'+datamodule2.ADOTable1.Fields[3].AsString);
sDecimalSpinEdit19.Text:=datamodule2.ADOTable1.Fields[3].AsString;

end;
sCurrencyEdit19.Text:=IntToStr(strtoint(datamodule2.ADOTable1.Fields[2].AsString)*strtoint(sDecimalSpinEdit19.Text))

end;
p_u16 :=strtoint(datamodule2.ADOTable1.Fields[2].AsString);

end;
datamodule2.ADOTable1.Next;
end;
end;

procedure TForm1.sComboBox22Click(Sender: TObject);
var i,j:integer;
begin

DataModule2.ADOTable1.Close;
datamodule2.DataSource1.DataSet:=datamodule2.ADOTable1;

DataModule2.ADOTable1.Open;
i:=datamodule2.ADOTable1.RecordCount;
for j := 1 to i do
begin
with DataModule2.ADOTable1 do
begin

if (sComboBox12.Items[scombobox22.itemIndex]=datamodule2.ADOTable1.FieldByName('nom_produit').AsString) then
begin
if (strtoint(sDecimalSpinEdit20.Text)>strtoint(datamodule2.ADOTable1.Fields[3].AsString)) then
begin
ShowMessage('Erreur !! La quantité voulue dépasse la quantité stockée, Le max est :'+datamodule2.ADOTable1.Fields[3].AsString);
sDecimalSpinEdit20.Text:=datamodule2.ADOTable1.Fields[3].AsString;

end;
sCurrencyEdit20.Text:=IntToStr(strtoint(datamodule2.ADOTable1.Fields[2].AsString)*strtoint(sDecimalSpinEdit20.Text))

end;
p_u17 :=strtoint(datamodule2.ADOTable1.Fields[2].AsString);

end;
datamodule2.ADOTable1.Next;
end;
end;

procedure TForm1.sComboBox23Click(Sender: TObject);
var i,j:integer;
begin

DataModule2.ADOTable1.Close;
datamodule2.DataSource1.DataSet:=datamodule2.ADOTable1;

DataModule2.ADOTable1.Open;
i:=datamodule2.ADOTable1.RecordCount;
for j := 1 to i do
begin
with DataModule2.ADOTable1 do
begin

if (sComboBox23.Items[scombobox23.itemIndex]=datamodule2.ADOTable1.FieldByName('nom_produit').AsString) then
begin
if (strtoint(sDecimalSpinEdit21.Text)>strtoint(datamodule2.ADOTable1.Fields[3].AsString)) then
begin
ShowMessage('Erreur !! La quantité voulue dépasse la quantité stockée, Le max est :'+datamodule2.ADOTable1.Fields[3].AsString);
sDecimalSpinEdit21.Text:=datamodule2.ADOTable1.Fields[3].AsString;

end;
sCurrencyEdit21.Text:=IntToStr(strtoint(datamodule2.ADOTable1.Fields[2].AsString)*strtoint(sDecimalSpinEdit21.Text))

end;
p_u18 :=strtoint(datamodule2.ADOTable1.Fields[2].AsString);

end;
datamodule2.ADOTable1.Next;
end;
end;

procedure TForm1.sComboBox24Click(Sender: TObject);
var i,j:integer;
begin

DataModule2.ADOTable1.Close;
datamodule2.DataSource1.DataSet:=datamodule2.ADOTable1;

DataModule2.ADOTable1.Open;
i:=datamodule2.ADOTable1.RecordCount;
for j := 1 to i do
begin
with DataModule2.ADOTable1 do
begin

if (sComboBox24.Items[scombobox24.itemIndex]=datamodule2.ADOTable1.FieldByName('nom_produit').AsString) then
begin
if (strtoint(sDecimalSpinEdit22.Text)>strtoint(datamodule2.ADOTable1.Fields[3].AsString)) then
begin
ShowMessage('Erreur !! La quantité voulue dépasse la quantité stockée, Le max est :'+datamodule2.ADOTable1.Fields[3].AsString);
sDecimalSpinEdit22.Text:=datamodule2.ADOTable1.Fields[3].AsString;

end;
sCurrencyEdit22.Text:=IntToStr(strtoint(datamodule2.ADOTable1.Fields[2].AsString)*strtoint(sDecimalSpinEdit22.Text))

end;
p_u19 :=strtoint(datamodule2.ADOTable1.Fields[2].AsString);

end;
datamodule2.ADOTable1.Next;
end;
end;

procedure TForm1.sComboBox25Click(Sender: TObject);
var i,j:integer;
begin

DataModule2.ADOTable1.Close;
datamodule2.DataSource1.DataSet:=datamodule2.ADOTable1;

DataModule2.ADOTable1.Open;
i:=datamodule2.ADOTable1.RecordCount;
for j := 1 to i do
begin
with DataModule2.ADOTable1 do
begin

if (sComboBox25.Items[scombobox25.itemIndex]=datamodule2.ADOTable1.FieldByName('nom_produit').AsString) then
begin
if (strtoint(sDecimalSpinEdit23.Text)>strtoint(datamodule2.ADOTable1.Fields[3].AsString)) then
begin
ShowMessage('Erreur !! La quantité voulue dépasse la quantité stockée, Le max est :'+datamodule2.ADOTable1.Fields[3].AsString);
sDecimalSpinEdit23.Text:=datamodule2.ADOTable1.Fields[3].AsString;

end;
sCurrencyEdit23.Text:=IntToStr(strtoint(datamodule2.ADOTable1.Fields[2].AsString)*strtoint(sDecimalSpinEdit23.Text))

end;
p_u20 :=strtoint(datamodule2.ADOTable1.Fields[2].AsString);

end;
datamodule2.ADOTable1.Next;
end;
end;

procedure TForm1.sComboBox26Click(Sender: TObject);
var i,j:integer;
begin

DataModule2.ADOTable1.Close;
datamodule2.DataSource1.DataSet:=datamodule2.ADOTable1;

DataModule2.ADOTable1.Open;
i:=datamodule2.ADOTable1.RecordCount;
for j := 1 to i do
begin
with DataModule2.ADOTable1 do
begin

if (sComboBox26.Items[scombobox26.itemIndex]=datamodule2.ADOTable1.FieldByName('nom_produit').AsString) then
begin
if (strtoint(sDecimalSpinEdit24.Text)>strtoint(datamodule2.ADOTable1.Fields[3].AsString)) then
begin
ShowMessage('Erreur !! La quantité voulue dépasse la quantité stockée, Le max est :'+datamodule2.ADOTable1.Fields[3].AsString);
sDecimalSpinEdit24.Text:=datamodule2.ADOTable1.Fields[3].AsString;

end;
sCurrencyEdit24.Text:=IntToStr(strtoint(datamodule2.ADOTable1.Fields[2].AsString)*strtoint(sDecimalSpinEdit24.Text))

end;
p_u21 :=strtoint(datamodule2.ADOTable1.Fields[2].AsString);

end;
datamodule2.ADOTable1.Next;
end;
end;

procedure TForm1.sComboBox27Click(Sender: TObject);
var i,j:integer;
begin

DataModule2.ADOTable1.Close;
datamodule2.DataSource1.DataSet:=datamodule2.ADOTable1;

DataModule2.ADOTable1.Open;
i:=datamodule2.ADOTable1.RecordCount;
for j := 1 to i do
begin
with DataModule2.ADOTable1 do
begin

if (sComboBox27.Items[scombobox27.itemIndex]=datamodule2.ADOTable1.FieldByName('nom_produit').AsString) then
begin
if (strtoint(sDecimalSpinEdit25.Text)>strtoint(datamodule2.ADOTable1.Fields[3].AsString)) then
begin
ShowMessage('Erreur !! La quantité voulue dépasse la quantité stockée, Le max est :'+datamodule2.ADOTable1.Fields[3].AsString);
sDecimalSpinEdit25.Text:=datamodule2.ADOTable1.Fields[3].AsString;

end;
sCurrencyEdit25.Text:=IntToStr(strtoint(datamodule2.ADOTable1.Fields[2].AsString)*strtoint(sDecimalSpinEdit25.Text))

end;
p_u22 :=strtoint(datamodule2.ADOTable1.Fields[2].AsString);

end;
datamodule2.ADOTable1.Next;
end;
end;

procedure TForm1.sComboBox28Click(Sender: TObject);
var i,j:integer;
begin

DataModule2.ADOTable1.Close;
datamodule2.DataSource1.DataSet:=datamodule2.ADOTable1;

DataModule2.ADOTable1.Open;
i:=datamodule2.ADOTable1.RecordCount;
for j := 1 to i do
begin
with DataModule2.ADOTable1 do
begin

if (sComboBox28.Items[scombobox28.itemIndex]=datamodule2.ADOTable1.FieldByName('nom_produit').AsString) then
begin
if (strtoint(sDecimalSpinEdit26.Text)>strtoint(datamodule2.ADOTable1.Fields[3].AsString)) then
begin
ShowMessage('Erreur !! La quantité voulue dépasse la quantité stockée, Le max est :'+datamodule2.ADOTable1.Fields[3].AsString);
sDecimalSpinEdit26.Text:=datamodule2.ADOTable1.Fields[3].AsString;

end;
sCurrencyEdit26.Text:=IntToStr(strtoint(datamodule2.ADOTable1.Fields[2].AsString)*strtoint(sDecimalSpinEdit26.Text))

end;
p_u23 :=strtoint(datamodule2.ADOTable1.Fields[2].AsString);

end;
datamodule2.ADOTable1.Next;
end;
end;

procedure TForm1.sComboBox29Click(Sender: TObject);
var i,j:integer;
begin

DataModule2.ADOTable1.Close;
datamodule2.DataSource1.DataSet:=datamodule2.ADOTable1;

DataModule2.ADOTable1.Open;
i:=datamodule2.ADOTable1.RecordCount;
for j := 1 to i do
begin
with DataModule2.ADOTable1 do
begin

if (sComboBox29.Items[scombobox26.itemIndex]=datamodule2.ADOTable1.FieldByName('nom_produit').AsString) then
begin
if (strtoint(sDecimalSpinEdit27.Text)>strtoint(datamodule2.ADOTable1.Fields[3].AsString)) then
begin
ShowMessage('Erreur !! La quantité voulue dépasse la quantité stockée, Le max est :'+datamodule2.ADOTable1.Fields[3].AsString);
sDecimalSpinEdit27.Text:=datamodule2.ADOTable1.Fields[3].AsString;

end;
sCurrencyEdit27.Text:=IntToStr(strtoint(datamodule2.ADOTable1.Fields[2].AsString)*strtoint(sDecimalSpinEdit27.Text))

end;
p_u24 :=strtoint(datamodule2.ADOTable1.Fields[2].AsString);

end;
datamodule2.ADOTable1.Next;
end;
end;

procedure TForm1.sComboBox30Click(Sender: TObject);
var i,j:integer;
begin

DataModule2.ADOTable1.Close;
datamodule2.DataSource1.DataSet:=datamodule2.ADOTable1;

DataModule2.ADOTable1.Open;
i:=datamodule2.ADOTable1.RecordCount;
for j := 1 to i do
begin
with DataModule2.ADOTable1 do
begin

if (sComboBox30.Items[scombobox30.itemIndex]=datamodule2.ADOTable1.FieldByName('nom_produit').AsString) then
begin
if (strtoint(sDecimalSpinEdit28.Text)>strtoint(datamodule2.ADOTable1.Fields[3].AsString)) then
begin
ShowMessage('Erreur !! La quantité voulue dépasse la quantité stockée, Le max est :'+datamodule2.ADOTable1.Fields[3].AsString);
sDecimalSpinEdit28.Text:=datamodule2.ADOTable1.Fields[3].AsString;

end;
sCurrencyEdit28.Text:=IntToStr(strtoint(datamodule2.ADOTable1.Fields[2].AsString)*strtoint(sDecimalSpinEdit28.Text))

end;
p_u25 :=strtoint(datamodule2.ADOTable1.Fields[2].AsString);

end;
datamodule2.ADOTable1.Next;
end;
end;

procedure TForm1.sComboBox31Click(Sender: TObject);
var i,j:integer;
begin

DataModule2.ADOTable1.Close;
datamodule2.DataSource1.DataSet:=datamodule2.ADOTable1;

DataModule2.ADOTable1.Open;
i:=datamodule2.ADOTable1.RecordCount;
for j := 1 to i do
begin
with DataModule2.ADOTable1 do
begin

if (sComboBox31.Items[scombobox31.itemIndex]=datamodule2.ADOTable1.FieldByName('nom_produit').AsString) then
begin
if (strtoint(sDecimalSpinEdit29.Text)>strtoint(datamodule2.ADOTable1.Fields[3].AsString)) then
begin
ShowMessage('Erreur !! La quantité voulue dépasse la quantité stockée, Le max est :'+datamodule2.ADOTable1.Fields[3].AsString);
sDecimalSpinEdit29.Text:=datamodule2.ADOTable1.Fields[3].AsString;

end;
sCurrencyEdit29.Text:=IntToStr(strtoint(datamodule2.ADOTable1.Fields[2].AsString)*strtoint(sDecimalSpinEdit29.Text))

end;
p_u26 :=strtoint(datamodule2.ADOTable1.Fields[2].AsString);

end;
datamodule2.ADOTable1.Next;
end;
end;

procedure TForm1.sComboBox32Click(Sender: TObject);
var i,j:integer;
begin

DataModule2.ADOTable1.Close;
datamodule2.DataSource1.DataSet:=datamodule2.ADOTable1;

DataModule2.ADOTable1.Open;
i:=datamodule2.ADOTable1.RecordCount;
for j := 1 to i do
begin
with DataModule2.ADOTable1 do
begin

if (sComboBox32.Items[scombobox32.itemIndex]=datamodule2.ADOTable1.FieldByName('nom_produit').AsString) then
begin
if (strtoint(sDecimalSpinEdit30.Text)>strtoint(datamodule2.ADOTable1.Fields[3].AsString)) then
begin
ShowMessage('Erreur !! La quantité voulue dépasse la quantité stockée, Le max est :'+datamodule2.ADOTable1.Fields[3].AsString);
sDecimalSpinEdit30.Text:=datamodule2.ADOTable1.Fields[3].AsString;

end;
sCurrencyEdit30.Text:=IntToStr(strtoint(datamodule2.ADOTable1.Fields[2].AsString)*strtoint(sDecimalSpinEdit30.Text))

end;
p_u27 :=strtoint(datamodule2.ADOTable1.Fields[2].AsString);

end;
datamodule2.ADOTable1.Next;
end;
end;

procedure TForm1.sComboBox33Click(Sender: TObject);
var i,j:integer;
begin

DataModule2.ADOTable1.Close;
datamodule2.DataSource1.DataSet:=datamodule2.ADOTable1;

DataModule2.ADOTable1.Open;
i:=datamodule2.ADOTable1.RecordCount;
for j := 1 to i do
begin
with DataModule2.ADOTable1 do
begin

if (sComboBox33.Items[scombobox33.itemIndex]=datamodule2.ADOTable1.FieldByName('nom_produit').AsString) then
begin
if (strtoint(sDecimalSpinEdit31.Text)>strtoint(datamodule2.ADOTable1.Fields[3].AsString)) then
begin
ShowMessage('Erreur !! La quantité voulue dépasse la quantité stockée, Le max est :'+datamodule2.ADOTable1.Fields[3].AsString);
sDecimalSpinEdit31.Text:=datamodule2.ADOTable1.Fields[3].AsString;

end;
sCurrencyEdit31.Text:=IntToStr(strtoint(datamodule2.ADOTable1.Fields[2].AsString)*strtoint(sDecimalSpinEdit31.Text))

end;
p_u28 :=strtoint(datamodule2.ADOTable1.Fields[2].AsString);

end;
datamodule2.ADOTable1.Next;
end;
end;

procedure TForm1.sComboBox34Click(Sender: TObject);
var i,j:integer;
begin

DataModule2.ADOTable1.Close;
datamodule2.DataSource1.DataSet:=datamodule2.ADOTable1;

DataModule2.ADOTable1.Open;
i:=datamodule2.ADOTable1.RecordCount;
for j := 1 to i do
begin
with DataModule2.ADOTable1 do
begin

if (sComboBox34.Items[scombobox34.itemIndex]=datamodule2.ADOTable1.FieldByName('nom_produit').AsString) then
begin
if (strtoint(sDecimalSpinEdit32.Text)>strtoint(datamodule2.ADOTable1.Fields[3].AsString)) then
begin
ShowMessage('Erreur !! La quantité voulue dépasse la quantité stockée, Le max est :'+datamodule2.ADOTable1.Fields[3].AsString);
sDecimalSpinEdit32.Text:=datamodule2.ADOTable1.Fields[3].AsString;

end;
sCurrencyEdit32.Text:=IntToStr(strtoint(datamodule2.ADOTable1.Fields[2].AsString)*strtoint(sDecimalSpinEdit32.Text))

end;
p_u29 :=strtoint(datamodule2.ADOTable1.Fields[2].AsString);

end;
datamodule2.ADOTable1.Next;
end;
end;

procedure TForm1.sComboBox35Click(Sender: TObject);
var i,j:integer;
begin

DataModule2.ADOTable1.Close;
datamodule2.DataSource1.DataSet:=datamodule2.ADOTable1;

DataModule2.ADOTable1.Open;
i:=datamodule2.ADOTable1.RecordCount;
for j := 1 to i do
begin
with DataModule2.ADOTable1 do
begin

if (sComboBox35.Items[scombobox35.itemIndex]=datamodule2.ADOTable1.FieldByName('nom_produit').AsString) then
begin
if (strtoint(sDecimalSpinEdit33.Text)>strtoint(datamodule2.ADOTable1.Fields[3].AsString)) then
begin
ShowMessage('Erreur !! La quantité voulue dépasse la quantité stockée, Le max est :'+datamodule2.ADOTable1.Fields[3].AsString);
sDecimalSpinEdit33.Text:=datamodule2.ADOTable1.Fields[3].AsString;

end;
sCurrencyEdit33.Text:=IntToStr(strtoint(datamodule2.ADOTable1.Fields[2].AsString)*strtoint(sDecimalSpinEdit33.Text))

end;
p_u30 :=strtoint(datamodule2.ADOTable1.Fields[2].AsString);

end;
datamodule2.ADOTable1.Next;
end;
end;

procedure TForm1.sComboBox36Click(Sender: TObject);
var i,j:integer;
begin

DataModule2.ADOTable1.Close;
datamodule2.DataSource1.DataSet:=datamodule2.ADOTable1;

DataModule2.ADOTable1.Open;
i:=datamodule2.ADOTable1.RecordCount;
for j := 1 to i do
begin
with DataModule2.ADOTable1 do
begin

if (sComboBox36.Items[scombobox36.itemIndex]=datamodule2.ADOTable1.FieldByName('nom_produit').AsString) then
begin
if (strtoint(sDecimalSpinEdit34.Text)>strtoint(datamodule2.ADOTable1.Fields[3].AsString)) then
begin
ShowMessage('Erreur !! La quantité voulue dépasse la quantité stockée, Le max est :'+datamodule2.ADOTable1.Fields[3].AsString);
sDecimalSpinEdit34.Text:=datamodule2.ADOTable1.Fields[3].AsString;

end;
sCurrencyEdit34.Text:=IntToStr(strtoint(datamodule2.ADOTable1.Fields[2].AsString)*strtoint(sDecimalSpinEdit34.Text))

end;
p_u31 :=strtoint(datamodule2.ADOTable1.Fields[2].AsString);

end;
datamodule2.ADOTable1.Next;
end;
end;

procedure TForm1.sComboBox37Click(Sender: TObject);
var i,j:integer;
begin

DataModule2.ADOTable1.Close;
datamodule2.DataSource1.DataSet:=datamodule2.ADOTable1;

DataModule2.ADOTable1.Open;
i:=datamodule2.ADOTable1.RecordCount;
for j := 1 to i do
begin
with DataModule2.ADOTable1 do
begin

if (sComboBox37.Items[scombobox37.itemIndex]=datamodule2.ADOTable1.FieldByName('nom_produit').AsString) then
begin
if (strtoint(sDecimalSpinEdit35.Text)>strtoint(datamodule2.ADOTable1.Fields[3].AsString)) then
begin
ShowMessage('Erreur !! La quantité voulue dépasse la quantité stockée, Le max est :'+datamodule2.ADOTable1.Fields[3].AsString);
sDecimalSpinEdit35.Text:=datamodule2.ADOTable1.Fields[3].AsString;

end;
sCurrencyEdit35.Text:=IntToStr(strtoint(datamodule2.ADOTable1.Fields[2].AsString)*strtoint(sDecimalSpinEdit35.Text))

end;
p_u32 :=strtoint(datamodule2.ADOTable1.Fields[2].AsString);

end;
datamodule2.ADOTable1.Next;
end;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
begin
 // Sauvegarde la WndProc actuelle du DBGrid1.
  OldWindowProc := DBGrid1.WindowProc;
  // Affecte une nouvelle procédure de fenêtre.
  DBGrid1.WindowProc := DBGridNewWindowProc;
end;
end;

procedure TForm1.FormMouseWheelDown(Sender: TObject; Shift: TShiftState;
  MousePos: TPoint; var Handled: Boolean);
begin
MousePos := ScreenToClient(MousePos);
If
(MousePos.X > sScrollBox1.Left) and
(MousePos.Y > sScrollBox1.Top) and
(MousePos.X < sScrollBox1.Left + sScrollBox1.Width) and
(MousePos.Y < sScrollBox1.Top + sScrollBox1.Height)
then sScrollBox1.Perform(WM_VSCROLL,1,0);

end;

procedure TForm1.FormMouseWheelUp(Sender: TObject; Shift: TShiftState;
  MousePos: TPoint; var Handled: Boolean);
begin
MousePos := ScreenToClient(MousePos);
If
(MousePos.X > sScrollBox1.Left) and
(MousePos.Y > sScrollBox1.Top) and
(MousePos.X < sScrollBox1.Left + sScrollBox1.Width) and
(MousePos.Y < sScrollBox1.Top + sScrollBox1.Height)
then sScrollBox1.Perform(WM_VSCROLL,0,0);

end;
procedure TForm1.DBGridNewWindowProc(var Msg: TMessage);
begin
  //.Interception de l'évènement WM_MOUSEWHEEL.
  if Msg.Msg = WM_MOUSEWHEEL then
  begin
    if (DBGrid1.DataSource.DataSet.Active) then
    begin
      if SmallInt(Msg.WParamHi) < 0 then
        DBGrid1.DataSource.DataSet.Next
      else
        DBGrid1.DataSource.DataSet.Prior;
    end;
    Exit;
  end;

  //.Traitement normal des autres message.
  OldWindowProc(Msg);
end;
end.
