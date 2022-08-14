unit Cadastro;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComCtrls, DBCtrls, Mask, ImgList, ExtCtrls, ToolWin,
  Grids, IdBaseComponent, IdComponent, IdTCPConnection, IdTCPClient, IdHTTP,
  XMLDoc, XMLIntf, IdMessage, IdMessageClient, IdSMTP;

type
  TFCadasto = class(TForm)
    Lb_Cad_Cli_CgcCpf: TLabel;
    Lb_Cad_Cli_InscRg: TLabel;
    MERg: TMaskEdit;
    Lb_Cad_Cli_Razao: TLabel;
    Lb_Cad_Cli_Telefone: TLabel;
    Label20: TLabel;
    PCEndereco: TPageControl;
    TSEndereco: TTabSheet;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    DBEdit4: TDBEdit;
    Label7: TLabel;
    CBEstado: TComboBox;
    Label8: TLabel;
    SBPRodape: TStatusBar;
    LDescUf: TLabel;
    ToolBar1: TToolBar;
    TBSair: TToolButton;
    ToolButton4: TToolButton;
    TBGravar: TToolButton;
    ImageList1: TImageList;
    Label9: TLabel;
    SGCadastro: TStringGrid;
    ENome: TEdit;
    Etelefone: TEdit;
    Email: TEdit;
    ECep: TEdit;
    ELogradouro: TEdit;
    ENro: TEdit;
    EBairro: TEdit;
    EComplemento: TEdit;
    ECidade: TEdit;
    CBPais: TComboBox;
    MECpf: TMaskEdit;
    IdHTTPClep: TIdHTTP;
    MCepApp: TMemo;
    ToolButton1: TToolButton;
    TBGeraXml: TToolButton;
    Timer: TTimer;
    OpenDialog1: TOpenDialog;
    ToolButton2: TToolButton;
    TBEmail: TToolButton;
    IdSMTP1: TIdSMTP;
    IdMessage1: TIdMessage;
    procedure Timer1Timer(Sender: TObject);
    procedure TBSairClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure MECpfClick(Sender: TObject);
    procedure ECepClick(Sender: TObject);
    procedure TBGravarClick(Sender: TObject);
    procedure LimpVar(cTipo : String);
    procedure MascCpf();
    procedure Cabecario();
    Procedure Cep();
    Procedure AlimentaCepJson();
    procedure TBGeraXmlClick(Sender: TObject);
    procedure TBEmailClick(Sender: TObject);
 private
    { Private declarations }
  public
    { Public declarations }
  end;

Function VerificaCpf(Cpf : String): Boolean;
Function DescEstado(Codigo : String): string;
Function Limpa(str : string): string;
Function LimpCa(str : string): string;
function Padr(cCaracter : String; cTexto : String; nTam : Integer): String;
function replicate_chr(sBase : string; iTam : Integer): String;

var
  FCadasto: TFCadasto;

var aRXml, aRJson, aRJsonI : Array of Array of String;
Var nCel : Integer;

implementation

{$R *.dfm}

procedure TFCadasto.Timer1Timer(Sender: TObject);
begin
  Screen.Cursors[crSqlWait] := Screen.Cursors[crDefault];

  SBPRodape.Panels[0].text := ' Teste Cadastro de Cliente ';
  SBPRodape.Panels[2].text := ' Data '+ copy(DateTimeToStr(now),1,10)+'  Hora '+copy(DateTimeToStr(now),11,10);
end;

procedure TFCadasto.TBSairClick(Sender: TObject);
begin
  if (Messagedlg('Confirma saida ?',mtConfirmation,[mbYes,mbNo],0) = mrYes) then
     application.terminate;
end;

procedure TFCadasto.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  TBSairClick(Sender);
end;

procedure TFCadasto.FormCreate(Sender: TObject);
begin
  MascCpf();
  Cep();

  Cabecario();
  LimpVar('T');

  nCel := 1;
end;

procedure TFCadasto.FormShow(Sender: TObject);
begin
  MECpf.Setfocus;
//  FCadasto.MECpf.Setfocus;
end;

procedure TFCadasto.MECpfClick(Sender: TObject);
begin
  MascCpf();
end;

procedure TFCadasto.ECepClick(Sender: TObject);
begin
  Cep();
end;

procedure TFCadasto.TBGravarClick(Sender: TObject);
var oOk : Boolean;
begin
// gravar
  oOk := True;
  if (trim(MECpf.Text) = '') then
     Begin
       application.messagebox('CPF não preenchido, favor verificar !!!','Atenção',MB_OK+MB_ICONInformation);
       oOk := False;
     end;
  if (trim(MERg.Text) = '') then
     Begin
       application.messagebox('Rg não preenchido, favor verificar !!!','Atenção',MB_OK+MB_ICONInformation);
       oOk := False;
     end;
  if (trim(ENome.Text) = '') then
     Begin
       application.messagebox('Nome não preenchido, favor verificar !!!','Atenção',MB_OK+MB_ICONInformation);
       oOk := False;
     end;
  if (trim(Etelefone.Text) = '') then
     Begin
       application.messagebox('Telefone não preenchido, favor verificar !!!','Atenção',MB_OK+MB_ICONInformation);
       oOk := False;
     end;
  if (trim(Email.Text) = '') then
     Begin
       application.messagebox('E-Mail não preenchido, favor verificar !!!','Atenção',MB_OK+MB_ICONInformation);
       oOk := False;
     end;

  if (trim(Ecep.Text) = '') then
     Begin
       application.messagebox('Cep não preenchido, favor verificar !!!','Atenção',MB_OK+MB_ICONInformation);
       oOk := False;
     end;

  if (trim(ELogradouro.Text) = '') then
     Begin
       application.messagebox('Logradouro não preenchido, favor verificar !!!','Atenção',MB_OK+MB_ICONInformation);
       oOk := False;
     end;

  if (trim(ENro.Text) = '') then
     Begin
       application.messagebox('Numero não preenchido, favor verificar !!!','Atenção',MB_OK+MB_ICONInformation);
       oOk := False;
     end;

  if (trim(EBairro.Text) = '') then
     Begin
       application.messagebox('Bairro não preenchido, favor verificar !!!','Atenção',MB_OK+MB_ICONInformation);
       oOk := False;
     end;

  if (trim(ECidade.Text) = '') then
     Begin
       application.messagebox('Cidade não preenchido, favor verificar !!!','Atenção',MB_OK+MB_ICONInformation);
       oOk := False;
     end;

  if oOk then
     Begin
       SGCadastro.Cells[00,nCel] := IntToStr(nCel);  // Codigo
       SGCadastro.Cells[01,nCel] := MECpf.Text;        // Cpf
       SGCadastro.Cells[02,nCel] := MERg.Text;         // Rg
       SGCadastro.Cells[03,nCel] := ENome.Text;        // Nome
       SGCadastro.Cells[04,nCel] := Etelefone.Text;    // Telefone
       SGCadastro.Cells[05,nCel] := Email.Text;        // e-mail
       SGCadastro.Cells[06,nCel] := Ecep.Text;         // Cep
       SGCadastro.Cells[07,nCel] := ELogradouro.Text;  // Logradouro
       SGCadastro.Cells[08,nCel] := ENro.Text;         // Numero
       SGCadastro.Cells[09,nCel] := EComplemento.Text; // Complemento
       SGCadastro.Cells[10,nCel] := EBairro.Text;      // Bairro
       SGCadastro.Cells[11,nCel] := ECidade.Text;      // Cidade
       SGCadastro.Cells[12,nCel] := CBEstado.Text;     // Estado
       SGCadastro.Cells[13,nCel] := CBPais.Text;       // Pais

       nCel := nCel + 1;
       if (nCel = 7) then
          SGCadastro.RowCount := SGCadastro.RowCount+1;

       LimpVar('T');
     end;
end;

procedure TFCadasto.Cabecario();
Begin
  SGCadastro.ColWidths[00] := (7*6);
  SGCadastro.Cells[00,nCel] := 'Codigo';

  SGCadastro.ColWidths[01] := (7*12);
  SGCadastro.Cells[01,nCel] := 'Cpf';

  SGCadastro.ColWidths[02] := (7*15);
  SGCadastro.Cells[02,nCel] := 'Rg';

  SGCadastro.ColWidths[03] := (7*50);
  SGCadastro.Cells[03,nCel] := 'Nome';

  SGCadastro.ColWidths[04] := (7*15);
  SGCadastro.Cells[04,nCel] := 'Telefone';

  SGCadastro.ColWidths[05] := (7*55);
  SGCadastro.Cells[05,nCel] := 'E-mail';

  SGCadastro.ColWidths[06] := (7*8);
  SGCadastro.Cells[06,nCel] := 'Cep';

  SGCadastro.ColWidths[07] := (7*50);
  SGCadastro.Cells[07,nCel] := 'Logradouro';

  SGCadastro.ColWidths[08] := (7*7);
  SGCadastro.Cells[08,nCel] := 'Numero';

  SGCadastro.ColWidths[09] := (7*50);
  SGCadastro.Cells[09,nCel] := 'Complemento';

  SGCadastro.ColWidths[10] := (7*40);
  SGCadastro.Cells[10,nCel] := 'Bairro';

  SGCadastro.ColWidths[11] := (7*40);
  SGCadastro.Cells[11,nCel] := 'Cidade';

  SGCadastro.ColWidths[12] := (7*6);
  SGCadastro.Cells[12,nCel] := 'Estado';

  SGCadastro.ColWidths[13] := (7*20);
  SGCadastro.Cells[13,nCel] := 'Pais';
end;

procedure TFCadasto.LimpVar(cTipo : String);
Begin
  if (cTipo = 'T') then
     Begin
       MECpf.Text := '';        // Cpf
       MERg.Text := '';         // Rg
       ENome.Text := '';        // Nome
       Etelefone.Text := '';    // Telefone
       Email.Text := '';        // e-mail
     end;
  Ecep.Text := '';         // Cep
  ELogradouro.Text := '';  // Logradouro
  ENro.Text := '';         // Numero
  EComplemento.Text := ''; // Complemento
  EBairro.Text := '';      // Bairro
  ECidade.Text := '';      // Cidade

  CBEstado.ItemIndex := 25; // Estado
  LDescUf.Caption := DescEstado( CBEstado.Text );

  CBPais.ItemIndex := 2;  // Pais
end;

procedure TFCadasto.MascCpf();
Begin
  MECpf.Text := Limpa(MECpf.Text);
  if (trim(MECpf.Text) <> '') and
     (length(trim(MECpf.Text)) = 11) then
     Begin
       if (VerificaCpf(MECpf.Text)) then
          MECpf.Text := copy(MECpf.Text,01,3)+'.'+
                        copy(MECpf.Text,04,3)+'.'+
                        copy(MECpf.Text,07,3)+'-'+
                        copy(MECpf.Text,10,2)
       else
          Begin
            application.messagebox('CPF Inválido !!!','Atenção',MB_OK+MB_ICONInformation);
            MECpf.Setfocus;
          end;
     end;
end;

procedure TFCadasto.Cep();
Var cCepp : String;
Begin
  cCepp := Limpa(ECep.Text);
  LimpVar('C');
  ECep.Text := cCepp;

  if (trim(ECep.Text) <> '') and
     (length(trim(ECep.Text)) = 8) then
     Begin
       AlimentaCepJson();

       ECep.Text := copy(ECep.Text,1,5)+'-'+copy(ECep.Text,6,3);
     end;
end;

procedure TFCadasto.AlimentaCepJson();
var URL1, cCepTex : string;
    JsonStreamRet : TStringStream;
    HTTP1 : TIdHTTP;
var nTam, nCont : Integer;
var cCamJson,  cConteudo, cAspa : String;
Begin
  URL1 := '';
  HTTP1 := IdHTTPClep.Create(nil);
  JsonStreamRet := TStringStream.Create('');

  HTTP1.Request.Accept := 'application/json';
  HTTP1.Request.ContentType := 'application/json';
  HTTP1.Request.BasicAuthentication := true;

  URL1 := pchar('http://viacep.com.br/ws/'+ECep.Text+'/json/');

  try
    HTTP1.Get(URL1, JsonStreamRet);
    JsonStreamRet.Position := 0;

    MCepApp.Lines.Clear;
    MCepApp.Text := JsonStreamRet.DataString;
    MCepApp.Text := trim(LimpCa(MCepApp.Text));

    nTam := length(MCepApp.Text);
    cCepTex := MCepApp.Text;

    if (nTam > 4) then
       Begin
         cAspa := 'N';
         for nCont := 0 to nTam do Begin
             if (cAspa = 'N') then
                Begin
                  if (copy(cCepTex,nCont,1) = ':') then
                     cAspa := 'C'
                  else if (copy(cCepTex,nCont,1) <> '"') then
                     cCamJson := cCamJson + copy(cCepTex,nCont,1);  // copy(cCepTex,nCont,1);
                end
             else if (cAspa = 'C') then
                Begin
                  if (copy(cCepTex,nCont  ,1) = ',') then
                     Begin
                       cAspa := ' ';
                       if (cConteudo = '') then
                          Begin
                            cCamJson  := '';
                            cConteudo := '';
                            cAspa := 'N';
                          end;
                     end
                  else if (copy(cCepTex,nCont,1) <> '"') then
                     cConteudo := cConteudo + copy(cCepTex,nCont,1);  // copy(cCepTex,nCont,1);
                end;

             if (cAspa = ' ') and (cCamJson <> '') and (cConteudo <> '') then
                Begin
                  if (cCamJson = 'logradouro') then
                     ELogradouro.Text := cConteudo
                  else if (cCamJson = 'complemento') then
                     EComplemento.Text := cConteudo
                  else if (cCamJson = 'bairro') then
                     EBairro.Text := cConteudo
                  else if (cCamJson = 'localidade') then
                     ECidade.Text := cConteudo
                  else if (cCamJson = 'uf') then
                     Begin
                       CBEstado.ItemIndex := CBEstado.Items.indexof(cConteudo);
                       LDescUf.Caption := DescEstado( CBEstado.Text );
                     end;

                  cCamJson  := '';
                  cConteudo := '';
                  cAspa     := 'N';
                end;
         end;
       end;
  except
    on E:EIdHTTPProtocolException do Begin
      MCepApp.Lines.Clear;

      application.messagebox(pchar(E.ErrorMessage),'Atenção',MB_OK+MB_ICONInformation);
    end;
  end;

end;

Function VerificaCPF(CPF: String): Boolean;
var  TextCPF : String;
     Laco, Soma, Digito1, Digito2 : Integer;
begin
  Result := False;
// retira os caracteres não numéricos
  TextCPF := '';
  for Laco := 1 to Length(CPF) do
      if CPF[Laco] in ['0'..'9'] then
         TextCPF := TextCPF + CPF[Laco];

  if TextCPF = '' then
     Result := True;

  if Length(TextCPF) <> 11 then
     Exit;
// verifica primeiro digito
  Soma := 0;
  for Laco := 1 to 9 do
      Soma := Soma + (StrToInt(TextCPF[Laco])*Laco);

  Digito1 := Soma mod 11;
  if Digito1 = 10 then
     Digito1 := 0;
// verifica segundo digito
  Soma := 0;
  for Laco := 1 to 8 do
      Soma := Soma + (StrToInt(TextCPF[Laco+1])*(Laco));

  Soma := Soma + (Digito1*9);
  Digito2 := Soma mod 11;
  if Digito2 = 10 then
     Digito2 := 0;
// faz a validação
  if Digito1 = StrToInt(TextCPF[10]) then
     if Digito2 = StrToInt(TextCPF[11]) then
        Result := True;
end;

Function DescEstado( Codigo : String ) : string;
var cDescEst: String;
begin
  cDescEst := '                   ';
  if Codigo = 'AC' then cDescEst := 'Acre';
  if Codigo = 'AL' then cDescEst := 'Alagoas';
  if Codigo = 'AM' then cDescEst := 'Amazonas';
  if Codigo = 'AP' then cDescEst := 'Amapa';
  if Codigo = 'BA' then cDescEst := 'Bahia';
  if Codigo = 'CE' then cDescEst := 'Ceara';
  if Codigo = 'DF' then cDescEst := 'Distrito Federal';
  if Codigo = 'ES' then cDescEst := 'Espirito Santo';
  if Codigo = 'GO' then cDescEst := 'Goias';
  if Codigo = 'MA' then cDescEst := 'Maranhão';
  if Codigo = 'MG' then cDescEst := 'Minas Gerais';
  if Codigo = 'MS' then cDescEst := 'Mato Grosso do Sul';
  if Codigo = 'MT' then cDescEst := 'Mato Grosso';
  if Codigo = 'PA' then cDescEst := 'Para';
  if Codigo = 'PB' then cDescEst := 'Paraiba';
  if Codigo = 'PE' then cDescEst := 'Pernambuco';
  if Codigo = 'PI' then cDescEst := 'Piaui';
  if Codigo = 'PR' then cDescEst := 'Parana';
  if Codigo = 'RJ' then cDescEst := 'Rio de Janeiro';
  if Codigo = 'RN' then cDescEst := 'Rio Grande do Norte';
  if Codigo = 'RO' then cDescEst := 'Rondonia';
  if Codigo = 'RR' then cDescEst := 'Roraima';
  if Codigo = 'RS' then cDescEst := 'Rio Grande do Sul';
  if Codigo = 'SC' then cDescEst := 'Santa Catarina';
  if Codigo = 'SE' then cDescEst := 'Sergipe';
  if Codigo = 'SP' then cDescEst := 'Sao Paulo';
  if Codigo = 'TO' then cDescEst := 'Tocantins';
  if Codigo = 'EX' then cDescEst := 'Exterior';
  result := cDescEst;
end;

Function Limpa(str: string): string;
begin
  Result := StringReplace(
              StringReplace(
                StringReplace(
                  StringReplace(
                    StringReplace(
                      StringReplace(
                        StringReplace( str,'.','',[rfReplaceAll]),
                                           '-','',[rfReplaceAll]),
                                           '/','',[rfReplaceAll]),
                                           '\','',[rfReplaceAll]),
                                           '(','',[rfReplaceAll]),
                                           ')','',[rfReplaceAll]),
                                           ' ','',[rfReplaceAll]);
end;

Function LimpCa(str: string): string;
begin
  Result := StringReplace(
              StringReplace(
                StringReplace(
                  StringReplace(
                    StringReplace(
                      StringReplace(
                        StringReplace(
                          StringReplace(
                            StringReplace(
                              StringReplace(
                                StringReplace(
                                  StringReplace(
                                    StringReplace(
                                      StringReplace(
                                        StringReplace(
                                          StringReplace(
                                            StringReplace(
                                              StringReplace(
                                                StringReplace(
                                                  StringReplace( str,'ß','',[rfReplaceAll]),
                                                                     '¢','',[rfReplaceAll]),
                                                                     '¥','',[rfReplaceAll]),
                                                                     'Ð','',[rfReplaceAll]),
                                                                     'Þ','',[rfReplaceAll]),
                                                                     'µ','',[rfReplaceAll]),
                                                                     '¶','',[rfReplaceAll]),
                                                                     'ƒ','',[rfReplaceAll]),
                                                                     '§','',[rfReplaceAll]),
                                                                     '©','',[rfReplaceAll]),
                                                                     'æ','',[rfReplaceAll]),
                                                                     'Ø','',[rfReplaceAll]),
                                                                     '®','',[rfReplaceAll]),
                                                                     '£','',[rfReplaceAll]),
                                                                     '¿','',[rfReplaceAll]),
                                                                     '{','',[rfReplaceAll]),
                                                                     '}','',[rfReplaceAll]),
                                                                 '": ','":',[rfReplaceAll]),
                                                                 ''#$A'','',[rfReplaceAll]),
                                                                '",  ','",',[rfReplaceAll]);
end;

function Padr(cCaracter : String; cTexto : String; nTam : Integer ): String;
begin
  cTexto := Trim(cTexto) ;
  if nTam > length(cTexto) Then
     Result := replicate_chr(cCaracter,(nTam - length(cTexto))) + cTexto
  Else
     Result := cTexto ;
End;

function replicate_chr(sBase : string; iTam : Integer): String;
var iLoop : Integer ;
var sTemp : string ;
begin
  sTemp := '';
  for iLoop := 1 to iTam do
  begin
    sTemp := sTemp + sBase ;
  end;
  Result := sTemp ;
end;

procedure TFCadasto.TBGeraXmlClick(Sender: TObject);
var XmlDoc: TXMLDocument;
var XmlTabela, XmlRegistro, XmlEndereco : IXMLNode;
var cNomeArq, cCodigo : String;
var nCel, nLinhas : Integer;
begin
  cCodigo := trim(SGCadastro.Cells[00,1]);  // Codigo
  if (cCodigo = '') then
     application.messagebox('Cadastro vazio, favor verificar !!!','Atenção',MB_OK+MB_ICONInformation)
  else
     Begin
       OpenDialog1.Title := 'Selecione a pasta';
       OpenDialog1.InitialDir := 'C:\';
       OpenDialog1.Filter     := '*.XML';
       OpenDialog1.FileName   := 'Cadastro.XML';

       if OpenDialog1.Execute then
          Begin
            cNomeArq := OpenDialog1.FileName;

            XmlDoc := TXMLDocument.Create(Self);
            try
              XmlDoc.Active := True;
              XmlTabela := XmlDoc.AddChild('Cadastro');

              nLinhas := SGCadastro.RowCount;
              for nCel := 1 to nLinhas do Begin
                 cCodigo := trim(SGCadastro.Cells[00,nCel]);  // Codigo
                 if (cCodigo <> '') then
                    Begin
                      XmlRegistro := XmlTabela.AddChild('Cliente');
                      XmlRegistro.ChildValues['Id']       := padr('0',cCodigo,3);
                      XmlRegistro.ChildValues['Cpf']      := SGCadastro.Cells[01,nCel];
                      XmlRegistro.ChildValues['Rg']       := SGCadastro.Cells[02,nCel];
                      XmlRegistro.ChildValues['Nome']     := SGCadastro.Cells[03,nCel];
                      XmlRegistro.ChildValues['Telefone'] := SGCadastro.Cells[04,nCel];
                      XmlRegistro.ChildValues['Email']    := SGCadastro.Cells[05,nCel];
                      XmlEndereco := XmlRegistro.AddChild('Endereco');
                      XmlEndereco.ChildValues['Cep']         := SGCadastro.Cells[06,nCel];
                      XmlEndereco.ChildValues['Logradouro']  := SGCadastro.Cells[07,nCel];
                      XmlEndereco.ChildValues['Nro']         := SGCadastro.Cells[08,nCel];
                      XmlEndereco.ChildValues['Complemento'] := SGCadastro.Cells[09,nCel];
                      XmlEndereco.ChildValues['Bairro']      := SGCadastro.Cells[10,nCel];
                      XmlEndereco.ChildValues['Cidade']      := SGCadastro.Cells[11,nCel];
                      XmlEndereco.ChildValues['Estado']      := SGCadastro.Cells[12,nCel];
                      XmlEndereco.ChildValues['Pais']        := SGCadastro.Cells[13,nCel];

//                      XmlDoc.SaveToFile(cNomeArq);
                    end;

                 XmlDoc.SaveToFile(cNomeArq);
              end;
            finally
              XmlDoc.Free;
          end;
     end;
  end;
end;

procedure TFCadasto.TBEmailClick(Sender: TObject);
var nCel, nLinhas, nPorta : Integer;
var cOk, cCodigo, cNomeArq, cEmailEnvi, cEmailOrig,
    cSmtp, cSmtpServ, cSenha, cComCopia, cCCo,
    cEmailConf, cNomeConf, cAssunto, cMensagem : String;
begin
  cOk := 'S';
  cCodigo := trim(SGCadastro.Cells[00,1]);  // Codigo
  if (cCodigo = '') then
     Begin
       application.messagebox('Cadastro vazio, favor verificar !!!','Atenção',MB_OK+MB_ICONInformation);
       cOk := 'N';
     end;

  if (cOk = 'S') and
     (Messagedlg('Confirma envio do e-mail  ?',mtConfirmation,[mbYes,mbNo],0) = mrNo) then
     cOk := 'N';

  cEmailEnvi := '';   //  'NomeEnvio'
  if (cOk = 'S') and
     (cEmailEnvi = '') then
     Begin
       application.messagebox('E-Mail do Destinatario não informado !!!','Atenção',MB_OK+MB_ICONInformation);
       cOk := 'N';
     end;

  if (cOk = 'S') then
     Begin
//       OpenDialog1.Title := 'Selecione a pasta';
//       OpenDialog1.InitialDir := 'C:\';
//       OpenDialog1.Filter     := '*.XML';
//       OpenDialog1.FileName   := 'Cadastro.XML';

       TBGeraXmlClick(Sender);

//       if OpenDialog1.Execute then
//          Begin
       cNomeArq := OpenDialog1.FileName;

       cEmailOrig := '';   //  'EmailOrigems'
       cSmtp      := '';   //  'Smtps'
       cSmtpServ  := '';   //  SmtpServs'
       nPorta     := 0;    //   'Portas'
       cSenha     := '';   //  'Senhas'
       cComCopia  := '';   //  'ComCopias'
       cCCo       := '';   //  'CCOs'
       cEmailConf := '';   //  'EMailConf'
       cNomeConf  := '';   //  'NomeConf'
// alimentar o campo memo com o cadastro
       MCepApp.clear;
       MCepApp.lines.add( 'Cadastro de Cliente;' );
       MCepApp.lines.add( '' );

       nLinhas := SGCadastro.RowCount;
       for nCel := 1 to nLinhas do Begin
          cCodigo := trim(SGCadastro.Cells[00,nCel]);  // Codigo
          if (cCodigo <> '') then
             Begin
               MCepApp.lines.add( 'Cpf .........'+SGCadastro.Cells[01,nCel] );
               MCepApp.lines.add( 'Rg ..........'+SGCadastro.Cells[02,nCel] );
               MCepApp.lines.add( 'Nome ........'+SGCadastro.Cells[03,nCel] );
               MCepApp.lines.add( 'Telefone ....'+SGCadastro.Cells[04,nCel] );
               MCepApp.lines.add( 'Email .......'+SGCadastro.Cells[05,nCel] );
               MCepApp.lines.add( 'Cep .........'+SGCadastro.Cells[06,nCel] );
               MCepApp.lines.add( 'Logradouro ..'+SGCadastro.Cells[07,nCel] );
               MCepApp.lines.add( 'Nro .........'+SGCadastro.Cells[08,nCel] );
               MCepApp.lines.add( 'Complemento .'+SGCadastro.Cells[09,nCel] );
               MCepApp.lines.add( 'Bairro ......'+SGCadastro.Cells[10,nCel] );
               MCepApp.lines.add( 'Cidade ......'+SGCadastro.Cells[11,nCel] );
               MCepApp.lines.add( 'Estado ......'+SGCadastro.Cells[12,nCel] );
               MCepApp.lines.add( 'Pais ........'+SGCadastro.Cells[13,nCel] );
               MCepApp.lines.add( '=========================================================' );
             end;
       end;

       cAssunto := 'Cadastro Cliente '+Copy(dateTimeToStr(now),1,10);
       cMensagem := trim(MCepApp.Text);

       IdMessage1.MessageParts.Clear;
       IdMessage1.Body.Clear;

       IdMessage1.From.Text := cEmailOrig;        // Email de quem vai enviar
       IdMessage1.From.Name := '';                // Nome de quem vai enviar
       IdMessage1.Recipients.EMailAddresses := cEmailEnvi;  // quem ira receber sua mensagem
       if (cComCopia <> '') then
          IdMessage1.CCList.EMailAddresses := cComCopia;      // Com Copia;
       if (cCCo <> '') then
          IdMessage1.BccList.EMailAddresses := cCCo;           // Copia oculta

       IdMessage1.Subject := cAssunto;                      // Assunto da Messangem
       IdMessage1.Body.Add(cMensagem);                      // ('Boa tarde Teste de envio');
       IdMessage1.Priority := mpNormal;                     // Prioridade da Messangem do tipo normal

       if (cNomeArq <> '') then
          TIdAttachment.Create(idmessage1.MessageParts, TFileName(pchar(cNomeArq)));

       IdSMTP1.Port := nPorta;                 // 587;                      // Porta usada
       IdSMTP1.Host := cSmtpServ;              //'smtp.globo.com';          // Servidor de SMTP
       IdSMTP1.Username := cSmtp;                //'losgrobo.nfe@globo.com';  // Usuario do Servidor de SMTP
       IdSMTP1.Password := cSenha;             //'j19d47i30';               // Senha do Servidor de SMTP
       IdSMTP1.AuthenticationType := atlogin;  //Autenticação usada por requerimentos do Servidor de SMTP   ou IdSMTP1.AuthType := atDefault

       if not IdSMTP1.Connected then
          Begin
            IdSMTP1.Connect;
            if IdSMTP1.Connected then
               Begin
                 IdSMTP1.Authenticate();
                 IdSMTP1.Send(IdMessage1);
                 IdSMTP1.Disconnect;

                 application.messagebox('E-Mail enviado ...','Atenção',MB_OK+MB_ICONInformation);
               end
            else
               application.messagebox('E-Mail não enviado erro na transmissão ...','Atenção',MB_OK+MB_ICONInformation);
          end;
//          end;
     end;
end;

end.
