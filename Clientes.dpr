program Clientes;

uses
  Forms,
  Cadastro in 'Cadastro.pas' {FCadasto};

{$R *.res}

begin
  Application.Initialize;
  Application.Title := 'Cadastro';
  Application.CreateForm(TFCadasto, FCadasto);
  Application.Run;
end.
