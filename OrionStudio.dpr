program OrionStudio;

uses
  Forms,
  main in 'main.pas' {Form1},
  osutil in 'osutil.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.Title := 'Orion Studio';
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
