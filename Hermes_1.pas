{Made by Yakovlev Aleksey, yakovlevap74@mail.ru}
program Hermes_1;

uses
  Forms, Interfaces,
  Controller in '..\New_Vers\Controller.pas' {Form1},
  Graph in '..\New_Vers\Graph.pas' {Form2};

{$R *.res}

begin
  Application.Scaled:=True;
  Application.Initialize;
  Application.CreateForm(TForm1, fmMain);
  Application.Run;
end.
