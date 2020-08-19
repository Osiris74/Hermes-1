unit Controller;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Forms,
  Buttons, ComObj, Menus, ExtCtrls, StdCtrls, Controls;

const
  cmRxByte = wm_User+$55;

type
  TForm1 = class(TForm)
    ReqPressureEdit: TEdit;
    okButton: TButton;
    standartRadio: TRadioButton;
    patchRadio: TRadioButton;
    cancelButton: TButton;
    Label1: TLabel;
    stopButton: TBitBtn;
    runButton: TBitBtn;
    PlVolume: TEdit;
    NegVolume: TEdit;
    Timer1: TTimer;
    Timer2: TTimer;
    MainMenu1: TMainMenu;
    File1: TMenuItem;
    Settings1: TMenuItem;
    COM1: TMenuItem;
    Graph1: TMenuItem;
    Label2: TLabel;
    Label3: TLabel;
    Label5: TLabel;
    About1: TMenuItem;
    Help1: TMenuItem;
    resetBtn: TSpeedButton;
    startBtn: TSpeedButton;
    edPipPressure: TEdit;
    lbCom: TLabel;
    stopBtn: TButton;
    N1: TMenuItem;
    procedure FormClose(Sender: TObject);
    procedure okButtonClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure runButtonClick(Sender: TObject);
    procedure stopButtonClick(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure Timer2Timer(Sender: TObject);
    procedure Graph1Click(Sender: TObject);
    procedure SaveToFile(FData:Variant;LogCounter:Integer);
    procedure About1Click(Sender: TObject);
    procedure RecivBytes(var Msg : TMessage); message cmRxByte;
    procedure cancelButtonClick(Sender: TObject);
    procedure resetBtnClick(Sender: TObject);
    procedure startBtnClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure ReqPressureEditKeyPress(Sender: TObject; var Key: Char);
    procedure stopBtnClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;


var
  fmMain: TForm1;
  time_const:double;
  isOpenHandle,isStopped:Boolean;
  cellCounter,LogCounter, Num, rowCounter:Integer;
  ExcelApp, ExcelSheet: OLEVariant;
  timeStartLog:TDateTime;
  CurrDate, TempDate : TDate;            // текущая дата


implementation
uses MyComm, Graph;
{$R *.dfm}

//=================================FormClose=====================================

  procedure TForm1.FormClose(Sender: TObject);
  begin
    if not isStopped then
    StopService;
  end;

//=================================FormClose=====================================


//*******************************************************************************


//=================================OkButtonClick=================================

  procedure TForm1.okButtonClick(Sender: TObject);
    var
      S:string;
      ctrl_char:Char;
  begin
    S:=Trim(ReqPressureEdit.Text);     //Get Required Pressure without spaces
    if standartRadio.Checked then      //Cheks Radio-button
    ctrl_char:='S';                    //Ctrl-Char for standart scan
    if patchRadio.Checked then
    ctrl_char:='P';                    //Ctrl-Char for patch
    WriteStrToPort(S,ctrl_char) ;      //Writing bytes to COM port
  end;


//==================================OkButtonClick==================================

//*******************************************************************************

//==================================GetComPort====================================

//Get Computer COM-Ports
  function GetComPort:BOOL;
    var
      ctrl_char:Char;
      Str:String;
      i:integer;
  begin
    ctrl_char:='F';             // Ctrl-char for
    Str:='99';                  // starting connection
    isOpenHandle:=False;        // Down the Flag of working COM

     for i:=1 to 10 do          // Starting loop for scanning COM-Ports
    begin
        if StartService(i) = true then
          begin
            sleep(2000);

           if not WriteStrToPort(Str,ctrl_char) then  //Trying write to port
             continue;

             sleep(1000);
             Application.ProcessMessages;

              if(isOpenHandle=True) then    //If port Opened, then Close it
               begin                       //Break the loop, and returning true
                   CloseComm();
                   Result:=true;
                   Num:=i;
                   break;
               end

              else
               begin
                  CloseComm();
                  Result:=false;
               end;

          end;
    end;
  end;

//==================================GetComPort====================================

//********************************************************************************

//==================================FilePreparation===============================


//Preparing Directory and files function
  function FilePrep:BOOL;
      var
        DirPath:String;
  begin
    DirPath:='C:\Pressure_Controller';
    if not DirectoryExists(DirPath) then
      begin
        CreateDir(DirPath);
      end;
    Result:=True;
  end;

//==================================FilePreparation===============================

//********************************************************************************

//==================================FormCreate====================================


  procedure TForm1.FormCreate(Sender: TObject);
    begin
            cellCounter:=0;
            FilePrep;                 //Preparation of log-files
            fmMain.Show;
    end;

//=====================================FormCreate================================


//********************************************************************************


//================================NameFormatting=================================


 function LogFileName(const ATime: TDateTime): string;
  begin
    Result := 'C:\Pressure_Controller\'+
            FormatDateTime('yyyy_mm_dd_hh_mm_ss',ATime)+'.xlsx';
  end;

//================================NameFormatting=================================


//*******************************************************************************


//=================================RunButtonClick================================


    procedure TForm1.runButtonClick(Sender: TObject);
      begin
        cellCounter:=0;
        runButton.Enabled:=False;
        stopButton.Enabled:=True;
        CurrDate := Now;

            //Creating OLE-object Excel
            ExcelApp:=CreateOleObject('Excel.Application');

            //Disable all notifications
            ExcelApp.Application.EnableEvents:=False;
            ExcelApp.DisplayAlerts:=False;

            ExcelApp.WorkBooks.Add;

            ExcelApp.WorkBooks.Item[1].SaveAs(LogFileName(CurrDate));


            ExcelApp.Quit;
            ExcelApp:=Unassigned;
            ExcelSheet:=Unassigned;


        Mycomm.LogFlag:=True;      //Flag for saving data to log

      end;

//=================================RunButtonClick================================

//******************************************************************************

//================================StopButtonClick================================

  procedure TForm1.stopButtonClick(Sender: TObject);
    begin
      MyComm.time_int:=0;
      stopButton.Enabled:=False;
      runButton.Enabled:=True;
      MyComm.LogFlag:=False;
      MyComm.LogArray(true);
      MyComm.time_int:=0;
    end;

//================================StopButtonClick================================

//********************************************************************************

//===================================SaveToFile===================================

  //Saving log-file(called from MyComm)
  procedure TForm1.SaveToFile(FData:Variant;LogCounter:Integer);
      var
        i:Integer;
   begin

    ExcelApp:=CreateOleObject('Excel.Application');

    ExcelApp.Visible:=False;

    ExcelApp.DisplayAlerts:=False;

    ExcelApp.Application.EnableEvents:=False;

    //Open existing log-file

    try
      ExcelApp.Workbooks.Open(LogFileName(CurrDate));

      ExcelSheet:=ExcelApp.Workbooks[1].WorkSheets[1];

      //Writing to log in loop
       for i:=1 to LogCounter do
        begin
          CellCounter:=CellCounter+1;
          ExcelSheet.Cells[cellCounter,1].Value:=FData[i,1];
          ExcelSheet.Cells[cellCounter,2].Value:=FData[i,2];
        end;

      ExcelApp.WorkBooks.Item[1].SaveAs(LogFileName(CurrDate));

      finally

        ExcelApp.Quit;

        ExcelApp:=Unassigned;

        ExcelSheet:=Unassigned;
      end;

   end;


//===================================SaveToFile===================================

//********************************************************************************

//=============================Timer_For_Pressure_In_Pipette=====================


  procedure TForm1.Timer1Timer(Sender: TObject);
    var
      ctrl_char:Char;
      S:String;
    begin
      Timer1.Enabled:=false;
      ctrl_char:='A';           //Ctrl-char for pressure in pippette
      S:='11111';
      if WriteStrToPort(S,ctrl_char) then Timer1.Enabled:=true;
end;


//=============================Timer_For_Pressure_In_Pipette=====================

//*******************************************************************************

//=============================Timer_For_Pressure_In_Volume======================


  procedure TForm1.Timer2Timer(Sender: TObject);
    var
      ctrl_char :Char;
      S:String;
    begin
      Timer1.Enabled:=false;
      Timer2.Enabled:=false;
      ctrl_char:='B';           //Ctrl-char for pressure in volume
      S:='11111';
      if WriteStrToPort(S,ctrl_char) then
        begin
          Timer2.Enabled:=true;
          Timer1.Enabled:=true;
        end;
    end;


//=============================Timer_For_Pressure_In_Volume======================

//*******************************************************************************

//==============================GraphClick=======================================

  //Open Graph form
  procedure TForm1.Graph1Click(Sender: TObject);
    begin
      if (not Assigned(Form2)) then  // Cheking Existing Form2, if not
       Form2:=TForm2.Create(Self);   // creating it
       Form2.Show;                   // showing it
    end;


//==============================GraphClick=======================================

//*******************************************************************************

//=============================AboutButton=======================================

    procedure TForm1.About1Click(Sender: TObject);
        begin
            MessageBox(0,'Hermes-1 v1.0'+#13#10+'Made by A.P.Yakovlev'+
              #13#10+'For more information,please contact me:'+#13#10+
              'yakovlevap74@mail.ru', 'About',0);
        end;


//=============================AboutButton=======================================


//********************************************************************************


//===================================RecivBytes===================================

      procedure TForm1.RecivBytes(var Msg: TMessage);
       var
        s:PChar;
          begin
              case Msg.WParam of

                 //Message with pipette pressure
                 1: begin
                      s:=PChar(Pointer(Msg.LParam)^);
                      edPipPressure.Text:=string(s);
                      Application.ProcessMessages;
                    end;

                 //Message with volume+ pressure
                 2: begin
                      s:=PChar(Pointer(Msg.LParam)^);
                      PlVolume.Text:=string(s);
                      Application.ProcessMessages;
                    end;

                 //Message with volume- pressure
                 3: begin
                       s:=PChar(Pointer(Msg.LParam)^);
                       NegVolume.Text:=string(s);
                       Application.ProcessMessages;
                    end;

                 //Connecting protocol
                 4: begin
                       isOpenHandle:=true;
                       Application.ProcessMessages;
                    end;
              end;
          end;

//===================================RecivBytes===================================


//********************************************************************************


//===================================CancelButton=================================
    procedure TForm1.cancelButtonClick(Sender: TObject);
        begin
             ReqPressureEdit.Clear;
        end;

//===================================CancelButton=================================


//********************************************************************************


//==================================ResetButton===================================


        procedure TForm1.resetBtnClick(Sender: TObject);
            begin
                StopService;
                isStopped:=True;
                startBtn.Visible:=True;
                resetBtn.Visible:=False;
                lbCom.Caption:='COM-port service stopped...';
                edPipPressure.Text:='';
            end;

//==================================ResetButton===================================


//********************************************************************************


//==================================StartButton===================================


      procedure TForm1.startBtnClick(Sender: TObject);

          begin
               StartService(Num);
               Timer1.Enabled:=True;       //Starting timers for requesting pressure
               Timer2.Enabled:=True;
               edPipPressure.Text:='';
               lbCom.Caption:='COM-port service started...';
               resetBtn.Visible:=True;
               startBtn.Visible:=False;
          end;


//==================================StartButton===================================


//********************************************************************************


//==================================FormShow=======================================

      procedure TForm1.FormShow(Sender: TObject);
        begin
            //Automatical protocol for connecting arduino
            if GetComPort = true then
              begin
                  StartService(Num);
                  Application.ProcessMessages;
                  //If succesfully, then starts Service
                  MessageBox(0,'Successfully','Connecting',0);

                  lbCom.Caption:='COM-port service started...';

                  Timer1.Enabled:=True;       //Starting timers for requesting pressure
                  Timer2.Enabled:=True;
              end
            else
              MessageBox(0,'Error with connecting protocol','Connecting',0);
        end;

//======================================FormShow====================================


//**********************************************************************************


//=============================Formatting the string================================


    procedure TForm1.ReqPressureEditKeyPress(Sender: TObject; var Key: Char);
      const
        Digit: set of Char=['0'..'9',Char($8),Char($2E),Char($2D)];
        // Char($8),Char($2E)  backspace and delete buttons
            begin
                 if (not (Key in Digit)) then
                  Key:=#0;
            end;

//=============================Formatting the string================================


//******************************************************************************


//==============================StopBtn=========================================


    procedure TForm1.stopBtnClick(Sender: TObject);
        var
          S:String;
          ctrl_char:char;
        begin
             ctrl_char:='S';
             S:=('200');
             WriteStrToPort(S,ctrl_char);
        end;


//==============================StopBtn=========================================


//******************************************************************************



end.
