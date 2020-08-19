unit MyComm;

interface

uses Windows, SysUtils, Classes, Controller, ActiveX,ComObj,Variants;


type

  // TComThread - child from the class TThread
  TCommThread = class(TThread)
  private
    { Private declarations }
  //procedure, of reading COM-Port
    Procedure QueryPort;
  protected
  //Method of starting thread
     Procedure Execute; override;
  end;

 function StartService(ComName:Integer):Bool;
 Procedure StopService;
 function WriteStrToPort(Str:String;ctrl_char:Char):boolean;
 function CloseComm : boolean;
 function LogArray(flag:Bool):Bool;

Var
CommThread:TCommThread; //Thread, where we will work with COM-port
hPort:Integer;          //port descriptor
isRecieved:boolean;
time_int:double;
dot_Pos:integer;
ComPort,pressureCounter,LogCounter:Integer;
LogBuff:array [0..200,0..1] of string;
FData:Variant;
LogFlag,Terminated:boolean;
implementation
uses Graph;

//******************************************************************************

//=================================StartComThread===============================

  Procedure StartComThread;
  //initialization of our thread
    Begin {StartComThread}
      //trying to initialize thread
      CommThread:=TCommThread.Create(False);
      CommThread.Priority:=tpLower;
      //checking the result
      If CommThread = Nil Then
        Begin {Nil}
        //Erorr, exit the application
          SysErrorMessage(GetLastError);
          Exit;
        End; {Nil}
    End; {StartComThread}


//=================================StartComThread===============================


//******************************************************************************


//================================Execute=======================================

  //Starting the procedure interviewing the port
  Procedure TCommThread.Execute;
    Begin {Execute}
      Repeat
        QueryPort;
        // Will work until Terminated
        Until Terminated;
    End;  {Execute}


//================================Execute=======================================


//******************************************************************************


//===============================LogArray=========================================


//Sending log for saving to the main form
function LogArray(flag:Bool):Bool;
var
  j,q:Integer;
  begin
            //if flag=true,then function was initialized from main form => LogCounter<200
            if(flag=true) then
              begin
                  CoInitialize(Nil);
                    //Creating variant array to send it to excel application
                    FData:=VarArrayCreate([1,LogCounter+1,1,2],varVariant);
                    for j:=1 to LogCounter+1 do
                        for q:=1 to 2  do
                             FData[j,q]:=LogBuff[j-1,q-1];
                                fmMain.SaveToFile(FData,LogCounter);
                                LogCounter:=0;
                                FillChar(LogBuff,LogCounter,0);
                                time_int:=0;
              end
            else
              begin
              //Initialization from ComThread
                  CoInitialize(Nil);
                    FData:=VarArrayCreate([1,LogCounter+1,1,2],varVariant);
                    for j:=1 to LogCounter+1 do
                        for q:=1 to 2  do
                             FData[j,q]:=LogBuff[j-1,q-1];
                                fmMain.SaveToFile(FData,LogCounter);
                                LogCounter:=0;
                                FillChar(LogBuff,LogCounter,0);
                    Result:=True;
            end;
  end;


//===============================LogArray=========================================


//******************************************************************************


//================================QueryPort=====================================


  //interviewing the port
  Procedure TCommThread.QueryPort;
    Var
      Ovr : TOverlapped;
      Events : array[0..1] of THandle;
      MyBuff:Array[0..255] Of Char;              //Buffer for readed inf
      ByteReaded:Dword;                          //Number of readed bytes
      pressureV1,pressureV2,pressureP,tmp,tmp1:String;
      pPressure,pPressureV1,pPressureV2:pChar;
      ctrl_charPos1,ctrl_charPos2, stop_charPos,
      ctrl_charPos3,ctrl_charPos4:integer;
      flag:Bool;

    Begin {QueryPort}
      flag:=false;
      //Read Buffer from Com-port
      FillChar(Ovr,SizeOf(TOverlapped),0);
      Ovr.hEvent:=CreateEvent(nil,TRUE,FALSE,#0);
      Events[0] := Ovr.hEvent;

        If Not ReadFile(hPort,MyBuff,SizeOf(MyBuff),ByteReaded,@Ovr) Then
          Begin {Error with readed files}
            //Error, close all and exit
            SysErrorMessage(GetLastError);
            Exit;
            CloseHandle(Ovr.hEvent);
          End;{Error with readed files}
      //Data recieved
      If ByteReaded>0 Then
        Begin {ByteReaded>0}
        //Making string from recieved buffer
              tmp:=string(MyBuff);
     // stop_charPos:=0;
     // ctrl_charPos1:=0;
     // ctrl_charPos2:=0;
      ctrl_charPos3:=0;
     // ctrl_charPos4:=0;
      tmp1:=tmp;
                                            //Parsing recieved string
      stop_charPos:=AnsiPos('#',tmp);       //Stop char
      ctrl_charPos1:=AnsiPos('A',tmp);      //Ctrl-char from volume +
      ctrl_charPos2:=AnsiPos('B',tmp);      //Ctrl-char from volume -
      ctrl_charPos3:=AnsiPos('D',tmp);      //Ctrl-char from pippette
      ctrl_charPos4:=AnsiPos('G',tmp);      //Ctrl-char for connecting protocol

      if (ctrl_charPos3<>0) then
        begin
          //Parsing pressure from pippette, and adding it to the main form
          pressureP:=copy(tmp,(ctrl_charPos3)+1,(stop_charPos-ctrl_charPos3)-1);
          pPressure:=PChar(pressureP);
         //PChar for SendMessage function

         if pressureCounter=10 then
         //to refresh the pressure every second,
         //and not to show it every 200 ms
          begin
            SendMessage(fmMain.Handle,cmRxByte,1,Integer(@pPressure));
            //Sending Message with parameters 1 and pressure
            pressureCounter:=0;
          end
         else
          pressureCounter:=pressureCounter+1;

          //If graph form Open, then send data to it
                  if (Assigned(Form2)=true) then
                      begin
                        if (StrToFloat(pressureP)<>0) then
                            try
                              Form2.MyGraph((StrToFloat(pressureP)));
                            except
                            end;
                      end;



          //Sending data for log
          if (LogFlag=True) then
             begin
                //Fill the array of data
                LogBuff[LogCounter,0]:=pressureP;
                time_int:=time_int+Integer((fmMain.Timer1.Interval)); //Time const
                LogBuff[LogCounter,1]:=FloatToStr(time_int/1000);

                //Starting saving function
                if LogCounter = 200 then
                  LogArray(flag)
                else
                  LogCounter:=LogCounter+1;
             end;
        end;

    //Parsing pressure in volumes and sending it to the main form
    if(ctrl_charPos1<>0) then
      begin
        //pressure in the Positive volume
        pressureV1:=copy(tmp,(ctrl_charPos1)+1,(ctrl_charPos2-ctrl_charPos1)-1);
        pPressureV1:=PChar(pressureV1);
        //PChar for SendMessage function

        SendMessage(fmMain.Handle,cmRxByte,2,Integer(@pPressureV1));
        //Sending Message with parameters 2 and pressure

        //pressure in the Negative volume
        pressureV2:=copy(tmp1,(ctrl_charPos2)+1,(stop_charPos-ctrl_charPos2)-1);
        pPressureV2:=PChar(pressureV2);
        //PChar for SendMessage function

        SendMessage(fmMain.Handle,cmRxByte,3,Integer(@pPressureV2));
        //Sending Message with parameters 3 and pressure

      end;

              //Connecting protocol
            if (ctrl_charPos4<>0) then
                  begin
                      SendMessage(fmMain.Handle,cmRxByte,4,Integer(0));
                      DeleteObject(ctrl_charPos4);
                     //Sending COM-start request to the main form
            end;
      End; {ByteReaded>0}
End; {QueryPort}


//================================QueryPort=====================================


//******************************************************************************


//==================================InitPort====================================


  //Initialization of port
  function InitPort:Bool;
    Var
      DCB: TDCB;         //Structure for settings of COM-Port
      CT: TCommTimeouts; //Structure for timeouts of COM-Port
    Begin {InitPort}
        hPort := CreateFile(PChar('\\.\COM' + IntToStr(ComPort)),
                            GENERIC_READ or GENERIC_WRITE,
                            FILE_SHARE_READ or FILE_SHARE_WRITE,
                            nil, OPEN_EXISTING,
                            FILE_ATTRIBUTE_NORMAL, 0);
        If (hPort < 0)                          //Couldn't create file(initialize port)
          Or Not SetupComm(hPort, 256, 256)     //Couldn't set up buffers
          Or Not GetCommState(hPort, DCB) Then //Couldn't get settings of COM-port
            Begin {Error}
              SysErrorMessage(GetLastError);
              Result:=False;
              Exit;
            End;  {Error}

        //Parameters of port
        DCB.BaudRate := CBR_115200; //velocity
        DCB.StopBits := 0;          //stop bits (0 - 1, 1 - 1,5, 2 - 2)
        DCB.Parity := 0;            //parity bits
        DCB.ByteSize := 8;          //bits with information

        If Not SetCommState(hPort, DCB) Then //Couldn't set up settings of COM-port
          Begin {Error}
            SysErrorMessage(GetLastError);
            Result:=False;
            Exit;
          End; {Error}

        //Setting up timeouts
        If Not GetCommTimeouts(hPort, CT) Then //Couldn't get timeouts
          Begin  {Error}
            SysErrorMessage(GetLastError);
            Result:=False;
            Exit;
          End; {Error}
        //Timeouts
        CT.ReadTotalTimeoutConstant := 50;
        CT.ReadIntervalTimeout := 50;
        CT.ReadTotalTimeoutMultiplier := 1;
        CT.WriteTotalTimeoutMultiplier := 10;
        CT.WriteTotalTimeoutConstant := 10;

        If Not SetCommTimeouts(hPort, CT) Then //Couldn't set up timeouts
          Begin {Error}
            SysErrorMessage(GetLastError);
            Result:=False;
            Exit;
          End; {Error}
          Result:=True;
    End;{InitPort}


//==================================InitPort====================================


//******************************************************************************


//================================WriteStrToPort================================


  //Write string to port
  function WriteStrToPort(Str:String; ctrl_char:Char): boolean;
    Var
      ByteWritten:DWord;
      MyBuff:Array[0..255] Of Char;
      charCtrl:String;
    Begin {WriteStrToPort}

      FillChar(MyBuff,SizeOf(MyBuff),0);    //Preparing buffer
      charCtrl:='C'+ctrl_char;             //Creating String with ctrl-chars
      Str:=Str+'&';
      Str:=charCtrl+Str;
      StrPCopy(MyBuff,Str);               //Copying final string,and writing it

      If Not WriteFile(hPort,MyBuff,Length(Str),ByteWritten,nil) Then
        Begin {Error}
          SysErrorMessage(GetLastError);
          Result:=false;
          Exit;
        End; {Error}
      Result:=true;
    End; {WriteStrToPort}


//================================WriteStrToPort================================


//******************************************************************************


//=====================================StopService==============================

  //Stop COM-port
  Procedure StopService;
    Begin {StopService}
      fmMain.Timer1.Enabled:=false;        //Disabling timers
      fmMain.Timer2.Enabled:=false;
      WriteStrToPort('99','G');           //Resetting Arduino
      sleep(200);
      CloseComm();                        //Close COM-port
      sleep(100);
    End; {StopService}


//==================================StopService==================================


//******************************************************************************


//==================================CloseComm====================================

  //Close COM-Port
  function CloseComm : boolean;
    begin
     SetCommMask(hPort,0);
     sleep(10);
     //Close handle
     CloseHandle(hPort);
     DeleteObject(hPort);
     sleep(100);
     CommThread.Free;                    //Thread free
     Result := True;
  end;


//=====================================CloseComm=================================


//******************************************************************************


//====================================StartService==============================


  //Start COM-port service
  function StartService(ComName:Integer):Bool;
    Begin {StartService}
  ComPort:=ComName;
            if  InitPort = true then         //Initialization of port
              begin
                StartComThread;             //Starting COM thread
                Result:=true;
              end
            else
              begin                        //If unsuccesfull(error,connected with portNum), then
                CloseHandle(hPort);        //Closing and deleting object
                DeleteObject(hPort);
                Result:=false;
              end;
    End;  {StartService}

//====================================StartService==============================


//******************************************************************************
end.
