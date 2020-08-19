unit Graph;

interface

uses
  Windows, Messages, SysUtils,  Classes,  Forms,
  ExtCtrls, StdCtrls, Buttons, Controls;

type
  TForm2 = class(TForm)
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    BitBtn3: TBitBtn;
    BitBtn4: TBitBtn;
    procedure FormCreate(Sender: TObject);
    procedure PaintGraf;
    procedure MyGraph(y2:Double);
    procedure BitBtn1Click(Sender: TObject);
    procedure BitBtn3Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
 end;

 const
 TIKS=25;

 type
 TMas=array[1..TIKS] of Integer;

var
  Form2: TForm2;
  CoorMas: TMas;        // Buffer with data
  counter: integer;
  scale:double;
  koef:integer;
  pressure:double;

implementation
uses MyComm,Controller;
{$R *.dfm}

//================================DrawLine========================================

  // procedure drawing horizontal line
  procedure DrawLine(dc:HDC; x1,y1,x2,y2:integer);
    begin
      MoveToEx(dc,x1,y1,nil);
      LineTo(dc,x2,y2);
    end;

//================================DrawLine========================================


//********************************************************************************


//==============================VerticalLine========================================

  //procedure drawing vertical line
  procedure VerticalLine(dc:HDC; x1,y1,x2,y2:integer);
    begin
      MoveToEx(dc,x1,y1,nil);
      LineTo(dc,x2,y2);
    end;

//==============================VerticalLine========================================


//**********************************************************************************


//===============================MyGraph==========================================


  //Fill buffer with coordinates of dots
  procedure TForm2.MyGraph(y2:Double);
    var
      j:Integer;
    begin
      if counter>TIKS then begin
        for j:=1 to TIKS-1 do                //Fill the buffer in loop
          CoorMas[j]:=CoorMas[j+1];
          CoorMas[TIKS]:=Round(y2);
        end
      else
        CoorMas[counter]:=Round(y2);
        inc(counter);
        if counter>40 then counter:=26;      //Resetting the counter for long-term work
        PaintGraf;                           //Calling procedure of drawing graph
    end;


//====================================MyGraph======================================


//********************************************************************************


//====================================PaintGraf==================================


    //Paint graph
    procedure TForm2.PaintGraf;
     var
        b1,b2,b3: HBRUSH;
        x,j:integer;
        dc: HDC;
        f: HFONT;
        lf: TLogFont;
        tmpPressure:double;
        textPressure:double;
    begin
      Form2.Repaint;                             //Repaint form with
      dc:= GetDC(Form2.Handle);                  //new data
      b1:= CreatePen(PS_SOLID,2,RGB(255,0,0));   //Creating brushes for axis
      b2:= CreatePen(PS_DASH,1,RGB(55,155,55));  //lines
      b3:= CreatePen(PS_SOLID,2,RGB(0,0,210));   //and graph


      SelectObject(dc,b2);                          //Drawing additional axis

      pressure:=Trunc((((Form2.Height -100) div 2)) * scale * 0.6625);  //Setting up
                                                                        //the pressure
      tmpPressure:=pressure;
      textPressure:=pressure / 5;                //Pressure for one segment

      for j:=1 to 10 do                          //Drawing and filling axis
        begin
          DrawLine(dc,0,((Form2.Height-100) div 10)*j,Form2.Width,((Form2.Height-100) div 10 )*j);
          TextOut(dc,10,((Form2.Height-100) div 10)*j,PAnsiChar(FloatToStr(pressure)),6);
          pressure:=pressure-textPressure;
        end;

      pressure:=tmpPressure;                        //This parametr for non-resetting
                                                    //pressure Param

      for j:=1 to 16 do                             //Vertical additional axis
        begin
          VerticalLine(dc,(Form2.Width div 15)*j,0,(Form2.Width div 15)*j,Form2.Height);
        end;

    lf.lfHeight:=   14;                         //Font settings
    lf.lfEscapement:=0;
    lf.lfWidth:=7;
    lf.lfOrientation:=0;
    lf.lfWeight:=0;
    lf.lfCharSet:=ANSI_CharSet;
    lf.lfItalic:= 0;
    lf.lfUnderline:= 0;
    lf.lfStrikeOut:= 0;
    f:= CreateFontIndirect(lf);
    SelectObject(dc,f);                         //Writing text pressure
    TextOut(dc,50,Form2.Height-75,'200 ms',6);  //time interval


    SelectObject(dc,b3);                        //Drawing the graph
    x  := 20;

    //Moving the pen to the beging of the form
    MoveToEx(dc,x,Round((((Form2.Height-100) div 2) - (CoorMas[1] / scale))),nil);

    for j := 2 to TIKS do begin
      x := x + 20;
        //Drawing lines from one element to another
      LineTo(dc,x,Round(((Form2.Height-100) div 2) - (CoorMas[j] / scale)));
    end;

    DeleteObject(SelectObject(dc,b1));               //Deleting brushes
    DeleteObject(SelectObject(dc,b2));
    DeleteObject(SelectObject(dc,b3));

  end;

//====================================PaintGraf==================================


//*******************************************************************************


//====================================FormCreate=================================
  procedure TForm2.FormCreate(Sender: TObject);
    var
      x0,y0:integer;
    begin
      // Setting up some paarmeters for drawing the graph
      counter := 1;
      scale := 1;
    end;


//******************************************************************************

//================================Scale for y-axis===============================

  procedure TForm2.BitBtn1Click(Sender: TObject);
    begin
      scale := scale * 1.5;
      pressure:=((((Form2.Height - 100) div 2)) * scale); //Maximum pressure
                                                         //for this scale
    end;

  procedure TForm2.BitBtn3Click(Sender: TObject);
    begin
      scale:=scale / 1.5;
      if(scale=0) then scale := 10;
      pressure:=((((Form2.Height - 100) div 2)) * scale);
    end;


//==============================Scale for y-axis=================================

//******************************************************************************

//==============================                        ========================


end.
