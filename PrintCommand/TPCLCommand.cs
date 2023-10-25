using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PrintCommand
{
    /// <summary>
    /// TPCL Command Class 입니다
    /// </summary>
    public class TPCLCommand
    {
        /// <summary>
        /// 중간 사이즈 라벨 출력 메서드
        /// </summary>
        /// <param name="labelSizeX">라벨 사이즈 X축(넓이)</param>
        /// <param name="labelSizeY">라벨 사이즈 Y축(높이)</param>
        /// <param name="PrintSizeX">프린팅 영역 X축(넓이)</param>
        /// <param name="PrintSizeY">프린팅 영역 Y축(높이)</param>
        /// <param name="groundX">지역 데이터의 위치 X축(넓이)</param>
        /// <param name="groundY">지역 데이터의 위치 Y축(높이)</param>
        /// <param name="inputText">넣어줄 textInput</param>
        /// <returns></returns>
        public string _MiddleLabelCommand(double labelSizeX, double labelSizeY, double PrintSizeX, double PrintSizeY, double groundX, double groundY, string inputText)
        {
            StringBuilder builder = new StringBuilder();
            //클리어
            builder.Append(_SetClearImageBuffer());
            //라벨 사이즈 지정
            builder.Append(_SetLabelSize(labelSizeY,labelSizeX,PrintSizeY,PrintSizeX));
            // 텍스트 위치 지정1
            builder.Append(_SetTrueFont(01, groundX, groundY, 80, 80, "E", 90, "B"));
            // 텍스트 인풋
            builder.Append(_SetTrueValueInput(01, inputText));
            // Event 진행
            builder.Append(_SetStartPrinting(1,0,1,0,1,0,0,0));

            return builder.ToString();
        }
        #region Define Consts
        public const byte pointUP = 10;
        const char NUL = (char)0;
        const char SOH = (char)1;
        const char STX = (char)2;
        const char ETX = (char)3;
        const char EOT = (char)4;
        const char ENQ = (char)5;
        const char ACK = (char)6;
        const char BEL = (char)7;
        const char BS = (char)8;
        const char HT = (char)9;
        const char LF = (char)10;
        const char VT = (char)11;
        const char FF = (char)12;
        const char CR = (char)13;
        const char SO = (char)14;
        const char SI = (char)15;
        const char DLE = (char)16;
        const char DC1 = (char)17;
        const char DC2 = (char)18;
        const char DC3 = (char)19;
        const char DC4 = (char)20;
        const char NAK = (char)21;
        const char SYN = (char)22;
        const char ETB = (char)23;
        const char CAN = (char)24;
        const char EM = (char)25;
        const char SUB = (char)26;
        const char ESC = (char)27;
        const char FS = (char)28;
        const char GS = (char)29;
        const char RS = (char)30;
        const char US = (char)31;
        const char SPACE = (char)32;
        #endregion

        /// <summary>
        /// [ESC] 를 뜻하며 구문의 시작 = "{"
        /// </summary>
        public string _StartCommand { get; private set; } = "{";
        /// <summary>
        /// [LF][NUL] 구문의 끝 = "|}"
        /// </summary>
        public string _EndCommand { get; private set; } = "|}";
        /// <summary>
        /// [LF] 구문 And = "|"
        /// </summary>
        public string _AndCommand { get; private set; } = "|";

        //C
        /// <summary>
        /// 초기 이미지버퍼 클리어 메서드
        /// </summary>
        /// <returns></returns>
        public string _SetClearImageBuffer()
        {
            StringBuilder builder = new StringBuilder();
            builder.Append(_StartCommand);
            builder.Append("C");
            builder.Append(_EndCommand);
            return builder.ToString();
        }

        //D
        /// <summary>
        /// 라벨 사이즈 지정 메서드 각 라벨의 높이,넓이,인쇄구간의 높이,넓이 를뜻함
        /// </summary>
        /// <param name="LabelHeight">라벨 높이</param>
        /// <param name="LabelWidth">라벨 넓이</param>
        /// <param name="PrintingHeight">인쇄 영역 높이</param>
        /// <param name="PrintingWidth">인쇄 영역 넓이</param>
        /// <returns></returns>
        public string _SetLabelSize(double LabelHeight, double LabelWidth, double PrintingHeight, double PrintingWidth)
        {
            StringBuilder builder = new StringBuilder();
            builder.Append(_StartCommand);

            builder.Append("D");
            builder.AppendFormat("{0:0000},{1:0000},{2:0000},{3:0000}", LabelHeight, PrintingWidth, PrintingHeight, LabelWidth);

            builder.Append(_EndCommand);
            return builder.ToString();
        }

        //AX
        /// <summary>
        /// 라벨의 상세 조정 메서드, true = "+"(plus), false = "-"(minus)
        ///  0~500, 0~500,0~99 단위 : 1=0.1mm
        /// </summary>
        /// <param name="LabelStartAdjust">시작 포지션 위치조정 방향</param>
        /// <param name="StartPositionValue">시작 포지션 위치조정 값</param>
        /// <param name="LabelCutAdjust">커트 포지션 위치조정 방향</param>
        /// <param name="CutPositionValue">커트 포지션 위치조정 값</param>
        /// <param name="LabelBackFeedAdjust">후방 포지션 위치조정 방향</param>
        /// <param name="BackFeedPositionValue">후방 포지션 위치조정 값</param>
        /// <returns></returns>
        public string _SetFineFositionCommand(bool LabelStartAdjust, double StartPositionValue, bool LabelCutAdjust, double CutPositionValue, bool LabelBackFeedAdjust, double BackFeedPositionValue)
        {
            char startAdjustment = LabelStartAdjust ? '+' : '-';
            char cutAdjustment = LabelCutAdjust ? '+' : '-';
            char backFeedAdjustment = LabelBackFeedAdjust ? '+' : '-';

            StringBuilder builder = new StringBuilder();
            builder.Append(_StartCommand)
                   .Append("AX;")
                   .Append(startAdjustment)
                   .Append(StartPositionValue.ToString("000"))
                   .Append(",")
                   .Append(cutAdjustment)
                   .Append(CutPositionValue.ToString("000"))
                   .Append(",")
                   .Append(backFeedAdjustment)
                   .Append(BackFeedPositionValue.ToString("00"))
                   .Append(_EndCommand);

            return builder.ToString();
        }
        
        //AY
        /// <summary>
        /// 열전도 방법과, 인쇄 밀도를 지정합니다.
        /// density가 true면, Darker, false면 Light
        /// densityValue = 0~10 Lv, thermal가 true면 열전도, false면 직접 열전도를 시행합니다
        /// </summary>
        /// <param name="densityValue"></param>
        /// <param name="thermal"></param>
        /// <returns></returns>
        public string _SetPrintDensity(string densityValue, bool thermal)
        {
            
            int thermalChoice = thermal ? 0 : 1;

            StringBuilder builder = new StringBuilder();
            builder.Append(_StartCommand)
                   .Append("AY;")
                   .Append(densityValue)
                   .Append(",")
                   .Append(thermalChoice.ToString())
                   .Append(_EndCommand);

            return builder.ToString();
        }

        //RM
        /// <summary>
        /// 리본 모터 Volt출력값을 조정합니다
        /// input = 1 당 5% 감소 최대 15 (45%)
        /// </summary>
        /// <param name="ribbonVolt"></param>
        /// <param name="ribbonBackVolt"></param>
        /// <returns></returns>
        public string _SetMotorVoltCommand(double ribbonVolt, double ribbonBackVolt)
        {
            if (ribbonVolt > 15)
            {
                ribbonVolt = 15;
            }
            if (ribbonBackVolt > 15)
            {
                ribbonBackVolt = 15;
            }

            StringBuilder builder = new StringBuilder();
            builder.Append(_StartCommand)
                   .Append("RM;")
                   .Append('-')
                   .Append(ribbonVolt.ToString("00"))
                   .Append('-')
                   .Append(ribbonBackVolt.ToString("00"))
                   .Append(_EndCommand);

            return builder.ToString();
        }
        
        //XR
        /// <summary>
        /// 특정위치의 영역대를 지정합니다.
        /// startX,Y의 좌표로부터 endX,Y의 좌표까지 사각형으로 지정
        /// TypeClear ? True = 지정된 영역을 Clear , False = 지정된 영역을 White,Black 패턴 반전
        /// </summary>
        /// <param name="startX"></param>
        /// <param name="startY"></param>
        /// <param name="endX"></param>
        /// <param name="endY"></param>
        /// <param name="TypeClear"></param>
        /// <returns></returns>
        public string _SetStartEndPoint(double startX, double startY, double endX, double endY, bool TypeClear)
        {
            string TypeClearChoice = TypeClear ? "A" : "B";

            StringBuilder builder = new StringBuilder();
            builder.Append(_StartCommand)
                   .Append("XR;")
                   .Append(startX.ToString("0000"))
                   .Append(",")
                   .Append(startY.ToString("0000"))
                   .Append(",")
                   .Append(endX.ToString("0000"))
                   .Append(",")
                   .Append(endY.ToString("0000"))
                   .Append(",")
                   .Append(TypeClearChoice)
                   .Append(_EndCommand);

            return builder.ToString();
        }
        
        //LC
        /// <summary>
        /// 포지션을 지정하여 선을 그립니다.
        /// TypeLine = true면 Line false면 Rectangle
        /// lineWidth 는 선두깨 입니다
        /// </summary>
        /// <param name="startX"></param>
        /// <param name="startY"></param>
        /// <param name="endX"></param>
        /// <param name="endY"></param>
        /// <param name="TypeLine"></param>
        /// <param name="lineWidth"></param>
        /// <returns></returns>
        public string _SetSelectedLine(double startX, double startY, double endX, double endY, bool TypeLine, double lineWidth)
        {
            string TypeClearChoice = TypeLine ? "0" : "1";
            if (lineWidth > 9)
            {
                lineWidth = 9;
            }

            StringBuilder builder = new StringBuilder();
            builder.Append(_StartCommand)
                   .Append("LC;")
                   .Append(startX.ToString("0000"))
                   .Append(",")
                   .Append(startY.ToString("0000"))
                   .Append(",")
                   .Append(endX.ToString("0000"))
                   .Append(",")
                   .Append(endY.ToString("0000"))
                   .Append(",")
                   .Append(lineWidth.ToString("0"))
                   .Append(_EndCommand);

            return builder.ToString();
        }

        //PC
        /// <summary>
        /// 비트맵 글꼴 형식을 지정합니다
        /// </summary>
        /// <param name="textNumber">textGroup을 지정 합니다 (0~199)</param>
        /// <param name="positionX">X좌표를 지정 합니다(0~9999) 1=0.1mm</param>
        /// <param name="positionY">Y좌표를 지정 합니다(0~9999) 1=0.1mm</param>
        /// <param name="textHorizontalMargin">가로 간격을 지정 합니다(1~9 or 05~95 = 0.5~9.5)</param>
        /// <param name="textVerticalMargin">새로 간격을 지정 합니다(1~9 or 05~95 = 0.5~9.5)</param>
        /// <param name="selectedFont"> FontStyle을 지정합니다
        /// [  "A" = Times Roman (Medium) 8 point  ]
        /// [  "B" = Times Roman (Medium) 10 point  ]
        /// [  "C" = Times Roman (Bold) 10 point  ]
        /// [  "D" = Times Roman (Bold) 12 point  ]
        /// [  "E" = Times Roman (Bold) 14 point  ]
        /// [  "F" = Times Roman (Italic) 12 point  ]
        /// [  "G" = Helvetica (Medium) 6 point  ]
        /// [  "H" = Helvetica (Medium) 10 point  ]
        /// [  "I" = Helvetica (Medium) 12 point  ]
        /// [  "J" = Helvetica (Bold) 12 point  ]
        /// [  "K" = Helvetica (Bold) 14 point  ]
        /// [  "L" = Helvetica (Italic) 12 point  ]
        /// [  "q" = Gothic725 Black</param>  ]
        /// <param name="rotate">문자열의 각도를 지정합니다 0,90,180,270도</param>
        /// <param name="textOption">문자 특성 "B" = 기본 Black Font</param>
        /// <returns></returns>
        public string _SetBitmapFont(double textNumber, double positionX, double positionY, double textHorizontalMargin, double textVerticalMargin, string selectedFont, int rotate, string textOption = "B")
        {
            string rotateValue = string.Empty;
            switch (rotate)
            {
                case 0:
                    rotateValue = "00";
                    break;
                case 90:
                    rotateValue = "11";
                    break;
                case 180:
                    rotateValue = "22";
                    break;
                case 270:
                    rotateValue = "33";
                    break;
            }


            StringBuilder builder = new StringBuilder();
            builder.Append(_StartCommand)
                   .Append("PC")
                   .Append(textNumber.ToString("000"))
                   .Append(";")
                   .Append(positionX.ToString("0000"))
                   .Append(",")
                   .Append(positionY.ToString("0000"))
                   .Append(",")
                   .Append(textHorizontalMargin.ToString())
                   .Append(",")
                   .Append(textVerticalMargin.ToString())
                   .Append(",")
                   .Append(selectedFont)
                   .Append(",")
                   .Append("+00")
                   .Append(",")
                   .Append(rotateValue)
                   .Append(",")
                   .Append(textOption)
                   .Append(_EndCommand);

            return builder.ToString();
        }

        //PV
        /// <summary>
        /// font형식을 지정합니다
        /// </summary>
        /// <param name="textNumber">textGroupNumber (00~99) </param>
        /// <param name="positionX">StartPoint X Value</param>
        /// <param name="positionY">StartPoint Y Value</param>
        /// <param name="fontWidth">폰트 넓이</param>
        /// <param name="fontHeight">폰트 높이</param>
        /// <param name="selectedFont">폰트 Style Selected
        /// [   A: TEC FONT1 (Helvetica [bold])   ]
        /// [   B: TEC FONT1 (Helvetica [bold] proportional)   ]
        /// [   E: Price Font 1   ]
        /// [   F: Price Font 2   ]
        /// [   G: Price Font 3   ]
        /// [   H: DUTCH801 Bold (Times Roman Proportional)   ]
        /// [   I: BRUSH738 Regular (Pop Proportional)   ]
        /// [   J: GOTHIC725 Black (Proportional)   ]</param>
        /// <param name="rotate">font 회전 각도(0,90,180,270)</param>
        /// <param name="textOption">문자 특성 "B" = 기본 Black Font</param>
        /// <returns></returns>
        public string _SetTrueFont(double textNumber, double positionX, double positionY, double fontWidth, double fontHeight, string selectedFont, int rotate, string textOption = "B")
        {
            string rotateValue = string.Empty;
            switch (rotate)
            {
                case 0:
                    rotateValue = "00";
                    break;
                case 90:
                    rotateValue = "11";
                    break;
                case 180:
                    rotateValue = "22";
                    break;
                case 270:
                    rotateValue = "33";
                    break;
            }


            StringBuilder builder = new StringBuilder();
            builder.Append(_StartCommand)
                   .Append("PV")
                   .Append(textNumber.ToString("00")) // a
                   .Append(";")
                   .Append(positionX.ToString("0000")) // b
                   .Append(",")
                   .Append(positionY.ToString("0000")) // c
                   .Append(",")
                   .Append(fontWidth.ToString("0000")) // d
                   .Append(",")
                   .Append(fontHeight.ToString("0000")) //e
                   .Append(",")
                   .Append(selectedFont) // f ( 여기에서 폰트 지정01~20 = 기본폰트, 21~25 추가폰트 ) + g 기능을 하나 추가해야함.(g 는 폰트가 지정된 경로)
                   .Append(",")
                   .Append(rotateValue) // ii
                   .Append(",")
                   .Append(textOption)  // j
                   .Append(_EndCommand);

            return builder.ToString();


        }

        //PV2
        /// <summary>
        /// font형식을 지정합니다
        /// </summary>
        /// <param name="textNumber">textGroupNumber (00~99) </param>
        /// <param name="positionX">StartPoint X Value</param>
        /// <param name="positionY">StartPoint Y Value</param>
        /// <param name="fontWidth">폰트 넓이</param>
        /// <param name="fontHeight">폰트 높이</param>
        /// <param name="selectedFont">21~25</param>
        /// <param name="rotate">font 회전 각도(0,90,180,270)</param>
        /// <param name="fontPath">폰트 지정 경로 1= 1슬롯, 2 = 2슬롯</param>
        /// <param name="textOption">문자 특성 "B" = 기본 Black Font</param>
        /// <returns></returns>
        public string _SetTrueFont(double textNumber, double positionX, double positionY, double fontWidth, double fontHeight, string selectedFont, int rotate, int fontPath,string textOption = "B")
        {
            string rotateValue = string.Empty;
            switch (rotate)
            {
                case 0:
                    rotateValue = "00";
                    break;
                case 90:
                    rotateValue = "11";
                    break;
                case 180:
                    rotateValue = "22";
                    break;
                case 270:
                    rotateValue = "33";
                    break;
            }


            StringBuilder builder = new StringBuilder();
            builder.Append(_StartCommand)
                   .Append("PV")
                   .Append(textNumber.ToString("00")) // a
                   .Append(";")
                   .Append(positionX.ToString("0000")) // b
                   .Append(",")
                   .Append(positionY.ToString("0000")) // c
                   .Append(",")
                   .Append(fontWidth.ToString("0000")) // d
                   .Append(",")
                   .Append(fontHeight.ToString("0000")) //e
                   .Append(",")
                   .Append(selectedFont) // f ( 여기에서 폰트 지정01~20 = 기본폰트, 21~25 추가폰트 ) + g 기능을 하나 추가해야함.(g 는 폰트가 지정된 경로)
                   .Append(",")
                   .Append(fontPath.ToString()) // g 폰트지정 경로
                   .Append(",")
                   .Append(rotateValue) // ii
                   .Append(",")
                   .Append(textOption)  // j
                   .Append(_EndCommand);

            return builder.ToString();


        }

        //XB
        /// <summary>
        /// 바코드 형식을 지정합니다
        /// </summary>
        /// <param name="barcodeNumber">지정 Barcode Number [ 0 ~ 31 ]</param>
        /// <param name="positionX"> X축 [ 0000 - 9999 ]</param>
        /// <param name="positionY"> Y축 [ 0000 - 9999 ]</param>
        /// <param name="barcodeType">BarcodeType
        ///     [  0: JAN8, EAN8  ]
        ///     [  5: JAN13, EAN13  ]
        ///     [  6: UPC-E  ]
        ///     [  7: EAN13 + 2 digits  ]
        ///     [  8: EAN13 + 5 digits  ]
        ///     [  9: CODE128 (with auto code selection)   ]
        ///     [  A: CODE128 (without auto code selection)  ]
        ///     [  C: CODE93  ]
        ///     [  G: UPC-E + 2 digits  ]
        ///     [  H: UPC-E + 5 digits  ]
        ///     [  I: EAN8 + 2 digits  ]
        ///     [  J: EAN8 + 5 digits  ]
        ///     [  K: UPC-A  ]
        ///     [  L: UPC-A + 2 digits  ]
        ///     [  M: UPC-A + 5 digits  ]
        ///     [  N: UCC/EAN128  ]
        ///     [  R: Customer bar code (Postal code for Japan)  ]
        ///     [  S: Highest priority customer bar code (Postal code for Japan)  ]
        ///     [  U: POSTNET (Postal code for U.S)  ]
        ///     [  V: RM4SCC (ROYAL MAIL 4 STATE CUSTOMER CODE) / (Postal code for U.K)  ]
        ///     [  W: KIX CODE (Postal code for Belgium)  ]
        /// </param>
        /// <param name="checkDigit">판독 오류 체크 default : 1</param>
        /// <param name="moduleWidth">01~15(in dots)</param>
        /// <param name="rotate">회전 각도 ( 0,90,180,270 )</param>
        /// <param name="barcodeHeigh">Barcode 높이( 0000 ~ 1000 -> {0.1mm units} )</param>
        /// <returns></returns>
        public string _SetBarcode(double barcodeNumber, double positionX, double positionY, string barcodeType, int checkDigit, double moduleWidth, int rotate, double barcodeHeigh)
        {
            string rotateValue = string.Empty;
            switch (rotate)
            {
                case 0:
                    rotateValue = "0";
                    break;
                case 90:
                    rotateValue = "1";
                    break;
                case 180:
                    rotateValue = "2";
                    break;
                case 270:
                    rotateValue = "3";
                    break;
            }


            StringBuilder builder = new StringBuilder();
            builder.Append(_StartCommand)
                   .Append("XB")
                   .Append(barcodeNumber.ToString("00")) // aa
                   .Append(";")
                   .Append(positionX.ToString("0000")) // bbbb
                   .Append(",")
                   .Append(positionY.ToString("0000")) // cccc
                   .Append(",")
                   .Append(barcodeType) // d
                   .Append(",")
                   .Append(checkDigit.ToString()) // e
                   .Append(",")
                   .Append(moduleWidth.ToString("00")) // ff
                   .Append(",")
                   .Append(rotateValue) // k
                   .Append(",")
                   .Append(barcodeHeigh.ToString("0000")) // llll
                   .Append(_EndCommand);

            return builder.ToString();


        }

        //RC
        /// <summary>
        /// _SetBitmapFont 의 데이터 주입
        /// </summary>
        /// <param name="textNumber">_SetSelectedLine와 동일한 textNumber</param>
        /// <param name="inputText">_SetSelectedLine에 들어갈 bitmapFont inputValue</param>
        /// <returns></returns>
        public string _SetBitmapValueInput(double textNumber, string inputText)
        {
            StringBuilder builder = new StringBuilder();
            builder.Append(_StartCommand)
                   .Append("RC")
                   .Append(textNumber.ToString("000"))
                   .Append(";")
                   .Append(inputText)
                   .Append(_EndCommand);

            return builder.ToString();
        }
        //RV
        /// <summary>
        /// _SetTrueFont 의 데이터 주입
        /// </summary>
        /// <param name="textNumber">_SetTrueFont와 동일한 textNumber</param>
        /// <param name="inputText">해당 textNumber의 위치에 들어갈 text DataValue</param>
        /// <returns></returns>
        public string _SetTrueValueInput(double textNumber, string inputText)
        {
            StringBuilder builder = new StringBuilder();
            builder.Append(_StartCommand)
                   .Append("RV")
                   .Append(textNumber.ToString("00"))
                   .Append(";")
                   .Append(inputText)
                   .Append(_EndCommand);

            return builder.ToString();
        }

        //RB
        /// <summary>
        /// _SetBarcode 의 데이터 주입
        /// </summary>
        /// <param name="barcodeNumber">_SetBarcode 동일한 textNumber</param>
        /// <param name="inputText">해당 textNumber의 위치에 들어갈 barcode DataValue</param>
        /// <returns></returns>
        public string _SetBarcodeValueInput(double barcodeNumber, string inputText)
        {
            StringBuilder builder = new StringBuilder();
            builder.Append(_StartCommand)
                   .Append("RB")
                   .Append(barcodeNumber.ToString("00"))
                   .Append(";")
                   .Append(inputText)
                   .Append(_EndCommand);

            return builder.ToString();
        }

        //XS
        /// <summary>
        /// 출력 시작 이벤트 메서드
        /// </summary>
        /// <param name="printCount">출력 할 개수를 입력합니다[0001~9999]</param>
        /// <param name="cutInterval">몇장 단위로 커팅 할 것인가 : 000은 커팅을 하지않습니다[000~100]</param>
        /// <param name="sensorType">라벨 센서 타입 선정[0:센서없음][1:반사][2:투과-일반라벨][3:투과-사전인쇄라벨][4:반사-수동임계값적용]</param>
        /// <param name="printMode">프린터 모드 0:배치모드,1:스트립모드-백센서on,2:스트립모드-백센서off</param>
        /// <param name="printSpeed">1단,2단,3단 [3ips,5ips,8ips] [1~3]</param>
        /// <param name="withRibbonValue">리본 포함 여부 [0:리본 미포함, 1:리본 포함-절약모드, 2:리본 포함-절약off, 3:리본 포함-헤드업]</param>
        /// <param name="tagRotation">0점 설정[0:좌상단, 1:우하단, 2:우상단, 3:좌하단</param>
        /// <param name="statusResponse">상태 응답 [0:off,1:on]</param>
        /// <returns></returns>
        public string _SetStartPrinting(double printCount, double cutInterval, int sensorType, int printMode, int printSpeed, int withRibbonValue, int tagRotation, int statusResponse)
        {
            string printModestr = string.Empty;
            switch (printMode)
            {
                case 0:
                    printModestr = "C";
                    break;
                case 1:
                    printModestr = "D";
                    break;
                case 2:
                    printModestr = "E";
                    break;
            }

            string printSpeedConvert = string.Empty;
            switch (printSpeed)
            {
                case 1:
                    printSpeedConvert = "3";
                    break;
                case 2:
                    printSpeedConvert = "4";
                    break;
                case 3:
                    printSpeedConvert = "8";
                    break;
            }

            if (sensorType > 4)
            {
                sensorType = 0;
            }

            StringBuilder builder = new StringBuilder();
            builder.Append(_StartCommand)
                   .Append("XS")
                   .Append(";")
                   .Append("I")
                   .Append(",")
                   .Append(printCount.ToString("0000")) // aaaa
                   .Append(",")
                   .Append(cutInterval.ToString("000")) // bbb
                   .Append(sensorType.ToString()) // c
                   .Append(printModestr) // d
                   .Append(printSpeedConvert) // e
                   .Append(withRibbonValue.ToString()) // f
                   .Append(tagRotation) // g
                   .Append(statusResponse) // h
                   .Append(_EndCommand); // end

            return builder.ToString();
        }
    }
}
