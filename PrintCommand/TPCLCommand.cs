﻿using System;
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
        public string _StartCommand = "{";
        /// <summary>
        /// [LF][NUL] = "|}"
        /// </summary>
        public string _EndCommand = "|}";
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

        /// <summary>
        /// 열전도 방법과, 인쇄 밀도를 지정합니다.
        /// density가 true면, Darker, false면 Light
        /// densityValue = 0~10 Lv, thermal가 true면 열전도, false면 직접 열전도를 시행합니다
        /// </summary>
        /// <param name="density"></param>
        /// <param name="densityValue"></param>
        /// <param name="thermal"></param>
        /// <returns></returns>
        public string _SetPrintDensity(bool density, double densityValue, bool thermal)
        {
            char densityChoice = density ? '+' : '-';
            int thermalChoice = thermal ? 0 : 1;
            if (densityValue > 10)
            {
                densityValue = 10;
            }

            StringBuilder builder = new StringBuilder();
            builder.Append(_StartCommand)
                   .Append("AY;")
                   .Append(densityChoice)
                   .Append(densityValue.ToString("00"))
                   .Append(",")
                   .Append(thermalChoice.ToString("00"))
                   .Append(_EndCommand);

            return builder.ToString();
        }

        /// <summary>
        /// 리본 모터 Volt출력값을 조정합니다
        /// input = 1 당 5% 감소 최대 15 (45%)
        /// </summary>
        /// <param name="ribbonVolt"></param>
        /// <param name="ribbonBackVolt"></param>
        /// <returns></returns>
        public string _SetMotorVoltCommand(double ribbonVolt,double ribbonBackVolt)
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


        /// <summary>
        /// 비트맵 글꼴 형식을 지정합니다
        /// </summary>
        /// <param name="textInputNum">textGroup을 지정 합니다 (0~199)</param>
        /// <param name="positionX">X좌표를 지정 합니다(0~9999) 1=0.1mm</param>
        /// <param name="positionY">Y좌표를 지정 합니다(0~9999) 1=0.1mm</param>
        /// <param name="textHorizontalMargin">가로 간격을 지정 합니다(1~9 or 05~95 = 0.5~9.5)</param>
        /// <param name="textVerticalMargin">새로 간격을 지정 합니다(1~9 or 05~95 = 0.5~9.5)</param>
        /// <param name="TypeFont"> FontStyle을 지정합니다
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
        /// <param name="textOption">잉크농도를 지정합니다 B = 기본 Black character</param>
        /// <returns></returns>
        public string _SetBitmapText(double textInputNum, double positionX, double positionY, double textHorizontalMargin, double textVerticalMargin, string TypeFont, int rotate, string textOption = "B")
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
                   .Append(textInputNum.ToString("000"))
                   .Append(";")
                   .Append(positionX.ToString("0000"))
                   .Append(",")
                   .Append(positionY.ToString("0000"))
                   .Append(",")
                   .Append(textHorizontalMargin.ToString())
                   .Append(",")
                   .Append(textVerticalMargin.ToString())
                   .Append(",")
                   .Append(TypeFont)
                   .Append(",")
                   .Append(rotateValue)
                   .Append(",")
                   .Append(textOption)
                   .Append(_EndCommand);

            return builder.ToString();
        }

        

    }
}