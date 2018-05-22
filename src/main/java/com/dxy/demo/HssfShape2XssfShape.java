package com.dxy.demo;

import com.microsoft.schemas.office.visio.x2012.main.ShapesType;
import org.apache.poi.hssf.usermodel.HSSFShapeTypes;
import org.apache.poi.ss.usermodel.ShapeTypes;

/**
 * Created by daixiyang on 2018/5/22
 */
public class HssfShape2XssfShape {

    public static int adapt(int hssfShapeType){
        switch (hssfShapeType){
            case HSSFShapeTypes.Rectangle :
                return ShapeTypes.RECT;
            case HSSFShapeTypes.RoundRectangle :
                return ShapeTypes.ROUND_RECT;
            case HSSFShapeTypes.Ellipse :
                return ShapeTypes.ELLIPSE;
            case HSSFShapeTypes.Diamond :
                return ShapeTypes.DIAMOND;
            case HSSFShapeTypes.IsocelesTriangle :
                return ShapeTypes.TRIANGLE;
            case HSSFShapeTypes.RightTriangle :
                return ShapeTypes.RT_TRIANGLE;
            case HSSFShapeTypes.Parallelogram :
                return ShapeTypes.PARALLELOGRAM;
            case HSSFShapeTypes.Trapezoid :
                return ShapeTypes.TRAPEZOID;
            case HSSFShapeTypes.Hexagon :
                return ShapeTypes.HEXAGON;
            case HSSFShapeTypes.Octagon :
                return ShapeTypes.OCTAGON;
            case HSSFShapeTypes.Plus :
                return ShapeTypes.PLUS;
            case HSSFShapeTypes.Star :
                return ShapeTypes.STAR_4; //todo
            case HSSFShapeTypes.Arrow :
                return ShapeTypes.RIGHT_ARROW; //todo
            case HSSFShapeTypes.ThickArrow :
                return ShapeTypes.RIGHT_ARROW; //todo
            case HSSFShapeTypes.HomePlate :
                return ShapeTypes.HOME_PLATE;
            case HSSFShapeTypes.Cube :
                return ShapeTypes.CUBE;
            case HSSFShapeTypes.Balloon :
                return ShapeTypes.RECT; //todo
            case HSSFShapeTypes.Seal :
                return ShapeTypes.IRREGULAR_SEAL_1; //todo
            case HSSFShapeTypes.Arc :
                return ShapeTypes.ARC;
            case HSSFShapeTypes.Line :
                return ShapeTypes.LINE;
            case HSSFShapeTypes.Plaque :
                return ShapeTypes.PLAQUE;
            case HSSFShapeTypes.Can :
                return ShapeTypes.CAN;
            case HSSFShapeTypes.Donut :
                return ShapeTypes.DONUT;
            case HSSFShapeTypes.TextSimple :
            case HSSFShapeTypes.TextOctagon :
            case HSSFShapeTypes.TextHexagon :
            case HSSFShapeTypes.TextCurve :
            case HSSFShapeTypes.TextWave :
            case HSSFShapeTypes.TextRing :
            case HSSFShapeTypes.TextOnCurve :
            case HSSFShapeTypes.TextOnRing :
                return ShapeTypes.RECT; //todo
            case HSSFShapeTypes.StraightConnector1 :
                return ShapeTypes.STRAIGHT_CONNECTOR_1;
            case HSSFShapeTypes.BentConnector2 :
                return ShapeTypes.BENT_CONNECTOR_2;
            case HSSFShapeTypes.BentConnector3 :
                return ShapeTypes.BENT_CONNECTOR_3;
            case HSSFShapeTypes.BentConnector4 :
                return ShapeTypes.BENT_CONNECTOR_4;
            case HSSFShapeTypes.BentConnector5 :
                return ShapeTypes.BENT_CONNECTOR_5;
            case HSSFShapeTypes.CurvedConnector2 :
                return ShapeTypes.CURVED_CONNECTOR_2;
            case HSSFShapeTypes.CurvedConnector3 :
                return ShapeTypes.CURVED_CONNECTOR_3;
            case HSSFShapeTypes.CurvedConnector4 :
                return ShapeTypes.CURVED_CONNECTOR_4;
            case HSSFShapeTypes.CurvedConnector5 :
                return ShapeTypes.CURVED_CONNECTOR_5;
            case HSSFShapeTypes.Callout1 :
                return ShapeTypes.CALLOUT_1;
            case HSSFShapeTypes.Callout2 :
                return ShapeTypes.CALLOUT_2;
            case HSSFShapeTypes.Callout3 :
                return ShapeTypes.CALLOUT_3;
            case HSSFShapeTypes.AccentCallout1 :
                return ShapeTypes.ACCENT_CALLOUT_1;
            case HSSFShapeTypes.AccentCallout2 :
                return ShapeTypes.ACCENT_CALLOUT_2;
            case HSSFShapeTypes.AccentCallout3 :
                return ShapeTypes.ACCENT_CALLOUT_3;
            case HSSFShapeTypes.BorderCallout1 :
                return ShapeTypes.BORDER_CALLOUT_1;
            case HSSFShapeTypes.BorderCallout2 :
                return ShapeTypes.BORDER_CALLOUT_2;
            case HSSFShapeTypes.BorderCallout3 :
                return ShapeTypes.BORDER_CALLOUT_3;
            case HSSFShapeTypes.AccentBorderCallout1 :
                return ShapeTypes.ACCENT_BORDER_CALLOUT_1;
            case HSSFShapeTypes.AccentBorderCallout2 :
                return ShapeTypes.ACCENT_BORDER_CALLOUT_2;
            case HSSFShapeTypes.AccentBorderCallout3 :
                return ShapeTypes.ACCENT_BORDER_CALLOUT_3;
            case HSSFShapeTypes.Ribbon :
                return ShapeTypes.RIBBON;
            case HSSFShapeTypes.Ribbon2 :
                return ShapeTypes.RIBBON_2;
            case HSSFShapeTypes.Chevron :
                return ShapeTypes.CHEVRON;
            case HSSFShapeTypes.Pentagon :
                return ShapeTypes.PENTAGON;
            case HSSFShapeTypes.NoSmoking :
                return ShapeTypes.NO_SMOKING;
            case HSSFShapeTypes.Star8 :
                return ShapeTypes.STAR_8;
            case HSSFShapeTypes.Star16 :
                return ShapeTypes.STAR_16;
            case HSSFShapeTypes.Star32 :
                return ShapeTypes.STAR_32;
            case HSSFShapeTypes.WedgeRectCallout :
                return ShapeTypes.WEDGE_RECT_CALLOUT;
            case HSSFShapeTypes.WedgeRRectCallout :
                return ShapeTypes.WEDGE_ROUND_RECT_CALLOUT;
            case HSSFShapeTypes.WedgeEllipseCallout :
                return ShapeTypes.WEDGE_ELLIPSE_CALLOUT;
            case HSSFShapeTypes.Wave :
                return ShapeTypes.WAVE;
            case HSSFShapeTypes.FoldedCorner :
                return ShapeTypes.FOLDED_CORNER;
            case HSSFShapeTypes.LeftArrow :
                return ShapeTypes.LEFT_ARROW;
            case HSSFShapeTypes.DownArrow :
                return ShapeTypes.DOWN_ARROW;
            case HSSFShapeTypes.UpArrow :
                return ShapeTypes.UP_ARROW;
            case HSSFShapeTypes.LeftRightArrow :
                return ShapeTypes.LEFT_RIGHT_ARROW;
            case HSSFShapeTypes.UpDownArrow :
                return ShapeTypes.UP_DOWN_ARROW;
            case HSSFShapeTypes.IrregularSeal1 :
                return ShapeTypes.IRREGULAR_SEAL_1;
            case HSSFShapeTypes.IrregularSeal2 :
                return ShapeTypes.IRREGULAR_SEAL_2;
            case HSSFShapeTypes.LightningBolt :
                return ShapeTypes.LIGHTNING_BOLT;
            case HSSFShapeTypes.Heart :
                return ShapeTypes.HEART;
            case HSSFShapeTypes.PictureFrame :
                return ShapeTypes.RECT; //todo
            case HSSFShapeTypes.QuadArrow :
                return ShapeTypes.QUAD_ARROW;
            case HSSFShapeTypes.LeftArrowCallout :
                return ShapeTypes.LEFT_ARROW_CALLOUT;
            case HSSFShapeTypes.RightArrowCallout :
                return ShapeTypes.RIGHT_ARROW_CALLOUT;
            case HSSFShapeTypes.UpArrowCallout :
                return ShapeTypes.UP_ARROW_CALLOUT;
            case HSSFShapeTypes.DownArrowCallout :
                return ShapeTypes.DOWN_ARROW_CALLOUT;
            case HSSFShapeTypes.LeftRightArrowCallout :
                return ShapeTypes.LEFT_RIGHT_ARROW_CALLOUT;
            case HSSFShapeTypes.UpDownArrowCallout :
                return ShapeTypes.UP_DOWN_ARROW_CALLOUT;
            case HSSFShapeTypes.QuadArrowCallout :
                return ShapeTypes.QUAD_ARROW_CALLOUT;
            case HSSFShapeTypes.Bevel :
                return ShapeTypes.BEVEL;
            case HSSFShapeTypes.LeftBracket :
                return ShapeTypes.LEFT_BRACKET;
            case HSSFShapeTypes.RightBracket :
                return ShapeTypes.RIGHT_BRACKET;
            case HSSFShapeTypes.LeftBrace :
                return ShapeTypes.LEFT_BRACE;
            case HSSFShapeTypes.RightBrace :
                return ShapeTypes.RIGHT_BRACE;
            case HSSFShapeTypes.LeftUpArrow :
                return ShapeTypes.LEFT_UP_ARROW;
            case HSSFShapeTypes.BentUpArrow :
                return ShapeTypes.BENT_UP_ARROW;
            case HSSFShapeTypes.BentArrow :
                return ShapeTypes.BENT_ARROW;
            case HSSFShapeTypes.Star24 :
                return ShapeTypes.STAR_24;
            case HSSFShapeTypes.StripedRightArrow :
                return ShapeTypes.STRIPED_RIGHT_ARROW;
            case HSSFShapeTypes.NotchedRightArrow :
                return ShapeTypes.NOTCHED_RIGHT_ARROW;
            case HSSFShapeTypes.BlockArc :
                return ShapeTypes.BLOCK_ARC;
            case HSSFShapeTypes.SmileyFace :
                return ShapeTypes.SMILEY_FACE;
            case HSSFShapeTypes.VerticalScroll :
                return ShapeTypes.VERTICAL_SCROLL;
            case HSSFShapeTypes.HorizontalScroll :
                return ShapeTypes.HORIZONTAL_SCROLL;
            case HSSFShapeTypes.CircularArrow :
                return ShapeTypes.CIRCULAR_ARROW;
            case HSSFShapeTypes.NotchedCircularArrow :
                return ShapeTypes.CIRCULAR_ARROW; //todo
            case HSSFShapeTypes.UturnArrow :
                return ShapeTypes.UTURN_ARROW;
            case HSSFShapeTypes.CurvedRightArrow :
                return ShapeTypes.CURVED_RIGHT_ARROW;
            case HSSFShapeTypes.CurvedLeftArrow :
                return ShapeTypes.CURVED_LEFT_ARROW;
            case HSSFShapeTypes.CurvedUpArrow :
                return ShapeTypes.CURVED_UP_ARROW;
            case HSSFShapeTypes.CurvedDownArrow :
                return ShapeTypes.CURVED_DOWN_ARROW;
            case HSSFShapeTypes.CloudCallout :
                return ShapeTypes.CLOUD_CALLOUT;
            case HSSFShapeTypes.EllipseRibbon :
                return ShapeTypes.ELLIPSE_RIBBON;
            case HSSFShapeTypes.EllipseRibbon2 :
                return ShapeTypes.ELLIPSE_RIBBON_2;
            case HSSFShapeTypes.FlowChartProcess :
                return ShapeTypes.FLOW_CHART_PROCESS;
            case HSSFShapeTypes.FlowChartDecision :
                return ShapeTypes.FLOW_CHART_DECISION;
            case HSSFShapeTypes.FlowChartInputOutput :
                return ShapeTypes.FLOW_CHART_INPUT_OUTPUT;
            case HSSFShapeTypes.FlowChartPredefinedProcess :
                return ShapeTypes.FLOW_CHART_PREDEFINED_PROCESS;
            case HSSFShapeTypes.FlowChartInternalStorage :
                return ShapeTypes.FLOW_CHART_INTERNAL_STORAGE;
            case HSSFShapeTypes.FlowChartDocument :
                return ShapeTypes.FLOW_CHART_DOCUMENT;
            case HSSFShapeTypes.FlowChartMultidocument :
                return ShapeTypes.FLOW_CHART_MULTIDOCUMENT;
            case HSSFShapeTypes.FlowChartTerminator :
                return ShapeTypes.FLOW_CHART_TERMINATOR;
            case HSSFShapeTypes.FlowChartPreparation :
                return ShapeTypes.FLOW_CHART_PREPARATION;
            case HSSFShapeTypes.FlowChartManualInput :
                return ShapeTypes.FLOW_CHART_MANUAL_INPUT;
            case HSSFShapeTypes.FlowChartManualOperation :
                return ShapeTypes.FLOW_CHART_MANUAL_OPERATION;
            case HSSFShapeTypes.FlowChartConnector :
                return ShapeTypes.FLOW_CHART_CONNECTOR;
            case HSSFShapeTypes.FlowChartPunchedCard :
                return ShapeTypes.FLOW_CHART_PUNCHED_CARD;
            case HSSFShapeTypes.FlowChartPunchedTape :
                return ShapeTypes.FLOW_CHART_PUNCHED_TAPE;
            case HSSFShapeTypes.FlowChartSummingJunction :
                return ShapeTypes.FLOW_CHART_SUMMING_JUNCTION;
            case HSSFShapeTypes.FlowChartOr :
                return ShapeTypes.FLOW_CHART_OR;
            case HSSFShapeTypes.FlowChartCollate :
                return ShapeTypes.FLOW_CHART_COLLATE;
            case HSSFShapeTypes.FlowChartSort :
                return ShapeTypes.FLOW_CHART_SORT;
            case HSSFShapeTypes.FlowChartExtract :
                return ShapeTypes.FLOW_CHART_EXTRACT;
            case HSSFShapeTypes.FlowChartMerge :
                return ShapeTypes.FLOW_CHART_MERGE;
            case HSSFShapeTypes.FlowChartOfflineStorage :
                return ShapeTypes.FLOW_CHART_OFFLINE_STORAGE;
            case HSSFShapeTypes.FlowChartOnlineStorage :
                return ShapeTypes.FLOW_CHART_ONLINE_STORAGE;
            case HSSFShapeTypes.FlowChartMagneticTape :
                return ShapeTypes.FLOW_CHART_MAGNETIC_TAPE;
            case HSSFShapeTypes.FlowChartMagneticDisk :
                return ShapeTypes.FLOW_CHART_MAGNETIC_DISK;
            case HSSFShapeTypes.FlowChartMagneticDrum :
                return ShapeTypes.FLOW_CHART_MAGNETIC_DRUM;
            case HSSFShapeTypes.FlowChartDisplay :
                return ShapeTypes.FLOW_CHART_DISPLAY;
            case HSSFShapeTypes.FlowChartDelay :
                return ShapeTypes.FLOW_CHART_DELAY;
            case HSSFShapeTypes.TextPlainText :
            case HSSFShapeTypes.TextStop :
            case HSSFShapeTypes.TextTriangle :
            case HSSFShapeTypes.TextTriangleInverted :
            case HSSFShapeTypes.TextChevron :
            case HSSFShapeTypes.TextChevronInverted :
            case HSSFShapeTypes.TextRingInside :
            case HSSFShapeTypes.TextRingOutside :
            case HSSFShapeTypes.TextArchUpCurve :
            case HSSFShapeTypes.TextArchDownCurve :
            case HSSFShapeTypes.TextCircleCurve :
            case HSSFShapeTypes.TextButtonCurve :
            case HSSFShapeTypes.TextArchUpPour :
            case HSSFShapeTypes.TextArchDownPour :
            case HSSFShapeTypes.TextCirclePour :
            case HSSFShapeTypes.TextButtonPour :
            case HSSFShapeTypes.TextCurveUp :
            case HSSFShapeTypes.TextCurveDown :
            case HSSFShapeTypes.TextCascadeUp :
            case HSSFShapeTypes.TextCascadeDown :
            case HSSFShapeTypes.TextWave1 :
            case HSSFShapeTypes.TextWave2 :
            case HSSFShapeTypes.TextWave3 :
            case HSSFShapeTypes.TextWave4 :
            case HSSFShapeTypes.TextInflate :
            case HSSFShapeTypes.TextDeflate :
            case HSSFShapeTypes.TextInflateBottom :
            case HSSFShapeTypes.TextDeflateBottom :
            case HSSFShapeTypes.TextInflateTop :
            case HSSFShapeTypes.TextDeflateTop :
            case HSSFShapeTypes.TextDeflateInflate :
            case HSSFShapeTypes.TextDeflateInflateDeflate :
            case HSSFShapeTypes.TextFadeRight :
            case HSSFShapeTypes.TextFadeLeft :
            case HSSFShapeTypes.TextFadeUp :
            case HSSFShapeTypes.TextFadeDown :
            case HSSFShapeTypes.TextSlantUp :
            case HSSFShapeTypes.TextSlantDown :
            case HSSFShapeTypes.TextCanUp :
            case HSSFShapeTypes.TextCanDown :
                return ShapeTypes.RECT; //todo
            case HSSFShapeTypes.FlowChartAlternateProcess :
                return ShapeTypes.FLOW_CHART_ALTERNATE_PROCESS;
            case HSSFShapeTypes.FlowChartOffpageConnector :
                return ShapeTypes.FLOW_CHART_OFFPAGE_CONNECTOR;
            case HSSFShapeTypes.Callout90 :
                return ShapeTypes.CALLOUT_1; //todo
            case HSSFShapeTypes.AccentCallout90 :
                return ShapeTypes.ACCENT_CALLOUT_1; //todo
            case HSSFShapeTypes.BorderCallout90 :
                return ShapeTypes.BORDER_CALLOUT_1; //todo
            case HSSFShapeTypes.AccentBorderCallout90 :
                return ShapeTypes.ACCENT_BORDER_CALLOUT_1; //todo
            case HSSFShapeTypes.LeftRightUpArrow :
                return ShapeTypes.LEFT_RIGHT_UP_ARROW;
            case HSSFShapeTypes.Sun :
                return ShapeTypes.SUN;
            case HSSFShapeTypes.Moon :
                return ShapeTypes.MOON;
            case HSSFShapeTypes.BracketPair :
                return ShapeTypes.BRACKET_PAIR;
            case HSSFShapeTypes.BracePair :
                return ShapeTypes.BRACE_PAIR;
            case HSSFShapeTypes.Star4 :
                return ShapeTypes.STAR_4;
            case HSSFShapeTypes.DoubleWave :
                return ShapeTypes.DOUBLE_WAVE;
            case HSSFShapeTypes.ActionButtonBlank :
                return ShapeTypes.ACTION_BUTTON_BLANK;
            case HSSFShapeTypes.ActionButtonHome :
                return ShapeTypes.ACTION_BUTTON_HOME;
            case HSSFShapeTypes.ActionButtonHelp :
                return ShapeTypes.ACTION_BUTTON_HELP;
            case HSSFShapeTypes.ActionButtonInformation :
                return ShapeTypes.ACTION_BUTTON_INFORMATION;
            case HSSFShapeTypes.ActionButtonForwardNext :
                return ShapeTypes.ACTION_BUTTON_FORWARD_NEXT;
            case HSSFShapeTypes.ActionButtonBackPrevious :
                return ShapeTypes.ACTION_BUTTON_BACK_PREVIOUS;
            case HSSFShapeTypes.ActionButtonEnd :
                return ShapeTypes.ACTION_BUTTON_END;
            case HSSFShapeTypes.ActionButtonBeginning :
                return ShapeTypes.ACTION_BUTTON_BEGINNING;
            case HSSFShapeTypes.ActionButtonReturn :
                return ShapeTypes.ACTION_BUTTON_RETURN;
            case HSSFShapeTypes.ActionButtonDocument :
                return ShapeTypes.ACTION_BUTTON_DOCUMENT;
            case HSSFShapeTypes.ActionButtonSound :
                return ShapeTypes.ACTION_BUTTON_SOUND;
            case HSSFShapeTypes.ActionButtonMovie :
                return ShapeTypes.ACTION_BUTTON_MOVIE;
            case HSSFShapeTypes.HostControl :
                return ShapeTypes.RECT; //todo
            case HSSFShapeTypes.TextBox :
                return ShapeTypes.RECT ;// todo
            default:
                return ShapeTypes.RECT;

        }
    }
}
