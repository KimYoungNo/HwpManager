from enum import Enum

def InsertText(hwp, *texts, /, sep='\n'):
    for text in texts:
        with hwp.HParameterSet("HInsertText", "InsertText") as HText:
            HText.Text = str(text)+sep


def PrintPDF(hwp, outputpath, printername, printrange=None):
    if printrange is None:
        printrange = f"1-{hwp.PageCount}"
    
    with hwp.HParameterSet("HPrint", "PrintToPDF") as HPrint:
        HPrint.filename = outputpath
        HPrint.PrinterName = printername	
        HPrint.PrintMethod = hwp.PrintType("Nomal")
        HPrint.Device = hwp.PrintDevice("PDF")
        HPrint.Range = hwp.PrintRange("Custom")
        HPrint.RangeCustom = printrange


class ParagraphAlignFlag(int, Enum):
    JUSTIFY = 0
    LEFT = 1
    RIGHT = 2
    CENTER = 3
    DIST = 4
    DISTEX = 5

def SetFontStyle(hwp, fontname, fontsize, align="JUSTIFY", linespace=150, bold=False, italic=False, offset=0):
    with hwp.HParameterSet("HCharShape", "CharShape") as HFont:
        HFont.Height = hwp.PointToHwpUnit(fontsize)
        HFont.Bold = bold
        HFont.Italic = italic
	
        HFont.FaceNameUser = fontname
        HFont.FaceNameSymbol = fontname
        HFont.FaceNameOther = fontname
        HFont.FaceNameJapanese = fontname
        HFont.FaceNameHanja = fontname
        HFont.FaceNameLatin = fontname
        HFont.FaceNameHangul = fontname
		
        HFont.FontTypeUser = 1
        HFont.FontTypeSymbol = 1
        HFont.FontTypeOther = 1
        HFont.FontTypeJapanese = 1
        HFont.FontTypeHanja = 1
        HFont.FontTypeLatin = 1
        HFont.FontTypeHangul = 1
		
        HFont.OffsetUser = offset
        HFont.OffsetSymbol = offset
        HFont.OffsetOther = offset
        HFont.OffsetJapanese = offset
        HFont.OffsetHanja = offset
        HFont.OffsetLatin = offset
        HFont.OffsetHangul = offset
        
    with hwp.HParameterSet("HParaShape", "ParagraphShape") as HPara:
        HPara.LineSpacing = linespace
        HPara.AlignType = getattr(ParagraphAlignFlag, align).value()
        
        
def SetPageBorder(hwp, linewidth, linetype):
    bordertype = hwp.HwpLineType(str(linetype))
    borderwidth = hwp.HwpLineWidth(f"{linewidth}mm")
    
    with hwp.HParameterSet("HSecDef", "PageBorder") as HSecDef:
        HSecDef.HideBorder = 0
        HSecDef.PageBorderFillBoth.TextBorder = 1
        
        HSecDef.PageBorderFillBoth.BorderTypeLeft = bordertype
        HSecDef.PageBorderFillBoth.BorderTypeRight = bordertype
        HSecDef.PageBorderFillBoth.BorderTypeTop = bordertype
        HSecDef.PageBorderFillBoth.BorderTypeBottom = bordertype
        
        HSecDef.PageBorderFillBoth.BorderWidthLeft = borderwidth
        HSecDef.PageBorderFillBoth.BorderWidthRight = borderwidth
        HSecDef.PageBorderFillBoth.BorderWidthTop = borderwidth
        HSecDef.PageBorderFillBoth.BorderWidthBottom = borderwidth
        
        HSecDef.HSet.SetItem("ApplyToPageBorderFill", 3)
        

def SetPageDimension(hwp, leftmargin, rightmargin, topmargin, bottommargin, headermargin, footermargin, paperwidth=210.0, paperheight=297.0):
    with hwp.HParameterSet("HSecDef", "PageBorder") as HSecDef:
        HSecDef.PageDef.LeftMargin = hwp.MiliToHwpUnit(leftmargin)
        HSecDef.PageDef.RightMargin = hwp.MiliToHwpUnit(rightmargin)
        HSecDef.PageDef.TopMargin = hwp.MiliToHwpUnit(topmargin)
        HSecDef.PageDef.BottomMargin = hwp.MiliToHwpUnit(bottommargin)
        
        HSecDef.PageDef.HeaderLen = hwp.MiliToHwpUnit(headermargin)
        HSecDef.PageDef.FooterLen = hwp.MiliToHwpUnit(footermargin)
        
        HSecDef.PageDef.PaperWidth = hwp.MiliToHwpUnit(paperwidth)
        HSecDef.PageDef.PaperHeight = hwp.MiliToHwpUnit(paperheight)
        
        HSecDef.HSet.SetItem("ApplyClass", 24)
        HSecDef.HSet.SetItem("ApplyTo", 3)
        
def TableInitialCell(hwp):
    hwp.Run.TableCellBlock()
    hwp.Run.TableColBegin()
    hwp.Run.TableColPageUp()
    hwp.Run.Cancel()
