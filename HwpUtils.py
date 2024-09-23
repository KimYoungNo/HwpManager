def PrintPDF(hwp, output_path, printer_name, print_range=None):
    if print_range is None:
        print_range = f"1-{hwp.PageCount}"
    
    with hwp.HParameterSet("HPrint", "PrintToPDF") as HPrint:
        HPrint.filename = output_path
        HPrint.PrinterName = printer_name	
        HPrint.PrintMethod = hwp.PrintType("Nomal")
        HPrint.Device = hwp.PrintDevice("PDF")
        HPrint.Range = hwp.PrintRange("Custom")
        HPrint.RangeCustom = print_range
