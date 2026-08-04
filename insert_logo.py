import win32com.client as win32
import os

# Use relative paths so it works when the project folder is moved
report_path = "Bao_Cao_Coaching_GSBH.xlsx"
logo_path = "logo_cj.png"

report_abs = os.path.abspath(report_path)
logo_abs = os.path.abspath(logo_path)

excel = None
wb = None
try:
    if not os.path.exists(logo_abs):
        print(f"Error: Logo file not found at {logo_abs}")
        exit(1)
    if not os.path.exists(report_abs):
        print(f"Error: Excel report not found at {report_abs}")
        exit(1)
        
    print("Opening Excel application to insert logo...")
    excel = win32.Dispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    
    wb = excel.Workbooks.Open(report_abs)
    
    # Process both sheets: Summary and Supervisor_Report
    sheet_names = ["Summary", "Supervisor_Report"]
    for s_name in sheet_names:
        try:
            ws = wb.Sheets(s_name)
            
            # 1. Clear any existing picture shapes to prevent overlays
            for shape in list(ws.Shapes):
                try:
                    shape.Delete()
                except:
                    pass
            
            # 2. Add the picture
            # Shapes.AddPicture(Filename, LinkToFile, SaveWithDocument, Left, Top, Width, Height)
            # Row height of rows 1-3 is around 15-20 each, so Top=10, Left=10, Width=55, Height=40 is good.
            ws.Shapes.AddPicture(logo_abs, False, True, 10, 10, 55, 40)
            
            # 3. Shift titles from column A to B so they don't get covered
            # Current title is at Cell(2,1). We move it to Cell(2,2).
            title_val = ws.Cells(2, 1).Value
            if title_val and ("coaching" in str(title_val).lower() or "cj" in str(title_val).lower()):
                ws.Cells(2, 2).Value = title_val
                ws.Cells(2, 1).Value = ""
                
                # Format Cell(2,2) with title style
                ws.Cells(2, 2).Font.Name = "Segoe UI"
                ws.Cells(2, 2).Font.Size = 16
                ws.Cells(2, 2).Font.Bold = True
                ws.Cells(2, 2).Font.Color = 0x272BC5 # CJ Red
                
            print(f"Logo successfully inserted in '{s_name}'.")
        except Exception as sheet_err:
            print(f"Error processing sheet '{s_name}': {sheet_err}")
            
    wb.Save()
    print("Report saved successfully with logo!")
    
except Exception as e:
    import traceback
    print("Error:")
    traceback.print_exc()
finally:
    if wb is not None:
        try:
            wb.Close(SaveChanges=True)
        except:
            pass
    if excel is not None:
        try:
            excel.Quit()
        except:
            pass
