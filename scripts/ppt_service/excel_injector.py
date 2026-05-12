import os
import logging

logger = logging.getLogger(__name__)

def inject_excel_visuals_into_ppt(excel_path: str, ppt_path: str):
    """
    Looks for shapes starting with 'excel_chart:' or 'excel_table:' in the PPTX.
    Copies the corresponding object from the Excel file and pastes it in the PPTX
    at the exact same coordinates, then deletes the placeholder.
    """
    try:
        import win32com.client
    except ImportError:
        logger.warning("win32com.client not installed. Skipping Excel injection.")
        return

    excel_path = os.path.abspath(excel_path)
    ppt_path = os.path.abspath(ppt_path)

    if not os.path.exists(excel_path):
        logger.warning(f"Excel file not found at {excel_path}. Skipping Excel injection.")
        return
        
    if not os.path.exists(ppt_path):
        logger.warning(f"PPT file not found at {ppt_path}. Skipping Excel injection.")
        return

    excel = None
    ppt = None
    wb = None
    prs = None

    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        ppt = win32com.client.DispatchEx("PowerPoint.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        # ppt.Visible = True  # Sometimes PPT needs to be visible to paste properly on some systems, but try hidden first
        # ppt.DisplayAlerts = False # PowerPoint doesn't have DisplayAlerts in the same way

        wb = excel.Workbooks.Open(excel_path)
        prs = ppt.Presentations.Open(ppt_path, WithWindow=False)

        shapes_replaced = 0

        # Loop through every slide
        for slide_index in range(1, prs.Slides.Count + 1):
            slide = prs.Slides(slide_index)
            
            # We collect shapes to process to avoid modifying the collection while iterating
            shapes_to_replace = []
            for shape_index in range(1, slide.Shapes.Count + 1):
                shape = slide.Shapes(shape_index)
                try:
                    name = shape.Name
                    if name.startswith("excel_chart:") or name.startswith("excel_table:"):
                        shapes_to_replace.append(name)
                        
                    # Check text frame for specific placeholders
                    if getattr(shape, "HasTextFrame", False):
                        if shape.TextFrame.HasText:
                            text = shape.TextFrame.TextFrame.Text if hasattr(shape.TextFrame, 'TextFrame') else shape.TextFrame.Text
                            text_stripped = text.strip() if text else ""
                            if text_stripped == "{{financial_summary_image}}":
                                # Rename the shape to use our logic
                                shape.Name = "excel_table:Fin_Summary"
                                shapes_to_replace.append(shape.Name)
                            elif text_stripped == "{{earnings_forecast_table}}":
                                shape.Name = "excel_table:Earnings_Forecast"
                                shapes_to_replace.append(shape.Name)
                            elif text_stripped == "{{financials_table}}":
                                shape.Name = "excel_table:Financials_Table"
                                shapes_to_replace.append(shape.Name)
                            elif text_stripped == "{{valuations_table}}":
                                shape.Name = "excel_table:Valuations_Table"
                                shapes_to_replace.append(shape.Name)
                            elif text_stripped == "{{key_risks_table}}":
                                shape.Name = "excel_table:Key_Risks"
                                shapes_to_replace.append(shape.Name)
                            elif text_stripped == "{{peer_comparision}}":
                                shape.Name = "excel_table:Peer_Compare"
                                shapes_to_replace.append(shape.Name)
                            elif text_stripped == "{{financial_model_from_excel_operational_sheet}}":
                                shape.Name = "excel_table:Operational_Data"
                                shapes_to_replace.append(shape.Name)
                            elif text_stripped == "{{governance_table}}":
                                shape.Name = "excel_table:Governance"
                                shapes_to_replace.append(shape.Name)
                            elif text_stripped == "{{timeline}}":
                                shape.Name = "excel_table:Timeline"
                                shapes_to_replace.append(shape.Name)
                            elif text_stripped == "{{financial_model_from_excel}}":
                                shape.Name = "excel_table:Op_Charts"
                                shapes_to_replace.append(shape.Name)
                except Exception:
                    pass

            # Process each found placeholder
            for shape_name in shapes_to_replace:
                try:
                    shape = slide.Shapes(shape_name)
                    
                    # 1. Save exact coordinates of the placeholder
                    top, left = shape.Top, shape.Left
                    width, height = shape.Width, shape.Height
                    
                    # 2. Extract Data from Excel based on the naming convention
                    if shape_name.startswith("excel_chart:"):
                        chart_info = shape_name.replace("excel_chart:", "")
                        if "!" in chart_info:
                            sheet_name, chart_name = chart_info.split("!", 1)
                        else:
                            # Default to first sheet or specific sheet if you have a convention
                            sheet_name = "Data Sheet"
                            chart_name = chart_info
                        
                        try:
                            # Try to select by name or index
                            if chart_name.isdigit():
                                wb.Sheets(sheet_name).ChartObjects(int(chart_name)).Copy()
                            else:
                                wb.Sheets(sheet_name).ChartObjects(chart_name).Copy()
                        except Exception as e:
                            logger.error(f"Failed to copy chart '{chart_name}' from sheet '{sheet_name}': {e}")
                            continue
                        
                    elif shape_name.startswith("excel_table:"):
                        # Example shape name: "excel_table:Data Sheet!A10:F20"
                        range_info = shape_name.replace("excel_table:", "")
                        if "!" in range_info:
                            sheet_name, cell_range = range_info.split("!", 1)
                            try:
                                wb.Sheets(sheet_name).Range(cell_range).Copy()
                            except Exception as e:
                                logger.error(f"Failed to copy range '{cell_range}' from sheet '{sheet_name}': {e}")
                                continue
                        else:
                            sheet_name = range_info
                            try:
                                # Copy the entire used range of the sheet
                                wb.Sheets(sheet_name).UsedRange.Copy()
                            except Exception as e:
                                logger.error(f"Failed to copy UsedRange from sheet '{sheet_name}': {e}")
                                continue

                    # 3. Paste into PowerPoint
                    # DataType=2 pastes as an Enhanced Metafile (Scalable Image), which preserves formatting nicely for tables
                    # For charts, it pastes it as a picture. If you want interactive charts, omit DataType.
                    pasted_shape = slide.Shapes.PasteSpecial(DataType=2) 
                    
                    # slide.Shapes.PasteSpecial can return a ShapeRange. Take the first item.
                    if getattr(pasted_shape, "Count", 0) > 0:
                        pasted_shape = pasted_shape(1)
                    
                    # 4. Resize and reposition to match the placeholder perfectly
                    pasted_shape.Top = top
                    pasted_shape.Left = left
                    pasted_shape.Width = width
                    pasted_shape.Height = height
                    
                    # 5. Delete the original placeholder rectangle
                    shape.Delete()
                    shapes_replaced += 1
                    logger.info(f"Successfully replaced placeholder: {shape_name}")
                    
                except Exception as e:
                    logger.error(f"Failed to process placeholder {shape_name}: {e}")

        if shapes_replaced > 0:
            prs.Save()
            logger.info(f"Injected {shapes_replaced} Excel visuals into {ppt_path}")
        else:
            logger.info("No excel_chart: or excel_table: placeholders found.")

    except Exception as e:
        logger.error(f"Error during Excel injection: {e}")
    finally:
        # Always close and quit to avoid hanging background processes
        if wb:
            try: wb.Close(False) 
            except: pass
        if prs:
            try: prs.Close() 
            except: pass
        if excel:
            try: excel.Quit()
            except: pass
        if ppt:
            try: ppt.Quit()
            except: pass
