import os
import win32com.client
import pythoncom
from pathlib import Path
import tempfile
import shutil
import uuid

class DocConverter:
    @staticmethod
    def ensure_docx(file_path):
        """
        Checks if file is .doc, if so converts to .docx in temp dir.
        Returns the path to the .docx file (either original or converted).
        """
        path = Path(file_path)
        if not path.exists():
            raise FileNotFoundError(f"File not found: {file_path}")
            
        if path.suffix.lower() == '.docx':
            return str(path)
            
        if path.suffix.lower() == '.doc':
            return DocConverter.convert_doc_to_docx(str(path))
            
        return str(path)

    @staticmethod
    def convert_doc_to_docx(doc_path):
        """
        Converts .doc to .docx using MS Word or WPS COM automation.
        Saves to a temporary file.
        Uses DispatchEx and safe temporary filenames to avoid path issues.
        Support for WPS Office (Kwps.Application / WPS.Application).
        """
        print(f"Converting .doc to .docx: {doc_path}")
        
        # Initialize COM library
        pythoncom.CoInitialize()
        
        word = None
        doc = None
        safe_source_path = None
        
        try:
            # Create a safe temporary copy of the source file
            temp_dir = tempfile.gettempdir()
            safe_source_name = f"src_{uuid.uuid4().hex[:8]}.doc"
            safe_source_path = os.path.join(temp_dir, safe_source_name)
            
            shutil.copy2(doc_path, safe_source_path)
            
            # Application ProgIDs to try (WPS preferred if installed as it's lighter sometimes, or Word)
            # Actually Word is usually preferred for .doc, but user has WPS.
            app_progids = ["Kwps.Application", "WPS.Application", "Word.Application"]
            
            for progid in app_progids:
                try:
                    word = win32com.client.DispatchEx(progid)
                    # print(f"DEBUG: Successfully dispatched: {progid}")
                    break
                except Exception as e:
                    # print(f"DEBUG: Failed to dispatch {progid}: {e}")
                    try:
                        word = win32com.client.Dispatch(progid)
                        # print(f"DEBUG: Successfully dispatched (standard): {progid}")
                        break
                    except:
                        pass
            
            if not word:
                raise Exception(f"Could not launch Word or WPS. Tried: {', '.join(app_progids)}")
                
            word.Visible = False
            try:
                word.DisplayAlerts = False
            except: pass
            
            # Open the SAFE temp file
            doc = word.Documents.Open(safe_source_path)
            
            # Generate output path
            base_name = os.path.basename(doc_path)
            stem = os.path.splitext(base_name)[0]
            # Sanitize stem for output filename
            safe_stem = "".join([c for c in stem if c.isalnum() or c in (' ', '-', '_')])
            if not safe_stem: safe_stem = "converted_doc"
            
            output_name = f"{safe_stem}_{uuid.uuid4().hex[:8]}_converted.docx"
            output_path = os.path.join(temp_dir, output_name)
            
            # SaveAs format: 
            # Word uses 12 (wdFormatXMLDocument)
            # WPS usually handles 12 too. If fails, might need 16 (wdFormatDocumentDefault) or 0 (wdFormatDocument)
            # But we want .docx.
            doc.SaveAs(output_path, FileFormat=12)
            
            return output_path
            
        except Exception as e:
            print(f"Error converting .doc: {e}")
            raise e
        finally:
            if doc:
                try:
                    doc.Close(False)
                except: pass
            if word:
                try:
                    word.Quit()
                except: pass
            
            # Clean up safe source
            if safe_source_path and os.path.exists(safe_source_path):
                try:
                    os.remove(safe_source_path)
                except: pass
            
            pythoncom.CoUninitialize()
