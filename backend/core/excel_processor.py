import pandas as pd
import os
import io
import re
import json

class ExcelProcessor:
    def __init__(self):
        pass

    def parse_spectro_report(self, file_content: bytes):
        """
        Parses the uploaded Spectro Report (xlsx) and returns structured data.
        """
        try:
            file_obj = io.BytesIO(file_content)
            df = None
            errors = []

            # 1. Try modern XLSX (openpyxl)
            try:
                file_obj.seek(0)
                df = pd.read_excel(file_obj, engine='openpyxl')
            except Exception as e: errors.append(f"openpyxl: {e}")

            # 2. Try old XLS (xlrd)
            if df is None or df.empty:
                try:
                    file_obj.seek(0)
                    df = pd.read_excel(file_obj, engine='xlrd')
                except Exception as e: errors.append(f"xlrd: {e}")

            # 3. Try Binary XLSB (pyxlsb)
            if df is None or df.empty:
                try:
                    file_obj.seek(0)
                    df = pd.read_excel(file_obj, engine='pyxlsb')
                except Exception as e: errors.append(f"pyxlsb: {e}")

            # 4. Try CSV (sometimes machine exports rename .csv to .xlsx)
            if df is None or df.empty:
                try:
                    file_obj.seek(0)
                    # Try comma first
                    df = pd.read_csv(file_obj)
                except Exception as e:
                    try:
                        file_obj.seek(0)
                        # Try semicolon (common in Europe/machines)
                        df = pd.read_csv(file_obj, sep=';')
                    except Exception as e2:
                        errors.append(f"csv: {e2}")

            if df is None or df.empty:
                raise Exception(f"Failed to parse file with any engine. Engines tried: {', '.join(errors)}")

            if df.empty:
                raise Exception("The uploaded file appears to be empty.")

            # 1. Cleaning common columns
            if 'S.No' in df.columns:
                df['S.No'] = range(1, 1 + len(df))
            
            # 2. CE% Calculation
            df = self.calculate_ce(df)
            
            # 3. Extract Heats Safely
            heat_col = next((c for c in df.columns if 'HEAT' in c.upper()), df.columns[1] if len(df.columns) > 1 else None)
            heats = []
            if heat_col:
                heats = df[heat_col].dropna().astype(str).unique().tolist()

            # 4. JSON Sanitization (replace NaN/inf with "")
            df_plain = df.copy()
            df_plain = df_plain.replace([float('inf'), float('-inf')], 0)
            df_plain = df_plain.fillna("")

            return {
                "data": df_plain.to_dict(orient="records"),
                "columns": df_plain.columns.tolist(),
                "heats": heats
            }
        except Exception as e:
            import traceback
            print(f"ERROR: Spectro Analysis Failed: {traceback.format_exc()}")
            raise Exception(f"Excel parsing error: {str(e)}")

    def calculate_ce(self, df: pd.DataFrame):
        """
        Calculates Carbon Equivalent (CE%) = C + (Si/3) + P
        Based on user request to handle these specific elements.
        """
        c_col = next((col for col in df.columns if col.strip().upper() == 'C%'), None)
        si_col = next((col for col in df.columns if col.strip().upper() == 'SI%'), None)
        p_col = next((col for col in df.columns if col.strip().upper() == 'P%'), None)

        if c_col and si_col and p_col:
            try:
                c_val = pd.to_numeric(df[c_col], errors='coerce').fillna(0)
                si_val = pd.to_numeric(df[si_col], errors='coerce').fillna(0)
                p_val = pd.to_numeric(df[p_col], errors='coerce').fillna(0)
                
                ce_val = c_val + (si_val / 3) + p_val
                
                # Rounding
                df[c_col] = c_val.round(2)
                df[si_col] = si_val.round(2)
                ce_val = ce_val.round(2)
                
                # Insert next to Si
                loc_index = df.columns.get_loc(si_col) + 1
                if "CE%" not in df.columns:
                    df.insert(loc_index, "CE%", ce_val)
                else:
                    df["CE%"] = ce_val
            except:
                pass
        return df

    def deduplicate(self, df: pd.DataFrame, group_by_col: str):
        """
        Deduplicates records based on Sample Id or Heat No, keeping max 2 per group.
        """
        # Group by column and take top 2
        df = df.groupby(group_by_col).head(2).reset_index(drop=True)
        
        # Re-generate S.No
        if 'S.No' in df.columns:
            df['S.No'] = range(1, 1 + len(df))
            
        return df

    def parse_mtc_template(self, template_path: str):
        """
        Analyses the MTC template (Final correct.xlsx) to extract default labels and structure.
        """
        if not os.path.exists(template_path):
             return None
             
        try:
            df = pd.read_excel(template_path, header=None)
            
            # Extract defaults (logic mirror from app.py)
            raw_part_line = str(df.iloc[7, 1]) if len(df) > 7 else ""
            
            return {
                "part_details": raw_part_line,
                "customer": "SUNDRAM FASTENERS LTD",
                "reference": "REFERENCE - Ductile iron J434C GRADE 4512"
            }
        except Exception as e:
            print(f"Error parsing MTC template: {e}")
            return None
