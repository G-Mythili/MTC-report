from fastapi import FastAPI, UploadFile, File, HTTPException, Body
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
import os
import json
import io
import pandas as pd
from typing import List, Dict, Any

from core.excel_processor import ExcelProcessor
from core.excel_generator import ExcelGenerator

app = FastAPI(title="MTC Report API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Resolve paths relative to the project root (where the .bat is run from)
ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SETTINGS_FILE = os.path.join(ROOT_DIR, "settings.json")
FORMATS_FILE = os.path.join(ROOT_DIR, "report_formats.json")
GRADE_MASTER_FILE = os.path.join(ROOT_DIR, "grade_master.json")
processor = ExcelProcessor()

@app.get("/api/settings")
async def get_settings():
    if os.path.exists(SETTINGS_FILE):
        with open(SETTINGS_FILE, "r") as f:
            return json.load(f)
    return {
        "chem_title": "1. Chemical composition",
        "mech_title": "2. Mechanical Properties",
        "micro_title": "3. Microstructure",
        "matrix_title": "3.1 Matrix",
        "header_align": "center",
        "header_fill_color": "#d9e1f2",
        "border_style": "thin",
        "font_family": "Calibri",
        "font_size": 10
    }

@app.post("/api/settings")
async def update_settings(settings: dict):
    with open(SETTINGS_FILE, "w") as f:
        json.dump(settings, f, indent=4)
    return {"status": "success"}

@app.get("/api/formats")
async def get_formats():
    if os.path.exists(FORMATS_FILE):
        with open(FORMATS_FILE, "r") as f:
            return json.load(f)
    return {}

@app.post("/api/formats")
async def save_format(payload: dict):
    name = payload.get("name")
    cols = payload.get("columns")
    if not name:
        raise HTTPException(status_code=400, detail="Format name is required")
    
    formats = await get_formats()
    formats[name] = cols
    with open(FORMATS_FILE, "w") as f:
        json.dump(formats, f, indent=4)
    return {"status": "success"}

@app.delete("/api/formats/{name}")
async def delete_format(name: str):
    formats = await get_formats()
    if name in formats:
        del formats[name]
        with open(FORMATS_FILE, "w") as f:
            json.dump(formats, f, indent=4)
    # Return success even if not found to handle stale frontend states gracefully
    return {"status": "success"}

@app.put("/api/formats/{name}")
async def rename_format(name: str, payload: dict):
    new_name = payload.get("new_name")
    if not new_name:
        raise HTTPException(status_code=400, detail="New name is required")
    
    formats = await get_formats()
    if name in formats:
        cols = formats.pop(name)
        formats[new_name] = cols
        with open(FORMATS_FILE, "w") as f:
            json.dump(formats, f, indent=4)
        return {"status": "success"}
    raise HTTPException(status_code=404, detail="Format not found")

@app.get("/api/grade-master")
async def get_grade_master():
    if os.path.exists(GRADE_MASTER_FILE):
        with open(GRADE_MASTER_FILE, "r") as f:
            data = json.load(f)
            # Migration: If it's the old format (list per grade), wrap it
            migrated = False
            for grade in data:
                if isinstance(data[grade], list):
                    data[grade] = {
                        "chemistry": data[grade],
                        "mechanical": []
                    }
                    migrated = True
            if migrated:
                with open(GRADE_MASTER_FILE, "w") as fw:
                    json.dump(data, fw, indent=4)
            return data
    return {}

@app.post("/api/grade-master")
async def save_grade_specs(payload: dict):
    grade = payload.get("grade")
    # New format: { "chemistry": [...], "mechanical": [...] }
    specs = payload.get("specs") 
    if not grade:
        raise HTTPException(status_code=400, detail="Grade name is required")
    
    master = await get_grade_master()
    master[grade] = specs
    with open(GRADE_MASTER_FILE, "w") as f:
        json.dump(master, f, indent=4)
    return {"status": "success"}

@app.delete("/api/grade-master/{grade}")
async def delete_grade_master_entry(grade: str):
    master = await get_grade_master()
    if grade in master:
        del master[grade]
        with open(GRADE_MASTER_FILE, "w") as f:
            json.dump(master, f, indent=4)
    return {"status": "success"}

@app.post("/api/analyze")
async def analyze_report(file: UploadFile = File(...)):
    try:
        content = await file.read()
        result = processor.parse_spectro_report(content)
        return result
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

@app.post("/api/generate-excel")
async def generate_excel(payload: Dict[str, Any] = Body(...)):
    try:
        settings = payload.get("settings", {})
        data_to_fill = payload.get("data", {})
        
        # Resolve template path
        template_name = settings.get("mtc_template_path", "Final correct.xlsx")
        
        # Robust path resolution:
        # 1. Try absolute path as provided
        # 2. Try as relative path to ROOT_DIR
        # 3. Try just the filename in ROOT_DIR (handles Windows paths on Linux)
        
        possible_paths = [
            template_name,
            os.path.join(ROOT_DIR, template_name),
            os.path.join(ROOT_DIR, os.path.basename(template_name))
        ]
        
        template_path = None
        for p in possible_paths:
            if os.path.exists(p) and os.path.isfile(p):
                template_path = p
                break
        
        if not template_path:
             raise FileNotFoundError(f"Template file not found. Tried: {possible_paths}")
        
        generator = ExcelGenerator(settings)
        excel_bytes = generator.generate(template_path, data_to_fill)
        
        return StreamingResponse(
            io.BytesIO(excel_bytes),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=Generated_Report.xlsx"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/api/health")
async def health_check():
    return {"status": "healthy"}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)
