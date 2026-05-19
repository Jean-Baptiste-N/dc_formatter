"""
FastAPI application for DC Formatter
Provides REST API endpoints for document processing
"""
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pathlib import Path
import tempfile
import shutil
import logging
import sys
from typing import Optional

# Add project root to path
sys.path.insert(0, str(Path(__file__).parent))

from tools3.extract_xml_raw import export_all_xml
from tools3.parse_template import extract_page_dimensions_from_template
from tools3.parse_xml_raw_to_json_raw import xml_to_json
from tools3.process_json_raw_to_json_transformed import apply_tags_and_styles
from tools3.render_json_transformed_to_docx import json_to_docx

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Constants
TEMPLATE_PATH = Path("assets/TEMPLATE.docx")
UPLOAD_DIR = Path("uploads")
OUTPUT_DIR = Path("output")
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

app = FastAPI(
    title="DC Formatter API",
    description="Document transformation pipeline API",
    version="1.0.0"
)

# Mount static files (HTML interface)
static_dir = Path(__file__).parent
if (static_dir / "index.html").exists():
    app.mount("/static", StaticFiles(directory=str(static_dir), html=True), name="static")


@app.get("/health")
async def health_check():
    """Health check endpoint for container monitoring"""
    return {
        "status": "healthy",
        "service": "dc-formatter",
        "version": "1.0.0"
    }


@app.get("/")
async def root():
    """Serve the web interface"""
    html_path = Path(__file__).parent / "index.html"
    if html_path.exists():
        return FileResponse(html_path, media_type="text/html")
    
    # Fallback if HTML not found
    return {
        "name": "DC Formatter API",
        "description": "Transform DOCX documents through extraction, conversion, and rendering pipeline",
        "endpoints": {
            "health": "/health",
            "ui": "/static/index.html",
            "process": "/process (POST)",
            "docs": "/docs"
        }
    }


@app.post("/process")
async def process_document(
    file: UploadFile = File(...),
    keep_temp: Optional[bool] = False
):
    """
    Process a DOCX document through the full pipeline
    
    Args:
        file: DOCX file to process
        keep_temp: Keep temporary files for debugging (default: False)
    
    Returns:
        Processed DOCX file
    """
    if not file.filename.endswith('.docx'):
        raise HTTPException(
            status_code=400,
            detail="File must be a DOCX document (.docx extension required)"
        )
    
    # Create temporary directory for processing
    temp_dir = Path(tempfile.mkdtemp(prefix="dc_formatter_"))
    try:
        logger.info(f"Processing file: {file.filename}")
        
        # Save uploaded file
        source_docx = temp_dir / "source.docx"
        with open(source_docx, "wb") as f:
            content = await file.read()
            f.write(content)
        logger.info(f"Saved upload to: {source_docx}")
        
        # Create output paths
        output_xml_raw = temp_dir / "output_raw.xml"
        output_json_raw = temp_dir / "output_raw.json"
        output_json_transformed = temp_dir / "output_transformed.json"
        output_docx = temp_dir / "output.docx"
        
        # Stage 1: Extract XML and dimensions
        logger.info("Stage 1: Extracting XML and template dimensions...")
        export_all_xml(source_docx, output_xml_raw)
        dims = extract_page_dimensions_from_template(TEMPLATE_PATH)
        logger.info(f"Template dimensions: {dims}")
        
        # Stage 2: Convert XML to JSON raw
        logger.info("Stage 2: Converting XML to JSON (raw)...")
        xml_to_json(output_xml_raw, output_json_raw, dims)
        
        # Stage 3: Transform JSON (apply tags and styles)
        logger.info("Stage 3: Transforming JSON (applying tags and styles)...")
        apply_tags_and_styles(output_json_raw, output_json_transformed)
        
        # Stage 4: Render back to DOCX
        logger.info("Stage 4: Rendering JSON to DOCX...")
        json_to_docx(output_json_transformed, TEMPLATE_PATH, output_docx)
        
        logger.info(f"Successfully processed document: {file.filename}")
        
        # Save the processed file to output directory
        output_filename = f"processed_{file.filename}"
        output_path = OUTPUT_DIR / output_filename
        shutil.copy2(output_docx, output_path)
        logger.info(f"Saved processed document to: {output_path}")
        
        # Return the processed file
        return FileResponse(
            path=output_docx,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            filename=output_filename
        )
        
    except Exception as e:
        logger.error(f"Error processing document: {str(e)}", exc_info=True)
        raise HTTPException(
            status_code=500,
            detail=f"Error processing document: {str(e)}"
        )
    finally:
        # Cleanup temporary files (unless keep_temp is True for debugging)
        if not keep_temp:
            try:
                shutil.rmtree(temp_dir)
                logger.info(f"Cleaned up temporary directory: {temp_dir}")
            except Exception as e:
                logger.warning(f"Could not remove temp directory {temp_dir}: {e}")


@app.post("/process-batch")
async def process_batch(
    files: list[UploadFile] = File(...)
):
    """
    Process multiple DOCX documents
    
    Note: Returns a JSON response with file paths (not direct file downloads)
    """
    results = []
    
    for file in files:
        try:
            if not file.filename.endswith('.docx'):
                results.append({
                    "filename": file.filename,
                    "status": "error",
                    "error": "File must be DOCX format"
                })
                continue
            
            temp_dir = Path(tempfile.mkdtemp(prefix="dc_formatter_"))
            
            # Save and process (simplified for batch)
            source_docx = temp_dir / "source.docx"
            with open(source_docx, "wb") as f:
                content = await file.read()
                f.write(content)
            
            output_docx = temp_dir / "output.docx"
            
            # Process pipeline
            export_all_xml(source_docx, temp_dir / "raw.xml")
            dims = extract_page_dimensions_from_template(TEMPLATE_PATH)
            xml_to_json(temp_dir / "raw.xml", temp_dir / "raw.json", dims)
            apply_tags_and_styles(temp_dir / "raw.json", temp_dir / "transformed.json")
            json_to_docx(temp_dir / "transformed.json", TEMPLATE_PATH, output_docx)
            
            # Save to output directory
            output_filename = f"processed_{file.filename}"
            output_path = OUTPUT_DIR / output_filename
            shutil.copy2(output_docx, output_path)
            
            results.append({
                "filename": file.filename,
                "status": "success",
                "output_path": str(output_path)
            })
            logger.info(f"Batch: Processed {file.filename} -> {output_path}")
            
        except Exception as e:
            logger.error(f"Batch error for {file.filename}: {str(e)}")
            results.append({
                "filename": file.filename,
                "status": "error",
                "error": str(e)
            })
    
    return {"results": results}


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(
        app,
        host="0.0.0.0",
        port=8000,
        log_level="info"
    )
