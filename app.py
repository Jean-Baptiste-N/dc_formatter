"""
FastAPI application for DC Formatter
Provides REST API endpoints for document processing
"""
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from urllib.parse import quote
from pathlib import Path
import tempfile
import shutil
import logging
import sys

# Add project root to path
sys.path.insert(0, str(Path(__file__).parent))

from tools.extract_xml_raw import export_all_xml
from tools.parse_template import extract_page_dimensions_from_template
from tools.parse_xml_raw_to_json_raw import xml_to_json
from tools.process_json_raw_to_json_transformed import apply_tags_and_styles
from tools.render_json_transformed_to_docx import json_to_docx

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Constants
TEMPLATE_PATH = Path("/app/TEMPLATE/TEMPLATE.docx")
UPLOAD_DIR = Path("/app/DC_SOURCES")  # Bind volume from ./uploads
OUTPUT_DIR = Path("/app/OUTPUTS_FORMATTED")  # Bind volume to ./output
OUTPUT1_XML_RAW = Path("/app/OUTPUT1_XML-RAW")  # Intermediate: Raw XML
OUTPUT2_JSON_RAW = Path("/app/OUTPUT2_JSON-RAW")  # Intermediate: Raw JSON
OUTPUT3_JSON_TRANSFORMED = Path("/app/OUTPUT3_JSON-TRANSFORMED")  # Intermediate: Transformed JSON
OUTPUT4_DOCX_RESULT = Path("/app/OUTPUT4_DOCX-RESULT")  # Final: DOCX result

# Create output directories with proper permissions
for output_path in [UPLOAD_DIR, OUTPUT_DIR, OUTPUT1_XML_RAW, OUTPUT2_JSON_RAW, OUTPUT3_JSON_TRANSFORMED, OUTPUT4_DOCX_RESULT]:
    try:
        output_path.mkdir(exist_ok=True, mode=0o777)
        # Ensure directory is writable
        output_path.chmod(0o777)
        logger.info(f"✓ Directory ready: {output_path.absolute()}")
    except PermissionError as e:
        logger.warning(f"⚠ Permission warning for {output_path}: {e} (may be OK if parent volume has correct perms)")
    except Exception as e:
        logger.warning(f"Could not set permissions for {output_path}: {e}")

app = FastAPI(
    title="DC Formatter API",
    description="Document transformation pipeline API",
    version="1.0.0"
)

# Configure CORS to expose Content-Disposition header for file downloads
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
    expose_headers=["Content-Disposition"],
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
    file: UploadFile = File(...)
):
    """
    Process a DOCX document through the full pipeline

    Args:
        file: DOCX file to process

    Returns:
        Processed DOCX file into OUTPUTS_FORMATTED directory
    """
    if not file.filename.endswith('.docx'):
        raise HTTPException(
            status_code=400,
            detail="File must be a DOCX document (.docx extension required)"
        )

    # Create temporary directory for intermediate processing
    temp_dir = Path(tempfile.mkdtemp(prefix="dc_formatter_"))
    try:
        logger.info(f"Processing file: {file.filename}")

        # Stage 0: Save uploaded file temporarily (NOT to bindable volume yet)
        temp_source_docx = temp_dir / file.filename
        try:
            logger.info(f"Saving temp file: {temp_source_docx}")
            with open(temp_source_docx, "wb") as f:
                content = await file.read()
                f.write(content)
            logger.info(f"✓ Temp file saved: {temp_source_docx}")
        except Exception as e:
            logger.error(f"Error saving temp file: {e}")
            raise HTTPException(status_code=500, detail=f"Error saving file: {str(e)}")

        # Also try to save to DC_SOURCES for reference (but don't fail if it doesn't work)
        try:
            source_docx_backup = UPLOAD_DIR / file.filename
            logger.info(f"Attempting to backup to DC_SOURCES: {source_docx_backup}")
            shutil.copy2(temp_source_docx, source_docx_backup)
            logger.info(f"✓ Backup saved to DC_SOURCES")
        except PermissionError:
            logger.warning(f"⚠ Could not save backup to DC_SOURCES (permission denied) - continuing with processing")
        except Exception as e:
            logger.warning(f"⚠ Could not save backup to DC_SOURCES: {e} - continuing with processing")

        # Extract page dimensions once
        logger.info("Extracting template dimensions...")
        dims = extract_page_dimensions_from_template(TEMPLATE_PATH)
        logger.info(f"Template dimensions: {dims}")

        # Stage 1: Extract XML (saves to OUTPUT1_XML_RAW)
        logger.info("Stage 1: Extracting XML from DOCX...")
        global_xml_path = export_all_xml(temp_source_docx, str(OUTPUT1_XML_RAW))
        logger.info(f"✓ XML extracted to OUTPUT1: {global_xml_path}")

        # Stage 2: Convert XML to JSON raw (saves to OUTPUT2_JSON_RAW)
        logger.info("Stage 2: Converting XML to JSON (raw)...")
        raw_json_path = xml_to_json(global_xml_path, str(OUTPUT2_JSON_RAW))
        logger.info(f"✓ Raw JSON created in OUTPUT2: {raw_json_path}")

        # Stage 3: Transform JSON (apply tags and styles, saves to OUTPUT3_JSON_TRANSFORMED)
        logger.info("Stage 3: Transforming JSON (applying tags and styles)...")
        transformed_json_path = apply_tags_and_styles(raw_json_path, str(OUTPUT3_JSON_TRANSFORMED), dims)
        logger.info(f"✓ Transformed JSON created in OUTPUT3: {transformed_json_path}")

        # Stage 4: Render back to DOCX (saves to OUTPUT4 and OUTPUT_DIR)
        logger.info("Stage 4: Rendering JSON to DOCX...")
        final_docx_path = json_to_docx(transformed_json_path, TEMPLATE_PATH, str(OUTPUT4_DOCX_RESULT))
        logger.info(f"✓ Final DOCX created in OUTPUT4: {final_docx_path}")

        # Also copy to OUTPUTS_FORMATTED for download
        try:
            output_filename = f"{Path(file.filename).stem}_formatted.docx"
            output_path = OUTPUT_DIR / output_filename
            logger.info(f"Copying result to OUTPUTS_FORMATTED: {output_path}")
            shutil.copy2(final_docx_path, output_path)
            logger.info(f"✓ Result copied to OUTPUTS_FORMATTED: {output_path}")
            final_docx_path = output_path
        except PermissionError as e:
            logger.warning(f"⚠ Could not copy to OUTPUTS_FORMATTED (permission denied): {e} - using OUTPUT4 version")
        except Exception as e:
            logger.warning(f"⚠ Could not copy to OUTPUTS_FORMATTED: {e} - using OUTPUT4 version")

        logger.info(f"Successfully processed document: {file.filename}")

        # Return the processed file for download with proper filename header
        output_filename = f"{Path(file.filename).stem}_formatted.docx"
        return FileResponse(
            path=final_docx_path,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            filename=output_filename,
            headers={"Content-Disposition": f'attachment; filename="{output_filename}"'}
        )

    except Exception as e:
        logger.error(f"Error processing document: {str(e)}", exc_info=True)
        raise HTTPException(
            status_code=500,
            detail=f"Error processing document: {str(e)}"
        )
    finally:
        # Cleanup temporary files
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
    Process multiple DOCX documents through the full pipeline

    Returns: JSON response with file paths and status for each file
    """
    results = []

    # Extract page dimensions once
    dims = extract_page_dimensions_from_template(TEMPLATE_PATH)

    for file in files:
        if not file.filename.endswith('.docx'):
            results.append({
                "filename": file.filename,
                "status": "error",
                "error": "File must be DOCX format"
            })
            continue

        try:
            logger.info(f"Batch: Processing {file.filename}")

            # Save to DC_SOURCES
            source_docx = UPLOAD_DIR / file.filename
            with open(source_docx, "wb") as f:
                content = await file.read()
                f.write(content)
            logger.info(f"Batch: Saved {file.filename} to DC_SOURCES")

            # Execute pipeline with proper output directories
            global_xml_path = export_all_xml(source_docx, str(OUTPUT1_XML_RAW))
            raw_json_path = xml_to_json(global_xml_path, str(OUTPUT2_JSON_RAW))
            transformed_json_path = apply_tags_and_styles(raw_json_path, str(OUTPUT3_JSON_TRANSFORMED), dims)
            final_docx_path = json_to_docx(transformed_json_path, TEMPLATE_PATH, str(OUTPUT4_DOCX_RESULT))

            # Copy to OUTPUTS_FORMATTED for download
            output_filename = f"{Path(file.filename).stem}_formatted.docx"
            output_path = OUTPUT_DIR / output_filename
            shutil.copy2(final_docx_path, output_path)

            results.append({
                "filename": file.filename,
                "status": "success",
                "output_path": str(output_path)
            })
            logger.info(f"Batch: Successfully processed {file.filename}")

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
