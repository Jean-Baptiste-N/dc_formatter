"""
Test suite for DC Formatter API
"""
import pytest
from pathlib import Path
from fastapi.testclient import TestClient
from app import app

client = TestClient(app)


def test_health_check():
    """Test the health check endpoint"""
    response = client.get("/health")
    assert response.status_code == 200
    data = response.json()
    assert data["status"] == "healthy"
    assert data["service"] == "dc-formatter"


def test_root_endpoint():
    """Test the root endpoint"""
    response = client.get("/")
    assert response.status_code == 200
    data = response.json()
    assert "name" in data
    assert "endpoints" in data


def test_process_invalid_file_type():
    """Test that non-DOCX files are rejected"""
    # Create a fake text file
    with open("test.txt", "w") as f:
        f.write("This is not a DOCX file")
    
    with open("test.txt", "rb") as f:
        response = client.post(
            "/process",
            files={"file": ("test.txt", f, "text/plain")}
        )
    
    assert response.status_code == 400
    assert "must be a DOCX document" in response.json()["detail"]
    
    # Cleanup
    Path("test.txt").unlink()


def test_process_missing_file():
    """Test that missing file returns an error"""
    response = client.post("/process")
    assert response.status_code == 422  # Validation error


@pytest.mark.asyncio
async def test_process_valid_docx():
    """Test processing a valid DOCX file"""
    # This test requires a real DOCX file
    docx_path = Path("DC_SOURCES") / "DC_JNZ_2026.docx"
    
    if not docx_path.exists():
        pytest.skip(f"Test DOCX file not found: {docx_path}")
    
    with open(docx_path, "rb") as f:
        response = client.post(
            "/process",
            files={"file": ("test.docx", f, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")}
        )
    
    # Check response
    assert response.status_code == 200
    assert response.headers["content-type"] == "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    assert len(response.content) > 0


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
