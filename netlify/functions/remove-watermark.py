import os
import re
import base64
import zipfile
import shutil
import tempfile
import xml.etree.ElementTree as ET
from email.parser import BytesParser
from email.policy import default

# Since Netlify runs on AWS Lambda, we can import fitz (PyMuPDF) from our requirements
try:
    import fitz
except ImportError:
    fitz = None

def clean_pptx(in_path, out_path):
    with tempfile.TemporaryDirectory() as temp_dir:
        with zipfile.ZipFile(in_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
            
        rels_dir = os.path.join(temp_dir, "ppt", "slideLayouts", "_rels")
        layouts_dir = os.path.join(temp_dir, "ppt", "slideLayouts")
        
        cleaned_any = False
        if os.path.exists(rels_dir):
            for rels_file in os.listdir(rels_dir):
                if not rels_file.endswith(".xml.rels"):
                    continue
                    
                rels_path = os.path.join(rels_dir, rels_file)
                with open(rels_path, "r", encoding="utf-8", errors="ignore") as f:
                    rels_content = f.read()
                    
                match = re.search(r'Id="([^"]+)"[^>]+Target="[^"]*gamma\.app', rels_content)
                if not match:
                    continue
                    
                rel_id = match.group(1)
                layout_filename = rels_file.replace(".rels", "")
                layout_path = os.path.join(layouts_dir, layout_filename)
                
                if os.path.exists(layout_path):
                    with open(layout_path, "r", encoding="utf-8", errors="ignore") as lf:
                        layout_content = lf.read()
                    
                    pic_pattern = re.compile(r'<p:pic>.*?</p:pic>', re.DOTALL)
                    pics = pic_pattern.findall(layout_content)
                    
                    cleaned_layout = layout_content
                    for pic in pics:
                        if f'hlinkClick r:id="{rel_id}"' in pic or f'r:id="{rel_id}"' in pic:
                            cleaned_layout = cleaned_layout.replace(pic, "")
                            cleaned_any = True
                            
                    with open(layout_path, "w", encoding="utf-8") as lf:
                        lf.write(cleaned_layout)
        
        if os.path.exists(out_path):
            os.remove(out_path)
            
        with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as zip_out:
            for root, dirs, files in os.walk(temp_dir):
                for file in files:
                    full_path = os.path.join(root, file)
                    rel_path = os.path.relpath(full_path, temp_dir)
                    zip_out.write(full_path, rel_path)
                    
        return cleaned_any

def clean_pdf(in_path, out_path):
    if not fitz:
        raise ImportError("PyMuPDF is not installed on the serverless environment.")
        
    doc = fitz.open(in_path)
    cleaned_any = False
    
    for page_num in range(len(doc)):
        page = doc[page_num]
        
        links = page.get_links()
        gamma_links = []
        for link in links:
            uri = link.get("uri", "")
            if "gamma.app" in uri:
                gamma_links.append(link)
                
        watermark_rects = []
        for link in gamma_links:
            rect = link["from"]
            watermark_rects.append(rect)
            page.delete_link(link)
            cleaned_any = True
            
        if not watermark_rects:
            p_rect = page.rect
            fallback_rect = fitz.Rect(p_rect.width * 0.7, p_rect.height * 0.8, p_rect.width, p_rect.height)
            watermark_rects.append(fallback_rect)
            
        images = page.get_images()
        for img_info in images:
            xref = img_info[0]
            name = img_info[7]
            rects = page.get_image_rects(xref)
            
            for r in rects:
                is_watermark = False
                for w_rect in watermark_rects:
                    intersection = r & w_rect
                    if intersection.is_empty:
                        continue
                        
                    r_area = r.get_area()
                    intersect_area = intersection.get_area()
                    
                    if intersect_area / r_area > 0.8 and r_area < 80000:
                        is_watermark = True
                        break
                        
                    if abs(r.x0 - w_rect.x0) < 5 and abs(r.y0 - w_rect.y0) < 5 and abs(r.x1 - w_rect.x1) < 5 and abs(r.y1 - w_rect.y1) < 5:
                        is_watermark = True
                        break
                        
                if not is_watermark and r.x0 > page.rect.width * 0.7 and r.y0 > page.rect.height * 0.8:
                    if r.width < 250 and r.height < 60:
                        is_watermark = True
                        
                if is_watermark:
                    try:
                        page.delete_image(xref)
                        cleaned_any = True
                    except Exception:
                        pass
                    break
                    
    doc.save(out_path, garbage=3, deflate=True)
    return cleaned_any

def handler(event, context):
    # Enable CORS
    cors_headers = {
        "Access-Control-Allow-Origin": "*",
        "Access-Control-Allow-Headers": "Content-Type",
        "Access-Control-Allow-Methods": "POST, OPTIONS"
    }
    
    if event.get("httpMethod") == "OPTIONS":
        return {
            "statusCode": 200,
            "headers": cors_headers,
            "body": ""
        }
        
    try:
        # Get content type (case-insensitive)
        headers = event.get("headers", {})
        content_type = ""
        for k, v in headers.items():
            if k.lower() == "content-type":
                content_type = v
                break
                
        if not content_type:
            return {
                "statusCode": 400,
                "headers": cors_headers,
                "body": '{"error": "Missing Content-Type header."}'
            }
            
        # Parse body
        body = event.get("body", "")
        if event.get("isBase64Encoded", False):
            body_bytes = base64.b64decode(body)
        else:
            body_bytes = body.encode("utf-8") if isinstance(body, str) else body
            
        # Parse multipart body using email BytesParser
        payload = f"Content-Type: {content_type}\r\n\r\n".encode("utf-8") + body_bytes
        msg = BytesParser(policy=default).parsebytes(payload)
        
        file_bytes = None
        filename = ""
        
        for part in msg.iter_parts():
            part_filename = part.get_filename()
            if part_filename:
                file_bytes = part.get_content()
                filename = part_filename
                break
                
        if not file_bytes:
            return {
                "statusCode": 400,
                "headers": cors_headers,
                "body": '{"error": "No file uploaded in the request."}'
            }
            
        ext = os.path.splitext(filename)[1].lower()
        if ext not in [".pptx", ".pdf"]:
            return {
                "statusCode": 400,
                "headers": cors_headers,
                "body": '{"error": "Unsupported file format. Please upload a .pptx or .pdf file."}'
            }
            
        # Run cleaner using temp files
        with tempfile.TemporaryDirectory() as temp_dir:
            in_path = os.path.join(temp_dir, f"orig_{filename}")
            out_path = os.path.join(temp_dir, f"cleaned_{filename}")
            
            with open(in_path, "wb") as f:
                f.write(file_bytes)
                
            if ext == ".pptx":
                clean_pptx(in_path, out_path)
            else:
                clean_pdf(in_path, out_path)
                
            if not os.path.exists(out_path):
                return {
                    "statusCode": 500,
                    "headers": cors_headers,
                    "body": '{"error": "Processing failed."}'
                }
                
            with open(out_path, "rb") as f:
                cleaned_bytes = f.read()
                
        # Return cleaned file as a base64 encoded response
        response_headers = {
            **cors_headers,
            "Content-Type": "application/octet-stream",
            "Content-Disposition": f'attachment; filename="no_watermark_{filename}"'
        }
        
        return {
            "statusCode": 200,
            "headers": response_headers,
            "body": base64.b64encode(cleaned_bytes).decode("utf-8"),
            "isBase64Encoded": True
        }
        
    except Exception as e:
        return {
            "statusCode": 500,
            "headers": cors_headers,
            "body": f'{{"error": "Internal server error: {str(e)}"}}'
        }
