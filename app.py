import os
import re
import zipfile
import shutil
import tempfile
import fitz  # PyMuPDF
from flask import Flask, request, render_template, send_file, jsonify
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = tempfile.gettempdir()
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50 MB limits

def clean_pptx(in_path, out_path):
    """
    Strips Gamma watermarks from a PPTX file.
    Gamma embeds a watermark as a picture in the slide layouts pointing to a gamma.app link relation.
    """
    with tempfile.TemporaryDirectory() as temp_dir:
        # Extract the PPTX zip contents
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
                    
                # Find the relationship ID targeting gamma.app
                match = re.search(r'Id="([^"]+)"[^>]+Target="[^"]*gamma\.app', rels_content)
                if not match:
                    continue
                    
                rel_id = match.group(1)
                
                # Retrieve the corresponding slide layout XML
                layout_filename = rels_file.replace(".rels", "")
                layout_path = os.path.join(layouts_dir, layout_filename)
                
                if os.path.exists(layout_path):
                    with open(layout_path, "r", encoding="utf-8", errors="ignore") as lf:
                        layout_content = lf.read()
                    
                    # Locate and remove the entire <p:pic> XML element referencing this relationship ID
                    pic_pattern = re.compile(r'<p:pic>.*?</p:pic>', re.DOTALL)
                    pics = pic_pattern.findall(layout_content)
                    
                    cleaned_layout = layout_content
                    for pic in pics:
                        if f'hlinkClick r:id="{rel_id}"' in pic or f'r:id="{rel_id}"' in pic:
                            cleaned_layout = cleaned_layout.replace(pic, "")
                            cleaned_any = True
                            
                    with open(layout_path, "w", encoding="utf-8") as lf:
                        lf.write(cleaned_layout)
        
        # Build the new clean PPTX zip
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
    """
    Strips Gamma watermarks from a PDF file.
    Identifies 'gamma.app' link annotations, removes them, and wipes the overlapping watermark image XObjects.
    """
    doc = fitz.open(in_path)
    cleaned_any = False
    
    for page_num in range(len(doc)):
        page = doc[page_num]
        
        # Grab hyperlink annotations
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
            
        # Fallback to bottom right quadrant if no hyperlink matches
        if not watermark_rects:
            p_rect = page.rect
            fallback_rect = fitz.Rect(p_rect.width * 0.7, p_rect.height * 0.8, p_rect.width, p_rect.height)
            watermark_rects.append(fallback_rect)
            
        # Find images intersecting these zones
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
                    
                    # Watermark image should be relatively small (< 80000 sq points)
                    # and must represent > 80% of the intersection, or be an exact match
                    if intersect_area / r_area > 0.8 and r_area < 80000:
                        is_watermark = True
                        break
                        
                    if abs(r.x0 - w_rect.x0) < 5 and abs(r.y0 - w_rect.y0) < 5 and abs(r.x1 - w_rect.x1) < 5 and abs(r.y1 - w_rect.y1) < 5:
                        is_watermark = True
                        break
                        
                # Absolute fallback for small bottom-right images
                if not is_watermark and r.x0 > page.rect.width * 0.7 and r.y0 > page.rect.height * 0.8:
                    if r.width < 250 and r.height < 60:
                        is_watermark = True
                        
                if is_watermark:
                    try:
                        page.delete_image(xref)
                        cleaned_any = True
                    except Exception as e:
                        # Log error and continue
                        print(f"Error deleting image {xref} on page {page_num+1}: {e}")
                    break
                    
    doc.save(out_path, garbage=3, deflate=True)
    return cleaned_any

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/remove-watermark', methods=['POST'])
def remove_watermark():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part in request.'}), 400
        
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No file selected.'}), 400
        
    filename = secure_filename(file.filename)
    ext = os.path.splitext(filename)[1].lower()
    
    if ext not in ['.pptx', '.pdf']:
        return jsonify({'error': 'Unsupported file format. Please upload a .pptx or .pdf file.'}), 400
        
    # Generate local filenames
    input_path = os.path.join(app.config['UPLOAD_FOLDER'], f"orig_{filename}")
    output_path = os.path.join(app.config['UPLOAD_FOLDER'], f"cleaned_{filename}")
    
    try:
        # Save uploaded file
        file.save(input_path)
        
        # Clean based on extension
        if ext == '.pptx':
            cleaned = clean_pptx(input_path, output_path)
        else:
            cleaned = clean_pdf(input_path, output_path)
            
        if not os.path.exists(output_path):
            return jsonify({'error': 'Failed to process file.'}), 500
            
        # Read the file into memory to clean up disk immediately
        import io
        with open(output_path, "rb") as f:
            file_data = f.read()
            
        # Clean up temp files immediately
        for temp_path in [input_path, output_path]:
            try:
                if os.path.exists(temp_path):
                    os.remove(temp_path)
            except Exception:
                pass
                
        # Send back the cleaned file from memory
        return send_file(
            io.BytesIO(file_data),
            as_attachment=True,
            download_name=f"no_watermark_{filename}"
        )
        
    except Exception as e:
        # Clean up temp files if they still exist on error
        for temp_path in [input_path, output_path]:
            try:
                if os.path.exists(temp_path):
                    os.remove(temp_path)
            except Exception:
                pass
        return jsonify({'error': f"Internal error during cleaning: {str(e)}"}), 500

if __name__ == '__main__':
    app.run(debug=True, port=5000)
