from flask import Flask, request, send_file, jsonify, render_template, Response
from typing import Union
import os
import win32com.client
import pythoncom
from pathlib import Path
import tempfile
import fitz
import logging

app = Flask(__name__)

# Add this new route for the landing page
@app.route('/')
def index(): 
    return render_template('index.html')

def allowed_file(filename: str) -> bool:
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in {'pdf'}

@app.route('/compress', methods=['POST'])
def compress_pdf() -> Union[Response, tuple[Response, int]]:
    if 'file' not in request.files:
        return jsonify({'error': 'No file provided'}), 400
    
    file = request.files['file']
    try:
        compression_level = int(request.form.get('level', 50))
    except ValueError:
        return jsonify({'error': 'Invalid compression level'}), 400
    
    if not file or not file.filename:
        return jsonify({'error': 'No file selected'}), 400
        
    if not allowed_file(file.filename):
        return jsonify({'error': 'Invalid file type. Only PDF files are allowed'}), 400

    temp_input = None
    output = None
    doc = None
    
    try:
        # Save input file
        temp_input = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
        temp_input_path = temp_input.name
        temp_input.close()
        file.save(temp_input_path)
        
        # Open and compress PDF
        doc = fitz.open(temp_input_path)
        output = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
        output_path = output.name
        output.close()
        
        compression_level = int(request.form.get('level', 50))
        
        # Process each page
        for page_num in range(len(doc)):
            page = doc[page_num]
            
            # Process images with compression
            image_list = page.get_images()
            for img in image_list:
                xref = img[0]
                base_image = doc.extract_image(xref)
                
                if base_image:
                    # Calculate quality based on compression level
                    quality = max(5, min(95, 100 - compression_level))
                    
                    # Remove image and clean page
                    page.clean_contents()
                    
                    if compression_level > 50:
                        # Reduce image resolution for higher compression
                        page.set_rotation(0)
                        # Scale down images using mediabox
                        scale = 0.8 if compression_level > 70 else 0.9
                        new_width = page.rect.width * scale
                        new_height = page.rect.height * scale
                        page.set_mediabox((0, 0, new_width, new_height))
                        
            # Apply text compression
            if compression_level > 30:
                page.clean_contents()
                page.wrap_contents()
                
            # Extreme compression for high levels
            if compression_level > 70:
                page.clean_contents()
                page.wrap_contents()
                # Additional cleaning pass
                page.clean_contents()
        
        # Save with optimized parameters
        doc.save(
            output_path,
            garbage=4,
            clean=True,
            deflate=True,
            deflate_images=True,
            deflate_fonts=True,
            pretty=False,
            linear=True,
            ascii=False,
            expand=0
        )
        
        doc.close()
        doc = None
        
        # Clean up input file
        if os.path.exists(temp_input_path):
            os.unlink(temp_input_path)
        
        # Send compressed file
        response = send_file(
            output_path,
            mimetype='application/pdf',
            as_attachment=True,
            download_name='compressed.pdf'
        )
        
        # Delete the output file after sending
        @response.call_on_close
        def cleanup():
            if os.path.exists(output_path):
                try:
                    os.unlink(output_path)
                except:
                    pass
                    
        return response
        
    except Exception as e:
        # Clean up on error
        if doc:
            doc.close()
            
        if temp_input and os.path.exists(temp_input.name):
            try:
                os.unlink(temp_input.name)
            except:
                pass
                
        if output and os.path.exists(output.name):
            try:
                os.unlink(output.name)
            except:
                pass
                
        print(f"Compression error: {str(e)}")
        return jsonify({'error': f'Error compressing PDF: {str(e)}'}), 500

@app.route('/merge', methods=['POST'])
def merge_pdfs() -> Union[Response, tuple[Response, int]]:
    if 'files[]' not in request.files:
        return jsonify({'error': 'No files provided'}), 400
    
    files = request.files.getlist('files[]')
    if not files or not all(allowed_file(f.filename or '') for f in files):
        return jsonify({'error': 'Invalid file type'}), 400

    temp_files = []
    merged_pdf = None
    output = None
    
    try:
        merged_pdf = fitz.open()
        
        # Save and merge PDFs
        for file in files:
            temp_input = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
            temp_files.append(temp_input.name)
            file.save(temp_input.name)
            pdf_document = fitz.open(temp_input.name)
            merged_pdf.insert_pdf(pdf_document)
            pdf_document.close()
        
        # Save merged result
        output = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
        output_path = output.name
        output.close()  # Close the file handle immediately
        
        merged_pdf.save(output_path)
        merged_pdf.close()
        
        # Clean up temp files
        for temp_file in temp_files:
            try:
                if os.path.exists(temp_file):
                    os.unlink(temp_file)
            except Exception as e:
                print(f"Warning: Could not delete temp file {temp_file}: {e}")
        
        return send_file(
            output_path,
            mimetype='application/pdf',
            as_attachment=True,
            download_name='merged.pdf'
        )

    except Exception as e:
        if merged_pdf:
            merged_pdf.close()
            
        # Clean up on error
        for temp_file in temp_files:
            try:
                if os.path.exists(temp_file):
                    os.unlink(temp_file)
            except Exception as cleanup_error:
                print(f"Warning: Could not delete temp file {temp_file}: {cleanup_error}")
                
        if output and os.path.exists(output.name):
            try:
                os.unlink(output.name)
            except Exception as cleanup_error:
                print(f"Warning: Could not delete output file: {cleanup_error}")
                
        return jsonify({'error': f'Error merging PDFs: {str(e)}'}), 500


@app.route('/convert-ppt', methods=['POST'])
def convert_ppt():
    if not os.name == 'nt':  # Check if not running on Windows
        return jsonify({
            'error': 'PowerPoint conversion is only available on Windows servers. This feature is not supported in the cloud deployment.'
        }), 400
        
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file provided'}), 400
            
        file = request.files['file']
        if not file or not file.filename:
            return jsonify({'error': 'No file selected'}), 400
            
        if not file.filename.lower().endswith(('.ppt', '.pptx')):
            return jsonify({'error': 'Invalid file format'}), 400

        # Create temp directory if it doesn't exist
        temp_dir = Path('temp')
        temp_dir.mkdir(exist_ok=True, parents=True)

        # Generate unique filenames
        import uuid
        unique_id = str(uuid.uuid4())
        temp_ppt_path = temp_dir / f"{unique_id}_{file.filename}"
        temp_pdf_path = temp_dir / f"{unique_id}.pdf"

        try:
            pythoncom.CoInitialize()
            # Save the uploaded file
            file.save(str(temp_ppt_path))

            # Initialize PowerPoint with error handling
            powerpoint = None
            try:
                powerpoint = win32com.client.Dispatch("Powerpoint.Application")
                powerpoint.Visible = True  # Set visibility to True
                logging.info("PowerPoint application initialized")
                # Convert to PDF with absolute paths
                presentation = powerpoint.Presentations.Open(str(temp_ppt_path.absolute()))
                presentation.SaveAs(str(temp_pdf_path.absolute()), 32)  # 32 = PDF format
                presentation.Close()
            finally:
                if powerpoint:
                    powerpoint.Quit()
            pythoncom.CoUninitialize()

            if not temp_pdf_path.exists():
                raise Exception("PDF conversion failed - output file not created")

            # Send the converted file
            response = send_file(
                str(temp_pdf_path),
                mimetype='application/pdf',
                as_attachment=True,
                download_name=f"{Path(file.filename).stem}.pdf"
            )

            # Clean up after sending
            @response.call_on_close
            def cleanup():
                try:
                    if temp_ppt_path.exists():
                        os.unlink(str(temp_ppt_path))
                    if temp_pdf_path.exists():
                        os.unlink(str(temp_pdf_path))
                except Exception as e:
                    logging.error(f"Cleanup error: {e}")

            return response

        except Exception as e:
            # Clean up on error
            try:
                if temp_ppt_path.exists():
                    os.unlink(str(temp_ppt_path))
                if temp_pdf_path.exists():
                    os.unlink(str(temp_pdf_path))
            except:
                pass
            logging.error(f"PowerPoint conversion failed: {str(e)}")
            raise Exception(f"PowerPoint conversion failed: {str(e)}")

    except Exception as e:
        logging.error(f"PPT conversion error: {str(e)}")
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True)