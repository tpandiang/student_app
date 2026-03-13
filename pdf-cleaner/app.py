import os
import re
import uuid
from collections import Counter
import fitz
from flask import Flask, request, send_file, render_template_string

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = os.path.join(os.path.dirname(__file__), 'uploads')
app.config['CLEANED_FOLDER'] = os.path.join(os.path.dirname(__file__), 'cleaned')
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100 MB max

HTML = '''
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>PDF Footer Remover</title>
  <style>
    * { margin: 0; padding: 0; box-sizing: border-box; }
    body {
      font-family: 'Segoe UI', system-ui, sans-serif;
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      min-height: 100vh;
      display: flex;
      align-items: center;
      justify-content: center;
      padding: 20px;
    }
    .card {
      background: #fff;
      border-radius: 16px;
      box-shadow: 0 20px 60px rgba(0,0,0,0.2);
      padding: 48px 40px;
      max-width: 520px;
      width: 100%;
      text-align: center;
    }
    h1 {
      font-size: 1.8rem;
      color: #1a1a2e;
      margin-bottom: 8px;
    }
    .subtitle {
      color: #64748b;
      font-size: 0.95rem;
      margin-bottom: 32px;
    }
    .upload-area {
      border: 2px dashed #cbd5e1;
      border-radius: 12px;
      padding: 40px 20px;
      cursor: pointer;
      transition: all 0.2s;
      margin-bottom: 24px;
      position: relative;
    }
    .upload-area:hover, .upload-area.dragover {
      border-color: #667eea;
      background: #f0f0ff;
    }
    .upload-area svg {
      width: 48px;
      height: 48px;
      color: #94a3b8;
      margin-bottom: 12px;
    }
    .upload-area p {
      color: #64748b;
      font-size: 0.9rem;
    }
    .upload-area .filename {
      color: #1a1a2e;
      font-weight: 600;
      font-size: 1rem;
      margin-top: 8px;
    }
    input[type="file"] {
      position: absolute;
      inset: 0;
      opacity: 0;
      cursor: pointer;
    }
    .mode-toggle {
      display: flex;
      gap: 8px;
      margin-bottom: 16px;
      justify-content: center;
    }
    .mode-btn {
      padding: 8px 20px;
      border: 2px solid #e2e8f0;
      border-radius: 8px;
      background: #fff;
      font-size: 0.85rem;
      font-weight: 600;
      color: #64748b;
      cursor: pointer;
      transition: all 0.2s;
    }
    .mode-btn.active {
      border-color: #667eea;
      color: #667eea;
      background: #f0f0ff;
    }
    .manual-input {
      display: none;
    }
    .manual-input.show {
      display: block;
    }
    .footer-input {
      width: 100%;
      padding: 12px 16px;
      border: 1px solid #e2e8f0;
      border-radius: 8px;
      font-size: 0.9rem;
      margin-bottom: 24px;
      outline: none;
      transition: border 0.2s;
    }
    .footer-input:focus {
      border-color: #667eea;
    }
    label.input-label {
      display: block;
      text-align: left;
      font-size: 0.85rem;
      font-weight: 600;
      color: #334155;
      margin-bottom: 6px;
    }
    .auto-info {
      background: #f0f4ff;
      border: 1px solid #d0daf0;
      border-radius: 8px;
      padding: 12px 16px;
      margin-bottom: 24px;
      font-size: 0.82rem;
      color: #475569;
      text-align: left;
      line-height: 1.5;
    }
    .btn {
      background: linear-gradient(135deg, #667eea, #764ba2);
      color: #fff;
      border: none;
      padding: 14px 32px;
      font-size: 1rem;
      font-weight: 600;
      border-radius: 8px;
      cursor: pointer;
      width: 100%;
      transition: opacity 0.2s;
    }
    .btn:hover { opacity: 0.9; }
    .btn:disabled {
      opacity: 0.5;
      cursor: not-allowed;
    }
    .result {
      margin-top: 24px;
      padding: 16px;
      border-radius: 8px;
    }
    .result.success {
      background: #f0fdf4;
      border: 1px solid #86efac;
    }
    .result.success a {
      color: #16a34a;
      font-weight: 600;
      text-decoration: none;
      font-size: 1.05rem;
    }
    .result.success a:hover { text-decoration: underline; }
    .result.success .info {
      color: #64748b;
      font-size: 0.82rem;
      margin-top: 6px;
    }
    .result.success .detected {
      color: #475569;
      font-size: 0.8rem;
      margin-top: 8px;
      text-align: left;
      background: #f8fafc;
      padding: 8px 12px;
      border-radius: 6px;
      border: 1px solid #e2e8f0;
    }
    .result.error {
      background: #fef2f2;
      border: 1px solid #fca5a5;
      color: #dc2626;
    }
    .spinner {
      display: none;
      margin: 20px auto 0;
      width: 36px;
      height: 36px;
      border: 3px solid #e2e8f0;
      border-top-color: #667eea;
      border-radius: 50%;
      animation: spin 0.8s linear infinite;
    }
    .spinner-text {
      display: none;
      color: #64748b;
      font-size: 0.85rem;
      margin-top: 8px;
    }
    @keyframes spin { to { transform: rotate(360deg); } }
  </style>
</head>
<body>
  <div class="card">
    <h1>PDF Footer Remover</h1>
    <p class="subtitle">Upload a PDF and remove footers automatically</p>

    <form id="uploadForm" action="/upload" method="post" enctype="multipart/form-data">
      <div class="upload-area" id="dropZone">
        <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor">
          <path stroke-linecap="round" stroke-linejoin="round" stroke-width="1.5"
            d="M12 16v-8m0 0l-3 3m3-3l3 3M3 16.5v2.25A2.25 2.25 0 005.25 21h13.5A2.25 2.25 0 0021 18.75V16.5" />
        </svg>
        <p>Drag & drop your PDF here or click to browse</p>
        <p class="filename" id="fileName"></p>
        <input type="file" name="pdf" id="pdfInput" accept=".pdf" required>
      </div>

      <div class="mode-toggle">
        <button type="button" class="mode-btn active" id="autoBtn" onclick="setMode('auto')">Auto-Detect</button>
        <button type="button" class="mode-btn" id="manualBtn" onclick="setMode('manual')">Manual</button>
      </div>

      <input type="hidden" name="mode" id="modeInput" value="auto">

      <div class="auto-info" id="autoInfo">
        Automatically scans the PDF and detects repeating text at the bottom of pages (footers, copyright lines, page numbers, watermarks). No input needed.
      </div>

      <div class="manual-input" id="manualSection">
        <label class="input-label">Footer text to remove</label>
        <input type="text" name="footer_text" class="footer-input" placeholder="e.g. Copyright, Confidential, Company Name">
      </div>

      <button type="submit" class="btn" id="submitBtn">Remove Footers & Download</button>
    </form>

    <div class="spinner" id="spinner"></div>
    <div class="spinner-text" id="spinnerText">Scanning pages for footers...</div>
    <div id="result"></div>
  </div>

  <script>
    const dropZone = document.getElementById('dropZone');
    const pdfInput = document.getElementById('pdfInput');
    const fileName = document.getElementById('fileName');
    const form = document.getElementById('uploadForm');
    const spinner = document.getElementById('spinner');
    const spinnerText = document.getElementById('spinnerText');
    const result = document.getElementById('result');
    const submitBtn = document.getElementById('submitBtn');

    function setMode(mode) {
      document.getElementById('modeInput').value = mode;
      document.getElementById('autoBtn').classList.toggle('active', mode === 'auto');
      document.getElementById('manualBtn').classList.toggle('active', mode === 'manual');
      document.getElementById('autoInfo').style.display = mode === 'auto' ? 'block' : 'none';
      document.getElementById('manualSection').classList.toggle('show', mode === 'manual');
    }

    pdfInput.addEventListener('change', () => {
      if (pdfInput.files.length > 0) {
        fileName.textContent = pdfInput.files[0].name;
      }
    });

    document.addEventListener('dragover', e => e.preventDefault());
    document.addEventListener('drop', e => e.preventDefault());

    ['dragover', 'dragenter'].forEach(evt => {
      dropZone.addEventListener(evt, e => { e.preventDefault(); e.stopPropagation(); dropZone.classList.add('dragover'); });
    });
    ['dragleave'].forEach(evt => {
      dropZone.addEventListener(evt, e => { e.preventDefault(); e.stopPropagation(); dropZone.classList.remove('dragover'); });
    });
    dropZone.addEventListener('drop', e => {
      e.preventDefault();
      e.stopPropagation();
      dropZone.classList.remove('dragover');
      const files = e.dataTransfer.files;
      if (files.length > 0 && files[0].name.toLowerCase().endsWith('.pdf')) {
        pdfInput.files = files;
        fileName.textContent = files[0].name;
      }
    });

    form.addEventListener('submit', async (e) => {
      e.preventDefault();
      if (!pdfInput.files.length) return;

      submitBtn.disabled = true;
      spinner.style.display = 'block';
      spinnerText.style.display = 'block';
      result.innerHTML = '';

      const formData = new FormData(form);
      try {
        const resp = await fetch('/upload', { method: 'POST', body: formData });
        if (resp.ok) {
          const data = await resp.json();
          let detected = '';
          if (data.footers_found && data.footers_found.length > 0) {
            detected = '<div class="detected"><strong>Detected footers:</strong><br>'
              + data.footers_found.map(f => '&bull; ' + f).join('<br>') + '</div>';
          }
          result.innerHTML = '<div class="result success">'
            + '<a href="/download/' + data.file_id + '/' + encodeURIComponent(data.download_name) + '">Download: ' + data.download_name + '</a>'
            + '<div class="info">Removed footer from ' + data.pages_cleaned + ' of ' + data.total_pages + ' pages</div>'
            + detected
            + '</div>';
        } else {
          const err = await resp.json();
          result.innerHTML = '<div class="result error">' + err.error + '</div>';
        }
      } catch (err) {
        result.innerHTML = '<div class="result error">Something went wrong. Please try again.</div>';
      }
      submitBtn.disabled = false;
      spinner.style.display = 'none';
      spinnerText.style.display = 'none';
    });
  </script>
</body>
</html>
'''


def normalize_text(text):
    """Strip numbers and whitespace to detect repeating patterns regardless of page numbers."""
    text = text.strip()
    # Replace sequences of digits (page numbers) with a placeholder
    text = re.sub(r'\d+', '#', text)
    # Collapse whitespace
    text = re.sub(r'\s+', ' ', text)
    return text


def detect_footers(doc):
    """
    Scan the PDF and find text blocks that repeat near the bottom of pages.
    Returns a dict: normalized_text -> list of (page_index, rect) tuples
    """
    total_pages = len(doc)
    if total_pages < 2:
        # For single page, treat bottom 15% as footer
        page = doc[0]
        h = page.rect.height
        threshold = h * 0.80
        footer_blocks = []
        for b in page.get_text("blocks"):
            x0, y0, x1, y1, text, block_no, block_type = b
            if block_type == 0 and y0 >= threshold and text.strip():
                footer_blocks.append((0, fitz.Rect(x0, y0, x1, y1), text.strip()))
        return footer_blocks, []

    # Pass 1: collect bottom-region text blocks from all pages
    bottom_texts = {}  # normalized -> [(page_idx, rect, original_text)]

    for i, page in enumerate(doc):
        h = page.rect.height
        # Bottom 20% of the page
        threshold = h * 0.80

        for b in page.get_text("blocks"):
            x0, y0, x1, y1, text, block_no, block_type = b
            if block_type != 0:  # skip image blocks
                continue
            text_clean = text.strip()
            if not text_clean:
                continue
            if y0 >= threshold:
                key = normalize_text(text_clean)
                if key not in bottom_texts:
                    bottom_texts[key] = []
                bottom_texts[key].append((i, fitz.Rect(x0 - 2, y0 - 2, x1 + 2, y1 + 2), text_clean))

    # Pass 2: identify footers - text that appears on at least 20% of pages (min 2 pages)
    min_pages = max(2, int(total_pages * 0.15))
    footer_entries = []  # all (page_idx, rect) to redact
    footer_labels = []   # human-readable descriptions of what was found

    for key, entries in bottom_texts.items():
        page_count = len(set(e[0] for e in entries))
        if page_count >= min_pages:
            footer_entries.extend(entries)
            # Use first occurrence as label, truncate if long
            sample = entries[0][2]
            if len(sample) > 80:
                sample = sample[:80] + "..."
            footer_labels.append(f"{sample} ({page_count} pages)")

    # Also detect standalone page numbers (just a number at the bottom)
    for key, entries in bottom_texts.items():
        if key.strip() == '#':
            page_count = len(set(e[0] for e in entries))
            if page_count >= min_pages:
                already = any(e in footer_entries for e in entries)
                if not already:
                    footer_entries.extend(entries)
                    footer_labels.append(f"Page numbers ({page_count} pages)")

    return footer_entries, footer_labels


def remove_footer_auto(input_path, output_path):
    """Auto-detect and remove footers."""
    doc = fitz.open(input_path)
    total_pages = len(doc)

    footer_entries, footer_labels = detect_footers(doc)

    # Group by page
    pages_to_redact = {}
    for page_idx, rect, text in footer_entries:
        if page_idx not in pages_to_redact:
            pages_to_redact[page_idx] = []
        pages_to_redact[page_idx].append(rect)

    # Apply redactions
    for page_idx, rects in pages_to_redact.items():
        page = doc[page_idx]
        for rect in rects:
            page.add_redact_annot(rect)
        page.apply_redactions()

    doc.save(output_path)
    doc.close()
    return len(pages_to_redact), total_pages, footer_labels


def remove_footer_manual(input_path, output_path, footer_text):
    """Remove footer by matching specific text."""
    doc = fitz.open(input_path)
    removed_count = 0
    total_pages = len(doc)

    for page in doc:
        blocks = page.get_text("blocks")
        found = False
        for b in blocks:
            x0, y0, x1, y1, text, block_no, block_type = b
            if footer_text.lower() in text.lower():
                rect = fitz.Rect(x0 - 5, y0 - 2, x1 + 5, y1 + 2)
                page.add_redact_annot(rect)
                found = True
        if found:
            page.apply_redactions()
            removed_count += 1

    doc.save(output_path)
    doc.close()
    return removed_count, total_pages


@app.route('/')
def index():
    return render_template_string(HTML)


@app.route('/upload', methods=['POST'])
def upload():
    if 'pdf' not in request.files:
        return {'error': 'No file uploaded'}, 400

    pdf = request.files['pdf']
    if pdf.filename == '' or not pdf.filename.lower().endswith('.pdf'):
        return {'error': 'Please upload a valid PDF file'}, 400

    mode = request.form.get('mode', 'auto')

    file_id = str(uuid.uuid4())[:8]
    original_name = os.path.splitext(pdf.filename)[0]
    upload_path = os.path.join(app.config['UPLOAD_FOLDER'], f'{file_id}.pdf')
    cleaned_name = f'{original_name}-cleaned.pdf'
    cleaned_path = os.path.join(app.config['CLEANED_FOLDER'], f'{file_id}.pdf')

    pdf.save(upload_path)

    try:
        if mode == 'manual':
            footer_text = request.form.get('footer_text', '').strip()
            if not footer_text:
                return {'error': 'Please enter the footer text to remove'}, 400
            removed, total = remove_footer_manual(upload_path, cleaned_path, footer_text)
            footers_found = [footer_text]
        else:
            removed, total, footers_found = remove_footer_auto(upload_path, cleaned_path)
    except Exception as e:
        return {'error': f'Error processing PDF: {str(e)}'}, 500
    finally:
        if os.path.exists(upload_path):
            os.remove(upload_path)

    if removed == 0 and mode == 'auto':
        return {
            'file_id': file_id,
            'download_name': cleaned_name,
            'pages_cleaned': 0,
            'total_pages': total,
            'footers_found': ['No repeating footers detected. Try Manual mode if you see a footer.']
        }

    return {
        'file_id': file_id,
        'download_name': cleaned_name,
        'pages_cleaned': removed,
        'total_pages': total,
        'footers_found': footers_found
    }


@app.route('/download/<file_id>/<download_name>')
def download(file_id, download_name):
    cleaned_path = os.path.join(app.config['CLEANED_FOLDER'], f'{file_id}.pdf')
    if not os.path.exists(cleaned_path):
        return 'File not found or expired', 404
    return send_file(cleaned_path, as_attachment=True, download_name=download_name)


if __name__ == '__main__':
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    os.makedirs(app.config['CLEANED_FOLDER'], exist_ok=True)
    app.run(host='0.0.0.0', port=8080)
