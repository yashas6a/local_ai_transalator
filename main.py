import os, shutil, io, sys
import streamlit as st
from llama_cpp import Llama
import pytesseract
from PIL import Image
from pdf2image import convert_from_path
from docx import Document
import fitz
from typing import List, Optional
import tiktoken  # OpenAI tokenizer, works for most BPE-based models
import comtypes.client
#import win32com.client

###LANGUAGE
lang_code = "jpn"
//for pytessart ocr....feel free to add more supported languages here
SUPPORTED_LANGS = {
    "English (eng)": "eng",
    "Chinese (chi_sim)": "chi_sim",
    "Burmese (mya)": "mya",
    "Bur+Eng (mya_eng)": "mya+eng",
    "Russian (rus)": "rus",
    "Japanese (jpn)": "jpn",
    "Arabic (ara)" : "ara"
}

# === TESSARACT ===
def resource_path(relative_path):
    """Get absolute path to resource, works for PyInstaller."""
    try:
        base_path = sys._MEIPASS
        print(f"DEBUG: BASE_PATH = {base_path}")
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.normpath(os.path.join(base_path, relative_path))
# Set Tesseract path
try:
    pytesseract.pytesseract.tesseract_cmd = resource_path("tesseract_bundle/tesseract/tesseract.exe")
except Exception as e:
    print(f"ERROR: Failed to set Tesseract path: {str(e)}")

# Add Poppler to PATH for pdf2image
try:
    os.environ["PATH"] += os.pathsep + resource_path("tesseract_bundle/poppler/Library/bin")
except Exception as e:
    print(f"ERROR: Failed to set Poppler path: {str(e)}")

# Set TESSDATA_PREFIX for language files
try:
    os.environ["TESSDATA_PREFIX"] = resource_path("tesseract_bundle/tessdata")
except Exception as e:
    print(f"ERROR: Failed to set TESSDATA_PREFIX: {str(e)}")

# ---------------------- Configuration ----------------------
MODEL_PATH = "model/gemma-3-12b-it-Q3_K_L.gguf"  # 🔄 Update to your file
#MODEL_PATH = "C:\\Users\\DOMS\\Documents\\LLAMA\\models\\translate\\gemma-3-12b-it-Q3_K_L.gguf"
N_GPU_LAYERS = 0  # -1 = all on GPU, 0 = CPU‑only
N_CTX = 6000
MAX_TOKENS_PER_CHUNK = 2000  # Adjust based on your model's context limit

# ----------------------- Helpers ---------------------------
@st.cache_resource(show_spinner="Loading Gemma ⏳ …")
def load_model():
    return Llama(model_path=MODEL_PATH, n_gpu_layers=N_GPU_LAYERS, n_ctx=N_CTX, verbose=False)

llm = load_model()

# === Setup Tokenizer ===
#enc = tiktoken.get_encoding("cl100k_base")  # Compatible with LLaMA-style models
enc = tiktoken.get_encoding("cl100k_base")

# --------------------- Streamlit UI ------------------------
st.set_page_config(page_title="Gemma Translator", page_icon="🦙", layout="wide")

#st.title("🦙 Gemma‑powered Translator")
#st.header("🦙 Gemma‑powered Translator")

# Top‑level tabs similar to Google Translate
#mode = st.tabs(["📝 Text", "📄 Document", "🔍🖼️ OCR", "💬 Chat"])


# Function to programmatically switch tabs
if "mode" not in st.session_state:
    st.session_state.mode = "📝 Text"

if "input_text" not in st.session_state:
    st.session_state.input_text = ""

def switch_mode(new_mode, content: str):
    st.session_state.input_text = content
    st.session_state.mode = new_mode
    st.rerun()
    
selected_mode = st.radio(
    "🦙 Gemma‑powered Translator",
    ["📝 Text", "📄 Document", "🔍🖼️ OCR", "💬 Chat"],
    index=["📝 Text", "📄 Document", "🔍🖼️ OCR", "💬 Chat"].index(st.session_state.mode),
    horizontal=True,
    key="mode_radio",
    disabled=st.session_state.get("awaiting_reply", False)  # disable switching during stream
)

# Update session_state based on user selection
if selected_mode != st.session_state.mode and not st.session_state.get("awaiting_reply", False):
    st.session_state.mode = selected_mode
    st.rerun()

mode = st.session_state.mode


#----------------------- CHUNKING ----------------------------
def split_into_token_chunks(
    text: str,
    max_tokens: int,
    chars_per_token: int = 4
) -> List[str]:
    """
    Splits text into chunks of <= max_tokens tokens, ending on a sentence or paragraph boundary.
    Supports both Western (.?!), CJK (。？！), and paragraph (\n\n) boundaries.

    If no tokenizer is supplied, uses regex for sentence boundaries and estimates token counts.
    - enc: a tokenizer with .encode() and .decode(), like tiktoken or Hugging Face's tokenizer.
    """
    # Define sentence and paragraph boundaries
    sentence_boundaries = [
        ". ", "? ", "! ", "။ ", "。", "？", "！"
    ]
    paragraph_boundary = "\n\n"

    # --- Tokenization-aware path ---
    if enc:
        tokens = enc.encode(text)
        n = len(tokens)
        chunks = []
        start = 0

        while start < n:
            # Initial candidate chunk
            end = min(start + max_tokens, n)
            chunk_tokens = tokens[start:end]
            chunk_text = enc.decode(chunk_tokens)

            # Find the last sentence boundary within chunk
            last_boundary = -1
            last_mark_len = 0

            # Search for each boundary mark and record the furthest (rightmost) position
            for mark in sentence_boundaries + [paragraph_boundary]:
                pos = chunk_text.rfind(mark)
                if pos > last_boundary:
                    last_boundary = pos
                    last_mark_len = len(mark)

            # If a boundary exists and chunk isn't too small, adjust chunk to end at boundary
            if last_boundary != -1 and (end - start) > 30:
                chunk_text = chunk_text[:last_boundary + last_mark_len]
                chunk_tokens = enc.encode(chunk_text)
                end = start + len(chunk_tokens)
                chunk_text = enc.decode(tokens[start:end])  # Re-decode to handle edge cases

            chunks.append(chunk_text.strip())
            start = end

        return [c for c in chunks if c.strip()]

    # --- Regex fallback path (no tokenizer) ---
    import re
    # Split only after single-character sentence-ending punctuation (Western & CJK)
    # We avoid (?<=\n\n) since that's variable-width, handle paragraphs below
    # Uses Unicode-aware full stops
    sentence_pattern = re.compile(r'(?<=[\.\?\!။。？！])\s+')
    segments = sentence_pattern.split(text)

    # Now handle paragraph breaks within each segment
    new_segments = []
    for seg in segments:
        if paragraph_boundary in seg:
            parts = seg.split(paragraph_boundary)
            # Add paragraph break back for all but the last segment
            for i, part in enumerate(parts):
                # Restore paragraph break except for last one
                if i < len(parts) - 1:
                    new_segments.append(part + paragraph_boundary)
                else:
                    new_segments.append(part)
        else:
            new_segments.append(seg)

    # Join segments into chunks by estimated chars per token
    chunks = []
    buffer = ""
    max_chars = max_tokens * chars_per_token

    for seg in new_segments:
        seg = seg.strip()
        # If segment pushes buffer over max, push buffer as a chunk
        if len(buffer) + len(seg) > max_chars and buffer:
            chunks.append(buffer.strip())
            buffer = seg
        else:
            if buffer:
                buffer += " " + seg
            else:
                buffer = seg
    if buffer.strip():
        chunks.append(buffer.strip())

    return [c for c in chunks if c.strip()]
#----------------------------------------------------
def ocr_image_bytes(image_bytes: bytes, lang_code: str) -> str:
    with st.spinner("Running Tesseract 🧐 …"):
        """Run Tesseract OCR on raw image bytes and return extracted text."""
        with Image.open(io.BytesIO(image_bytes)) as img:
            text = pytesseract.image_to_string(img, lang=lang_code)
        return text.strip()

#@st.cache_data(show_spinner="OCR‑ing PDF pages …", hash_funcs={bytes:lambda _: None})
def pdf_to_text(file_path: str, lang_code: str) -> str:
    with st.spinner("OCR‑ing PDF pages …"):
        """Run Tesseract on every page obtained via **convert_from_path** (no pdftoimage CLI)."""
        images = convert_from_path(file_path)  # <- your requested helper
        text = ""
        for i, image in enumerate(images):
            try:
                page_text = pytesseract.image_to_string(image, lang=lang_code)
                text += f"--- Page {i+1} ---\n{page_text}\n"
            except Exception as e:
                text += f"[Error processing page {i+1}: {e}]\n"
                continue
        return text.strip()

#@st.cache_data(show_spinner="Scanning PDF pages for text/image …", hash_funcs={bytes:lambda _: None})
def extract_all_from_pdf_n(pdf_path: str ,lang_code: str) -> str:
    with st.spinner("Scanning PDF pages for text/image …"):
        """
        Extracts text from a PDF file using both direct text extraction and OCR fallback.
        """
        images = convert_from_path(pdf_path)
        combined_text = []
        doc = fitz.open(pdf_path) 
        i=0
        # --- Extract all text ---
        for page in doc:        
            text = page.get_text()
            if text.strip():
                #print(f"Text on page {i}")
                combined_text.append(f"Page {i+1} (Text)\n{text}")
            else:
                # Use OCR if no text found
                #print(f"OCR on page {i+1}")
                image = images[i]
                ocr_text = pytesseract.image_to_string(image, lang=lang_code)
                combined_text.append(f"Page {i+1} (OCR)\n{ocr_text}")

            combined_text.append("-" * 80)  # Page separator
            i+=1
        return "\n".join(combined_text)

# === DOC ===
#@st.cache_data(show_spinner="Converting DOC to DOCX…", hash_funcs={bytes:lambda _: None})
def convert_doc_to_docx(filepath):
    comtypes.CoInitialize()
    if not os.path.isfile(filepath):
        raise FileNotFoundError(f"File not found: {filepath}")
    else :        
        # Initialize Word application
        word = comtypes.client.CreateObject("Word.Application")
        word.Visible = False

        # Open the .doc file
        doc = word.Documents.Open(os.getcwd() +  "\\" + filepath)

        # Save as .docx
        doc.SaveAs(os.getcwd() + "\\temp_upload.docx", FileFormat=16)  # 16 corresponds to wdFormatXMLDocument (.docx)
        doc.Close()
        word.Quit()
        comtypes.CoUninitialize()
        #print(f"Conversion successful: {filepath}x")
    

#@st.cache_data(show_spinner="Scanning DOCX pages for text…", hash_funcs={bytes:lambda _: None})
def extract_text_from_docx(filepath: str) -> str:
    with st.spinner("Scanning DOCX pages for text …"):
        doc = Document(filepath)
        text = '\n'.join([p.text for p in doc.paragraphs])
        table_text = "\n"

        # Translate individual table cells
        i=0
        for table in doc.tables:
            i += 1
            table_text += f"table_{i}:\n"
            for row in table.rows:
                for cell in row.cells:
                    tt = cell.text.strip()
                    if tt:
                        table_text += "\t"+tt
                table_text += "\n"
        
        #translate_content_savetxt(text.strip() + table_text, filepath)
        return text.strip() + table_text

#@st.cache_data(show_spinner="Scanning DOCX pages for images…", hash_funcs={bytes:lambda _: None})
def extract_ocr_from_docx(filepath: str, lang_code: str) -> str:
    with st.spinner("Scanning DOCX pages for image …"):
        doc = Document(filepath)
        rels = doc.part.rels
        
        text = ""
        # --- Process OCR from embedded images ---
        for rel in rels.values():
            if "image" in rel.target_ref:
                img_data = rel.target_part.blob
                ocr_text = ocr_image_bytes(img_data, lang_code)
                text += ocr_text
        return text

# --------------------- TEXT MODE ---------------------------
#with mode[0]:
if st.session_state.mode == "📝 Text":    
    left, right = st.columns(2, gap="large")

    with left:
        st.subheader("Input 📥")
        input_text = st.text_area(" ", height=350, placeholder="Enter text to be translated here...!", key="input_text")
        submit = st.button("Translate 🚀", use_container_width=True, disabled=not input_text.strip())

    with right:
        st.subheader("Output 📤")
        st.markdown("")
        output_placeholder = st.empty()
        # Prepare placeholder; text_area will be rendered on first update to avoid duplicate keys
        if "translation" not in st.session_state:
            st.session_state.translation = ""
        output_placeholder.markdown(
            f"""
            <div style="height: 400px; min-height: 350px; overflow-y: auto; border: 1.5px solid #ccc;
                        border-radius: 8px; padding: 2px; background-color: #f9f9f9;
                        font-family: monospace; white-space: pre-wrap;">
                {st.session_state.translation}
            </div>
            """,
            unsafe_allow_html=True
        )
    
    with st.spinner("Gemma is translating… 💭"):
        if submit:
            st.session_state.awaiting_reply = True
            chunks = split_into_token_chunks(input_text.strip(), MAX_TOKENS_PER_CHUNK)
            translated_chunks = []
            translation = ""
            for idx, chunk_single in enumerate(chunks):
                #print(f"\nTranslating chunk_single {idx+1}/{len(chunks)} : \n{chunk_single.strip()}")
                message = st.success(f"Translating chunk {idx+1}/{len(chunks)}")                
                system_prompt = (
                    "You are a helpful assistant that translates any text to English. "
                    "Answer with only the translation—no explanations."
                )
                messages = [
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": chunk_single.strip()},
                ]

                stream = llm.create_chat_completion(messages=messages, 
                    max_tokens=4096,
                    temperature=0.6,
                    top_p=0.9, 
                    stream=True)
                                
                for chunk in stream:
                    token = chunk["choices"][0]["delta"].get("content", "")
                    translation += token
                    st.session_state.translation = translation
                # Display final result once, outside the loop
                #right.text_area("",
                 #   value=st.session_state.translation,
                #    height=350,
                #    disabled=False,
                #    key=f"output_box{idx}"
                #)
                #output_box.code(translation, language="markdown")
                output_placeholder.markdown(
                f"""
                <div style="height: 400px; min-height: 350px; overflow-y: auto; border: 1.5px solid #ccc;
                            border-radius: 8px; padding: 2px; background-color: #f9f9f9;
                            font-family: monospace; white-space: pre-wrap;">
                    {st.session_state.translation}
                </div>
                """,
                unsafe_allow_html=True
            )
                message.empty()
                translation += "\n"
            st.session_state.awaiting_reply = False

# ------------------- DOCUMENT MODE -------------------------
#with mode[1]:
elif st.session_state.mode == "📄 Document":
    st.subheader("Translate full documents 📄 → 📝")
    left, right = st.columns(2, gap="large")
    text = ""
    ocr = ""
    with left:
        st.markdown("#### Upload Docx or PDF")
        file_1 = st.file_uploader("Upload document", type=["pdf", "docx", "doc", "pptx", "xlsx"],  key="file_upload")
        lang_label_s = st.selectbox("Scan language", list(SUPPORTED_LANGS.keys()), index=0, key="scan_lang")
        do_scan= st.button("Scan Document 🧐", disabled=file_1 is None, key="scan_button")
    with right:
        st.markdown("#### Scanned Text Contents")
        st.text_area("Scan output", value=st.session_state.get("scan_text", ""), height=300, disabled=False, key="scan_output")
        if st.session_state.get("scan_text"):
            copy_btn = st.button("Copy to Clipboard", key="scan_copy_btn")
        else:
            copy_btn = False
        if do_scan and file_1 is not None:
            lang_code_s = SUPPORTED_LANGS[lang_label_s]
            
            if file_1.type == "application/pdf" or os.path.splitext(file_1.name)[1].lower() == ".pdf":
                # Save to a temporary file because convert_from_path works with file paths
                with st.spinner("Running PDF pages …"):
                    with open(os.getcwd() + "\\temp_upload.pdf", "wb") as f:
                        f.write(file_1.getbuffer())
                    text = extract_all_from_pdf_n("temp_upload.pdf", lang_code_s)
                    #ocr = extract_ocr_from_pdf("temp_upload.pdf", lang_code_s)
                    os.remove("temp_upload.pdf")

            elif os.path.splitext(file_1.name)[1].lower() == ".doc":                
                with open(os.getcwd() + "\\temp_upload.doc", "wb") as f:
                    f.write(file_1.getbuffer())
                convert_doc_to_docx(os.getcwd() + "\\temp_upload.doc")
                text = extract_text_from_docx(os.getcwd() + "\\temp_upload.docx")
                ocr = "\n\n[IMAGE_OCR]\n"+ extract_ocr_from_docx(os.getcwd() + "\\temp_upload.docx", lang_code_s)
                if os.path.isfile(os.getcwd() + "\\temp_upload.doc"):
                    os.remove(os.getcwd() + "\\temp_upload.doc")
                if os.path.isfile(os.getcwd() + "\\temp_upload.docx"):
                    os.remove(os.getcwd() + "\\temp_upload.docx") 
                    
            elif os.path.splitext(file_1.name)[1].lower() == ".docx":
                with st.spinner("Scanning docx pages…"):
                    with open("temp_upload.docx", "wb") as f:
                        f.write(file_1.getbuffer())
                    text = extract_text_from_docx("temp_upload.docx")
                    ocr = "\n\n[IMAGE_OCR]\n"+ extract_ocr_from_docx("temp_upload.docx", lang_code_s)
                    os.remove("temp_upload.docx")                    
            st.session_state.scan_text = text + ocr
            st.rerun()

            
        
        if copy_btn:
            switch_mode("📝 Text", st.session_state.scan_text)
            #st.session_state.input_text = st.session_state.ocr_text
            st.session_state.mode = "📝 Text"      
            st.rerun()
# ----------------------- OCR MODE --------------------------
#with mode[2]:
elif st.session_state.mode == "🔍🖼️ OCR":
    text = ""
    st.subheader("OCR 🔍 ✨")

    ocr_col, trans_col = st.columns(2, gap="large")

    with ocr_col:
        st.markdown("#### Upload Image or PDF")
        file = st.file_uploader("Choose an image or PDF", type=["png", "jpg", "jpeg", "pdf"], key="ocr_upload")
        lang_label = st.selectbox("OCR language", list(SUPPORTED_LANGS.keys()), index=0, key="ocr_lang")
        do_ocr = st.button("Run OCR 🧐", disabled=file is None, key="ocr_button")

    with trans_col:
        st.markdown("#### Extracted Text & Translation")
        st.text_area("OCR output", value=st.session_state.get("ocr_text", ""), height=300, disabled=False, key="ocr_output")
        if st.session_state.get("ocr_text"):
            copy_ocr = st.button("Copy to Clipboard", key="ocr_copy_btn")
        else:
            copy_ocr = False
        #st.text_area("English", value=st.session_state.get("ocr_translation", ""), height=200, disabled=False, key="ocr_translation_box")

    if do_ocr and file is not None:
        lang_code = SUPPORTED_LANGS[lang_label]
        if file.type == "application/pdf" or os.path.splitext(file.name)[1].lower() == ".pdf":
            # Save to a temporary file because convert_from_path works with file paths
            with st.spinner("Running OCR on PDF pages …"):
                with open("temp_upload.pdf", "wb") as f:
                    f.write(file.getbuffer())
                text += pdf_to_text("temp_upload.pdf", lang_code)
                os.remove("temp_upload.pdf")
        else:
            with st.spinner("Running OCR…"):
                text = ocr_image_bytes(file.getvalue(), lang_code)
        st.session_state.ocr_text = text
        st.rerun()

    if copy_ocr:
        switch_mode("📝 Text", st.session_state.ocr_text)
        #st.session_state.input_text = st.session_state.ocr_text
        st.session_state.mode = "📝 Text"            
        st.rerun()

# ----------------------- CHAT MODE -------------------------
#with mode[3]:
elif st.session_state.mode == "💬 Chat":
    st.subheader("Chat with Gemma 💬")

    # ------------------------------
    # Session State Initialization
    # ------------------------------
    st.session_state.setdefault("chat_history", [])
    st.session_state.setdefault("awaiting_reply", False)
    st.session_state.setdefault("clear_input", False)
    st.session_state.setdefault("chat_input", "")
    st.session_state.setdefault("partial_reply", "")

    # ------------------------------
    # Render Chat History
    # ------------------------------
    for message in st.session_state.chat_history:
        role = "You" if message["role"] == "user" else "Gemma"
        st.markdown(f"**{role}:** {message['content']}")

    # Placeholder for streaming assistant reply
    response_placeholder = st.empty()

    # ------------------------------
    # Clear chat input if flagged
    # ------------------------------
    if st.session_state.clear_input:
        st.session_state.chat_input = ""
        st.session_state.clear_input = False

    # ------------------------------
    # User Input Box
    # ------------------------------
    user_msg = st.text_input("Your message", key="chat_input")
    send_disabled = not user_msg.strip() or st.session_state.awaiting_reply
    send_btn = st.button("Send", disabled=send_disabled)

    # ------------------------------
    # On Send
    # ------------------------------
    if send_btn:
        st.session_state.awaiting_reply = True
        st.session_state.chat_history.append({"role": "user", "content": user_msg})
        st.session_state.clear_input = True
        st.session_state.partial_reply = ""

        # Compose messages with system prompt
        messages = [{"role": "system", "content": "You are a helpful assistant that translates text to English. Answer only with the translation."}]
        messages.extend(st.session_state.chat_history)

        # ------------------------------
        # Call LLM (Streamed)
        # ------------------------------
        with st.spinner("Gemma is thinking… 💭"):
            stream = llm.create_chat_completion(
                messages=messages,
                max_tokens=2048,
                temperature=0.6,
                top_p=0.9,
                stream=True
            )

            for chunk in stream:
                token = chunk["choices"][0]["delta"].get("content", "")
                st.session_state.partial_reply += token

                # Update the placeholder live
                response_placeholder.markdown(f"**Gemma:** {st.session_state.partial_reply}")

        # Add the complete assistant reply to history
        st.session_state.chat_history.append({
            "role": "assistant",
            "content": st.session_state.partial_reply
        })

        # Reset
        st.session_state.awaiting_reply = False
        st.session_state.partial_reply = ""
        st.rerun()


# ---------------------- Footer ------------------------------
with st.expander("ℹ️  How this works?"):
    st.markdown("Runs locally on downloaded gemma3 model.")
