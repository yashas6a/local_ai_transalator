from nicegui import ui, events
import ollama
import asyncio
import io

# Setup State
class State:
    def __init__(self):
        #self.model = "gemma3:12b"  # Your 5060 Ti friendly model
        self.model = "qwen3:8b"

state = State()

# --- SHARED OLLAMA FUNCTION ---
async def get_ollama_response(prompt):
    loop = asyncio.get_event_loop()
    response = await loop.run_in_executor(None, lambda: ollama.generate(model=state.model, prompt=prompt))
    return response['response']

# --- UI LAYOUT ---
@ui.page('/')
def index():
    ui.colors(primary='#4285F4') # Gemini Blue
    
    with ui.header().classes('items-center justify-between'):
        ui.label('Local AI Studio').classes('text-2xl font-bold')
        with ui.row():
            ui.button('Chat', on_click=lambda: nav.set_value('Chat')).props('flat color=white')
            ui.button('Translate', on_click=lambda: nav.set_value('Translate')).props('flat color=white')
            ui.button('Batch Files', on_click=lambda: nav.set_value('Batch')).props('flat color=white')

    nav = ui.radio(['Chat', 'Translate', 'Batch'], value='Chat').classes('hidden')

    # --- CHAT & TRANSLATE PAGES (Omitted for brevity, keep your existing code here) ---
    # PAGE 1: CHAT (Gemini Style)
    with ui.column().bind_visibility_from(nav, 'value', value='Chat').classes('w-full max-w-3xl mx-auto p-4'):
        ui.markdown('### ✨ Chat with Gemma')
        chat_container = ui.column().classes('w-full border p-4 rounded-lg bg-gray-50 h-96 overflow-y-auto')
        
        with ui.row().classes('w-full'):
            user_input = ui.input(placeholder='Type your message...').classes('flex-grow')
            
            async def send_chat():
                msg = user_input.value
                user_input.value = ''
                with chat_container:
                    ui.markdown(f"**You:** {msg}")
                    response_area = ui.markdown("Thinking...")
                    answer = await get_ollama_response(msg)
                    response_area.set_content(f"**Gemma:** {answer}")

            ui.button('Send', on_click=send_chat)

    # PAGE 2: TRANSLATE (Google Style)
    with ui.column().bind_visibility_from(nav, 'value', value='Translate').classes('w-full p-8'):
        ui.markdown('### 🌐 Local Translate')
        with ui.row().classes('w-full gap-4'):
            src_area = ui.textarea('Input').classes('w-1/2 h-64 border p-2')
            res_area = ui.textarea('Translation').classes('w-1/2 h-64 border p-2 bg-blue-50').props('readonly')
        
        target_lang = ui.select(['English', 'Spanish', 'French', 'German', 'Japanese'], value='English').classes('w-32')
        
        async def run_translate():
            prompt = f"Translate to {target_lang.value}. Only return translation: {src_area.value}"
            res_area.value = "Translating..."
            answer = await get_ollama_response(prompt)
            res_area.value = answer

        ui.button('Translate Now', on_click=run_translate).classes('w-full bg-blue-600 text-white')

    # --- PAGE 3: BATCH FILE TRANSLATION ---
    with ui.column().bind_visibility_from(nav, 'value', value='Batch').classes('w-full p-8 max-w-4xl mx-auto'):
        ui.markdown('### 📂 Batch File Translator')
        ui.label('Upload a text file to translate the entire content.')
        
        target_lang_batch = ui.select(['English', 'Spanish', 'French', 'German', 'Japanese', 'Hindi'], value='English').classes('w-48')
        
        async def handle_upload(e: events.UploadEventArguments):
            # 1. Read the uploaded file content as text
            text_content = e.content.read().decode('utf-8')
            ui.notify(f"Processing {e.name}...")

            # 2. Send to Gemma on 5060 Ti
            prompt = f"Translate the following document to {target_lang_batch.value}. Maintain formatting. Only return the translation:\n\n{text_content}"
            translated_text = await get_ollama_response(prompt)

            # 3. Create a download button for the result
            with results_container:
                ui.label(f"Finished: {e.name}").classes('font-bold mt-4')
                ui.button('Download Translated File', 
                          on_click=lambda: ui.download(translated_text.encode('utf-8'), f"translated_{e.name}"))
                ui.notify("File ready for download!", color='positive')

        ui.upload(on_upload=handle_upload, auto_upload=True, label="Pick a .txt or .md file").props('accept=.txt,.md').classes('w-full')
        
        results_container = ui.column().classes('w-full p-4 bg-gray-100 rounded-lg mt-4')

ui.run(title="Local Gemini", port=8080, reload=False)