import os
import io
import re
import copy
import logging
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes
import docx
from docx.shared import Pt

# --- Configuration ---
BOT_TOKEN = "8127720127:AAFeFVi4a2ZXmY-osUz9HjreJT4ZCfe4mtc" # <-- PASTE YOUR REAL TOKEN
TABLES_PER_FILE = 30
DOWNLOAD_DIR = "downloads"

# --- Setup Logging ---
logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)
logger = logging.getLogger(__name__)

# =====================================================================
# === HELPER & LOGIC FUNCTIONS (Unchanged) ===
# =====================================================================

def parse_individual_question(block: str):
    block = block.strip()
    if not block:
        return None
    lines = [line.strip() for line in block.split('\n') if line.strip()]
    if len(lines) < 5:
        raise ValueError("Incomplete block: must have a question and four options.")
    correct_option_match = re.search(r'Correct Option:\s*(\S+)', block, re.IGNORECASE)
    if not correct_option_match:
        raise ValueError("Missing 'Correct Option: [id]' line.")
    correct_option_id = correct_option_match.group(1).lower()
    question_text = re.sub(r'^(?:\d+\.|Q\.)\s*', '', lines[0])
    parsed_options = []
    option_ids = ['a', 'b', 'c', 'd']
    for i in range(4):
        option_text = re.sub(r'^[a-zA-Z\d]+[\.\)]\s*', '', lines[i + 1])
        parsed_options.append({'id': option_ids[i], 'text': option_text})
    explanation_lines = [line for line in lines[5:] if 'Correct Option:' not in line]
    explanation_text = "\n".join(explanation_lines).strip()
    return {
        'question_text': question_text,
        'options': parsed_options,
        'correct_option_id': correct_option_id,
        'explanation_text': explanation_text
    }

def format_tables_in_doc(document: docx.Document):
    logger.info(f"Formatting fonts for {len(document.tables)} tables.")
    for table in document.tables:
        for row in table.rows:
            row_identifier = row.cells[0].text.strip()
            target_cell = row.cells[1] if len(row.cells) > 1 else None
            if not target_cell: continue
            font_size = None
            if row_identifier == 'Question': font_size = 14
            elif row_identifier in ['Option', 'Solution']: font_size = 12
            if font_size:
                for paragraph in target_cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(font_size)

def clone_table(table, new_doc):
    p = new_doc.add_paragraph()
    p._p.addnext(copy.deepcopy(table._tbl))
    new_doc.add_paragraph()

# =====================================================================
# === NEW & MODIFIED FUNCTIONS FOR THE FEATURE ===
# =====================================================================

async def create_docx_from_text(text_content: str) -> tuple[io.BytesIO | None, list[str]]:
    """
    NEW REFACTORED FUNCTION: Takes a string of text, processes it, creates a 
    formatted docx file, and returns it as a stream along with any errors.
    """
    question_blocks = re.split(r'\n\s*\n', text_content.strip())
    valid_questions, failed_blocks = [], []

    for i, block in enumerate(question_blocks):
        if not block.strip(): continue
        try:
            parsed_data = parse_individual_question(block)
            if parsed_data: valid_questions.append(parsed_data)
        except ValueError as e:
            failed_blocks.append(f"❗️<b>Error in Question #{i+1}:</b> {e}\n<pre>{block}</pre>")
    
    if not valid_questions:
        return None, failed_blocks

    try:
        new_doc = docx.Document()
        for q_data in valid_questions:
            table = new_doc.add_table(rows=0, cols=3, style='Table Grid')
            row_cells = table.add_row().cells; row_cells[0].text = 'Question'; row_cells[1].merge(row_cells[2]).text = q_data['question_text']
            row_cells = table.add_row().cells; row_cells[0].text = 'Type'; row_cells[1].merge(row_cells[2]).text = 'multiple_choice'
            for option in q_data['options']:
                row_cells = table.add_row().cells; row_cells[0].text = 'Option'; row_cells[1].text = option['text']
                row_cells[2].text = 'correct' if option['id'] == q_data['correct_option_id'] else 'incorrect'
            row_cells = table.add_row().cells; row_cells[0].text = 'Solution'; row_cells[1].merge(row_cells[2]).text = q_data['explanation_text']
            row_cells = table.add_row().cells; row_cells[0].text = 'Marks'; row_cells[1].text = '4'; row_cells[2].text = '1'
            new_doc.add_paragraph('')
        
        format_tables_in_doc(new_doc)

        output_stream = io.BytesIO()
        new_doc.save(output_stream)
        output_stream.seek(0)
        return output_stream, failed_blocks
    except Exception as e:
        logger.error(f"Error during DOCX creation from text: {e}")
        failed_blocks.append(f"❌ A critical error occurred while generating the document: {e}")
        return None, failed_blocks

# =====================================================================
# === TELEGRAM HANDLERS (MODIFIED) ===
# =====================================================================

async def start_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """MODIFIED: Sends a welcome message explaining the new, smarter functionality."""
    await update.message.reply_html(
        "👋 <b>Welcome to the All-in-One Docx Bot!</b>\n\n"
        "I can help you in a few ways:\n\n"
        "1️⃣ <b>Send Plain Text:</b> Paste your questions directly into the chat, and I'll create a formatted .docx file for you.\n\n"
        "2️⃣ <b>Send a .docx File:</b> I will automatically inspect your file.\n"
        "   - If it contains <i>plain text</i>, I'll convert it to a table-based document immediately.\n"
        "   - If it contains <i>tables</i>, I'll add it to a queue. Send more files with tables, then use <code>/convert</code> to merge them all!"
    )

async def handle_text_message(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """MODIFIED: This now uses the new refactored function."""
    user_text = update.message.text
    await update.message.reply_text("🔄 Processing your text...")

    output_stream, failed_blocks = await create_docx_from_text(user_text)

    if failed_blocks:
        await update.message.reply_html("Some parts of your text could not be processed:\n\n" + "\n\n".join(failed_blocks))

    if output_stream:
        await update.message.reply_document(document=output_stream, filename="formatted_from_text.docx")
        logger.info("Successfully sent docx created from plain text message.")

async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """
    HEAVILY MODIFIED: This function now inspects the docx file and decides which
    workflow to trigger (merge vs. text conversion).
    """
    if not update.message.document or not update.message.document.file_name.endswith('.docx'):
        await update.message.reply_text("⚠️ Please send only `.docx` files.")
        return

    await update.message.reply_text("Inspecting your .docx file...")
    doc_file = await update.message.document.get_file()
    
    in_memory_stream = io.BytesIO()
    await doc_file.download_to_memory(in_memory_stream)
    in_memory_stream.seek(0)

    try:
        document = docx.Document(in_memory_stream)

        # DECISION POINT: Does the document contain tables?
        if document.tables:
            # WORKFLOW 1: File has tables, save for merging.
            logger.info("Document has tables. Adding to merge queue.")
            user_id = update.message.from_user.id
            if 'files' not in context.user_data: context.user_data['files'] = []
            
            file_path = os.path.join(DOWNLOAD_DIR, f"{user_id}_{doc_file.file_id}.docx")
            with open(file_path, 'wb') as f:
                f.write(in_memory_stream.getbuffer())
            
            context.user_data['files'].append(file_path)
            await update.message.reply_html("✅ File contains tables and has been added to the queue. Send more files or use <code>/convert</code> to merge.")
        else:
            # WORKFLOW 2: File has NO tables, so process its text content now.
            logger.info("Document has no tables. Processing its text content.")
            full_text = "\n".join([p.text for p in document.paragraphs if p.text.strip()])

            if not full_text:
                await update.message.reply_text("This .docx file appears to be empty or contains no text.")
                return

            await update.message.reply_text("Found plain text. Attempting to convert it into a formatted document...")
            output_stream, failed_blocks = await create_docx_from_text(full_text)
            
            if failed_blocks:
                await update.message.reply_html("Could not process the text from your document:\n\n" + "\n\n".join(failed_blocks))
            if output_stream:
                await update.message.reply_document(document=output_stream, filename=f"formatted_{update.message.document.file_name}")
                logger.info("Successfully sent docx created from docx text content.")

    except Exception as e:
        logger.error(f"Error processing document file: {e}")
        await update.message.reply_text("❌ An error occurred while processing your .docx file. It might be corrupted.")

# --- Unchanged convert_command and main function ---
async def convert_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    user_id = update.message.from_user.id
    input_files = context.user_data.get('files', [])
    if not input_files:
        await update.message.reply_text("You haven't added any files with tables to the queue yet.")
        return

    await update.message.reply_text(f"🔄 Processing {len(input_files)} file(s) from the queue...")
    all_tables, output_files = [], []
    try:
        for file_path in input_files:
            doc = docx.Document(file_path)
            all_tables.extend(doc.tables)
        if not all_tables:
            await update.message.reply_text("ℹ️ No tables were found in the queued document(s).")
            return

        for i in range(0, len(all_tables), TABLES_PER_FILE):
            chunk = all_tables[i:i + TABLES_PER_FILE]
            new_doc = docx.Document()
            new_doc.add_heading(f"Merged Tables - Part {len(output_files) + 1}", 0)
            for table in chunk: clone_table(table, new_doc)
            format_tables_in_doc(new_doc)
            output_filename = os.path.join(DOWNLOAD_DIR, f"{user_id}_merged_part_{len(output_files) + 1}.docx")
            new_doc.save(output_filename)
            output_files.append(output_filename)

        await update.message.reply_text(f"✅ Conversion complete! Found {len(all_tables)} tables. Sending you {len(output_files)} new file(s)...")
        for output_file in output_files:
            with open(output_file, 'rb') as f:
                await context.bot.send_document(chat_id=update.effective_chat.id, document=f)
    except Exception as e:
        logger.error(f"Error during conversion for user {user_id}: {e}")
        await update.message.reply_text("❌ An error occurred during the conversion process.")
    finally:
        logger.info(f"Cleaning up files for user {user_id}")
        for file_path in input_files + output_files:
            if os.path.exists(file_path): os.remove(file_path)
        context.user_data['files'] = []

def main() -> None:
    if "YOUR_TELEGRAM_BOT_TOKEN_HERE" in BOT_TOKEN:
        print("!!! ERROR: Please paste your real Telegram bot token in the BOT_TOKEN variable. !!!")
        return
    if not os.path.exists(DOWNLOAD_DIR): os.makedirs(DOWNLOAD_DIR)

    application = Application.builder().token(BOT_TOKEN).build()
    application.add_handler(CommandHandler("start", start_command))
    application.add_handler(CommandHandler("convert", convert_command))
    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text_message))
    application.add_handler(MessageHandler(filters.Document.ALL, handle_document))

    print("Unified All-in-One Bot is running...")
    application.run_polling()

if __name__ == '__main__':
    main()
    
