import os
import logging
import io
import copy
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes
import docx
from docx.shared import Pt
from docx.oxml.shared import OxmlElement, qn

# --- Configuration ---
# PASTE YOUR TELEGRAM BOT TOKEN HERE
BOT_TOKEN = "8127720127:AAFeFVi4a2ZXmY-osUz9HjreJT4ZCfe4mtc"
TABLES_PER_FILE = 30
DOWNLOAD_DIR = "downloads"
VALID_ROW_IDENTIFIERS = ['Question', 'Option', 'Solution']

# --- Setup Logging ---
logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)
logger = logging.getLogger(__name__)


# ==============================================================================
# === FEATURE 1: FONT MODIFICATION LOGIC ===
# ==============================================================================

def set_font_size_for_cell(cell, size_in_pt):
    """Iterates through paragraphs and runs in a cell to set the font size."""
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            run.font.size = Pt(size_in_pt)

def apply_font_modifications_to_file(file_path: str) -> bool:
    """
    Opens a docx file from a path, modifies table font sizes based on rules,
    and saves the changes back to the same file.
    """
    try:
        document = docx.Document(file_path)
        logger.info(f"Applying font modifications to: {os.path.basename(file_path)}")

        for table in document.tables:
            for row in table.rows:
                row_identifier = row.cells[0].text.strip()
                if row_identifier in VALID_ROW_IDENTIFIERS:
                    if row_identifier == 'Question':
                        set_font_size_for_cell(row.cells[1], 14)
                    else: # Option and Solution
                        set_font_size_for_cell(row.cells[1], 12)

        document.save(file_path)
        logger.info(f"Successfully saved font modifications for: {os.path.basename(file_path)}")
        return True

    except Exception as e:
        logger.error(f"Error modifying DOCX file {file_path}: {e}")
        return False


# ==============================================================================
# === FEATURE 2: ACCURATE TABLE CLONING LOGIC ===
# ==============================================================================

def clone_table(table, new_doc):
    """
    Clones a table by copying its underlying XML element to preserve all formatting.
    """
    p = new_doc.add_paragraph()
    tbl_xml = table._tbl
    new_tbl_xml = copy.deepcopy(tbl_xml)
    p._p.addnext(new_tbl_xml)
    new_doc.add_paragraph()


# ==============================================================================
# === TELEGRAM BOT HANDLERS ===
# ==============================================================================

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Sends a welcome message explaining all bot features, including error handling."""
    await update.message.reply_html(
        "👋 **Welcome to the Advanced DOCX Processor Bot!**\n\n"
        "I perform three main tasks:\n"
        "1️⃣ **Validate & Report Errors**: I check each table. If a row doesn't start with 'Question', 'Option', or 'Solution', I will report it back to you. Tables with 4 or more such errors will be skipped entirely.\n"
        "2️⃣ **Format Fonts**: I correctly format the font sizes for all valid rows.\n"
        "3️⃣ **Combine Tables**: I extract all valid tables and combine them into new documents.\n\n"
        "<b>How to use:</b>\n"
        "1. Send me one or more <code>.docx</code> files.\n"
        "2. Use the <code>/process</code> command to begin."
    )

async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Handles receiving a document, saves it locally for processing."""
    message = update.message
    if not message.document or not message.document.file_name.endswith('.docx'):
        await message.reply_text("⚠️ Please send only `.docx` files.")
        return

    user_id = message.from_user.id
    if 'files' not in context.user_data:
        context.user_data['files'] = []

    try:
        file = await message.document.get_file()
        file_path = os.path.join(DOWNLOAD_DIR, f"{user_id}_{message.document.file_name}")
        await file.download_to_drive(file_path)
        context.user_data['files'].append(file_path)
        logger.info(f"User {user_id} uploaded file: {file_path}")

        await message.reply_text(
            "✅ File received.\n\n"
            "You can send another file or use <b>/process</b> to validate, format, and combine everything.",
            parse_mode='HTML'
        )
    except Exception as e:
        logger.error(f"Error handling document for user {user_id}: {e}")
        await message.reply_text("An error occurred while receiving your file. Please try again.")

async def process(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """
    Processes files: validates tables, reports errors, applies fonts, and extracts valid tables.
    """
    user_id = update.message.from_user.id
    if 'files' not in context.user_data or not context.user_data['files']:
        await update.message.reply_text("You haven't sent any files. Please send a `.docx` file first.")
        return

    input_files = context.user_data['files']
    
    await update.message.reply_text(f"🔄 Starting processing for {len(input_files)} file(s)...")

    # --- STEP 1: FONT MODIFICATION (on all files first) ---
    for file_path in input_files:
        if not apply_font_modifications_to_file(file_path):
            await update.message.reply_text(f"❌ Critical error modifying fonts in {os.path.basename(file_path)}. Aborting.")
            return

    # --- STEP 2: VALIDATION AND TABLE GATHERING ---
    valid_tables_to_clone = []
    rejected_table_count = 0
    unidentified_row_details = []
    
    try:
        for file_path in input_files:
            logger.info(f"Validating tables in: {file_path}")
            doc = docx.Document(file_path)

            for table in doc.tables:
                unidentified_rows_in_table = 0
                
                # First, validate the entire table
                for row in table.rows:
                    identifier = row.cells[0].text.strip()
                    if identifier not in VALID_ROW_IDENTIFIERS:
                        unidentified_rows_in_table += 1
                        # Collect details of the unidentified row for reporting
                        row_text = f"'{row.cells[0].text.strip()}' | '{row.cells[1].text.strip()}'"
                        unidentified_row_details.append(f"- In file `{os.path.basename(file_path)}`: {row_text}")

                # Now, decide whether to reject or accept the table
                if unidentified_rows_in_table >= 4:
                    rejected_table_count += 1
                else:
                    valid_tables_to_clone.append(table)
        
        # --- STEP 3: REPORTING ERRORS TO USER ---
        if unidentified_row_details:
            error_message = "⚠️ Found rows with unidentified identifiers:\n" + "\n".join(unidentified_row_details)
            await update.message.reply_text(error_message, parse_mode='Markdown')
        
        if rejected_table_count > 0:
            await update.message.reply_text(f"‼️ Rejected *{rejected_table_count} table(s) due to having 4 or more unidentified rows.", parse_mode='Markdown')

        if not valid_tables_to_clone:
            await update.message.reply_text("ℹ️ No valid tables were found to process after validation.")
            return

        # --- STEP 4: CREATE NEW DOCS FROM VALID TABLES ---
        await update.message.reply_text(f"✅ Validation complete. Now creating new documents from {len(valid_tables_to_clone)} valid tables...")
        output_files = []
        file_counter = 1
        for i in range(0, len(valid_tables_to_clone), TABLES_PER_FILE):
            chunk = valid_tables_to_clone[i:i + TABLES_PER_FILE]
            
            new_doc = docx.Document()
            new_doc.add_heading(f"Processed Tables - Part {file_counter}", level=1)
            new_doc.add_paragraph(f"This document contains {len(chunk)} of the {len(valid_tables_to_clone)} total valid tables.")
            
            for table in chunk:
                clone_table(table, new_doc)

            output_filename = os.path.join(DOWNLOAD_DIR, f"{user_id}_output_part_{file_counter}.docx")
            new_doc.save(output_filename)
            output_files.append(output_filename)
            file_counter += 1
            
        await update.message.reply_text(f"✅ Processing complete! Sending you {len(output_files)} new file(s)...")
        for output_file in output_files:
            with open(output_file, 'rb') as f:
                await context.bot.send_document(chat_id=update.effective_chat.id, document=f)

    except Exception as e:
        logger.error(f"Error during main processing for user {user_id}: {e}")
        await update.message.reply_text("❌ A critical error occurred during the main processing stage.")
    
    finally:
        # --- FINAL STEP: CLEANUP ---
        logger.info(f"Cleaning up all temp files for user {user_id}")
        all_temp_files = input_files + locals().get('output_files', [])
        for file_path in all_temp_files:
            if os.path.exists(file_path):
                os.remove(file_path)
        context.user_data['files'] = []

# ==============================================================================
# --- MAIN BOT EXECUTION ---
# ==============================================================================

def main() -> None:
    """Start the bot."""
    if BOT_TOKEN == "YOUR_TELEGRAM_BOT_TOKEN":
        print("!!! ERROR: Please replace 'YOUR_TELEGRAM_BOT_TOKEN' with your actual bot token. !!!")
        return

    if not os.path.exists(DOWNLOAD_DIR):
        os.makedirs(DOWNLOAD_DIR)

    application = Application.builder().token(BOT_TOKEN).build()
    
    application.add_handler(CommandHandler("start", start))
    application.add_handler(CommandHandler("process", process))
    application.add_handler(MessageHandler(filters.Document.ALL, handle_document))

    print("Advanced DOCX Processor Bot is running...")
    application.run_polling()

if __name__ == '__main__':
    main()
