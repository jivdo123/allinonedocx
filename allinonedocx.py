import os
import logging
import copy
import re
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes
import docx
from docx.shared import Pt

# --- Configuration ---
BOT_TOKEN = "8127720127:AAFeFVi4a2ZXmY-osUz9HjreJT4ZCfe4mtc"
TABLES_PER_FILE = 30
DOWNLOAD_DIR = "downloads"

# --- Setup Logging ---
logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)
logger = logging.getLogger(__name__)


# ==============================================================================
# === GENERAL HELPER FUNCTIONS ===
# ==============================================================================

def set_font_size_for_cell(cell, size_in_pt):
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            run.font.size = Pt(size_in_pt)

def apply_font_modifications_to_table(table):
    for row in table.rows:
        identifier = row.cells[0].text.strip()
        if identifier.lower() == 'question':
            set_font_size_for_cell(row.cells[1], 14)
        else:
            set_font_size_for_cell(row.cells[1], 12)

def clone_table(table, new_doc):
    p = new_doc.add_paragraph()
    tbl_xml = table._tbl
    new_tbl_xml = copy.deepcopy(tbl_xml)
    p._p.addnext(new_tbl_xml)
    new_doc.add_paragraph()


# ==============================================================================
# === WORKFLOW 1: PROCESSING FILES WITH EXISTING TABLES (Corrected) ===
# ==============================================================================

def process_table_document(doc):
    """
    Validates tables based on the specific cell check for 'incorrect'.
    """
    valid_tables = []
    rejected_count = 0
    
    for table in doc.tables:
        is_rejected = False
        # Rule: Check 3rd, 4th, 5th, and 6th rows (if they exist)
        if len(table.rows) >= 6:
            incorrect_markers = 0
            # Rows to check are at index 2, 3, 4, 5
            rows_to_check = table.rows[2:6]
            for row in rows_to_check:
                # Check the second column (index 1)
                if len(row.cells) > 1 and 'incorrect' in row.cells[1].text.lower():
                    incorrect_markers += 1
            
            # Reject if all four specific cells contain 'incorrect'
            if incorrect_markers == 4:
                is_rejected = True
        
        if is_rejected:
            rejected_count += 1
        else:
            valid_tables.append(table)
            
    report = f"Rejected {rejected_count} table(s) due to 4 'incorrect' markers in specified cells." if rejected_count > 0 else ""
    return valid_tables, report


# ==============================================================================
# === WORKFLOW 2: PROCESSING FILES WITH PLAIN TEXT (Corrected) ===
# ==============================================================================

def process_text_document(doc):
    """
    Parses plain text, creating tables. The 'Explanation' field is now optional.
    """
    newly_created_tables = []
    error_report_lines = []
    
    temp_doc = docx.Document()
    paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    
    current_block = {}
    line_number = 0

    def create_table_from_block(block):
        # Helper to create a table; ensures 'explanation' is handled if missing.
        is_valid = all(k in block for k in ['question', 'options', 'correct']) and len(block['options']) == 4
        if not is_valid:
            error_report_lines.append(f"Skipped malformed block starting with: '{block.get('question', 'Unknown')[:30]}...'")
            return None

        table = temp_doc.add_table(rows=7, cols=2)
        table.cell(0, 0).text = "Question"
        table.cell(0, 1).text = block['question']
        for i, opt in enumerate(block['options']):
            table.cell(i + 1, 0).text = f"Option {chr(97 + i)}"
            table.cell(i + 1, 1).text = opt
        table.cell(5, 0).text = "Correct Option"
        table.cell(5, 1).text = block['correct']
        table.cell(6, 0).text = "Explanation"
        table.cell(6, 1).text = block.get('explanation', '') # Use .get() for optional key
        return table

    while line_number < len(paragraphs):
        line = paragraphs[line_number]
        
        match_q = re.match(r'^(?:Q\.|(?:\d{1,3}\.))\s*(.*)', line, re.IGNORECASE)
        if match_q:
            if current_block:
                table = create_table_from_block(current_block)
                if table: newly_created_tables.append(table)
            current_block = {'question': match_q.group(1).strip(), 'options': []}
        
        elif 'question' in current_block:
            match_opt = re.match(r'^[a-d]\.\s*(.*)', line, re.IGNORECASE)
            match_correct = re.match(r'^Correct Option:\s*(.*)', line, re.IGNORECASE)
            match_exp = re.match(r'^Explanation:\s*(.*)', line, re.IGNORECASE)

            if match_opt and len(current_block.get('options', [])) < 4:
                current_block['options'].append(match_opt.group(1).strip())
            elif match_correct:
                current_block['correct'] = match_correct.group(1).strip()
            elif match_exp:
                current_block['explanation'] = match_exp.group(1).strip()
        
        line_number += 1

    # Process the very last block in the file
    if current_block:
        table = create_table_from_block(current_block)
        if table: newly_created_tables.append(table)

    report = "\n".join(error_report_lines)
    return newly_created_tables, report


# ==============================================================================
# === TELEGRAM BOT HANDLERS ===
# ==============================================================================

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.message.reply_html(
        "👋 **Welcome to the Final DOCX Bot!**\n\n"
        "This version includes your latest corrections:\n"
        "1️⃣ **Table Files**: Rejects tables if the 2nd column of rows 3, 4, 5, and 6 all contain 'incorrect'.\n"
        "2️⃣ **Text Files**: `Explanation:` is now optional. If missing, the cell will be blank, and the block will NOT be an error.\n\n"
        "<b>How to use:</b>\n"
        "1. Send me your <code>.docx</code> files.\n"
        "2. Use the <code>/process</code> command."
    )

async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
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
        await message.reply_text("✅ File received. Send more files or use <b>/process</b>.", parse_mode='HTML')
    except Exception as e:
        logger.error(f"Error handling document for user {user_id}: {e}")
        await message.reply_text("An error occurred while receiving your file.")

async def process(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    # This main orchestrator function remains the same.
    user_id = update.message.from_user.id
    if not context.user_data.get('files'):
        await update.message.reply_text("You haven't sent any files yet.")
        return

    input_files = context.user_data['files']
    await update.message.reply_text(f"🔄 Processing {len(input_files)} file(s) with the final logic...")

    all_processed_tables = []
    error_reports = []
    output_files = []

    try:
        for file_path in input_files:
            doc = docx.Document(file_path)
            file_name = os.path.basename(file_path)
            
            if doc.tables:
                logger.info(f"Processing '{file_name}' as a table-based document.")
                valid_tables, report = process_table_document(doc)
                all_processed_tables.extend(valid_tables)
                if report: error_reports.append(f"In '{file_name}': {report}")
            else:
                logger.info(f"Processing '{file_name}' as a text-based document.")
                new_tables, report = process_text_document(doc)
                all_processed_tables.extend(new_tables)
                if report: error_reports.append(f"In '{file_name}':\n{report}")

        if error_reports:
            await update.message.reply_text("⚠️ **Processing Report:**\n\n" + "\n\n".join(error_reports))
        
        if not all_processed_tables:
            await update.message.reply_text("ℹ️ No valid tables could be found or created from your files.")
            return

        for table in all_processed_tables:
            apply_font_modifications_to_table(table)

        await update.message.reply_text(f"✅ Found/created {len(all_processed_tables)} valid tables. Creating final document(s)...")
        for i in range(0, len(all_processed_tables), TABLES_PER_FILE):
            chunk = all_processed_tables[i:i + TABLES_PER_FILE]
            
            new_doc = docx.Document()
            new_doc.add_heading(f"Processed Tables - Part {i//TABLES_PER_FILE + 1}", level=1)
            
            for table in chunk:
                clone_table(table, new_doc)

            output_filename = os.path.join(DOWNLOAD_DIR, f"{user_id}_output_part_{i//TABLES_PER_FILE + 1}.docx")
            new_doc.save(output_filename)
            output_files.append(output_filename)
            
        for output_file in output_files:
            with open(output_file, 'rb') as f:
                await context.bot.send_document(chat_id=update.effective_chat.id, document=f)

    except Exception as e:
        logger.error(f"A critical error occurred: {e}", exc_info=True)
        await update.message.reply_text("❌ A critical error occurred during processing.")
    
    finally:
        all_temp_files = input_files + output_files
        for file_path in all_temp_files:
            if os.path.exists(file_path):
                os.remove(file_path)
        context.user_data['files'] = []

# ==============================================================================
# --- MAIN BOT EXECUTION ---
# ==============================================================================

def main() -> None:
    if BOT_TOKEN == "YOUR_TELEGRAM_BOT_TOKEN":
        print("!!! ERROR: Please replace 'YOUR_TELEGRAM_BOT_TOKEN' with your actual bot token. !!!")
        return

    if not os.path.exists(DOWNLOAD_DIR):
        os.makedirs(DOWNLOAD_DIR)

    application = Application.builder().token(BOT_TOKEN).build()
    
    application.add_handler(CommandHandler("start", start))
    application.add_handler(CommandHandler("process", process))
    application.add_handler(MessageHandler(filters.Document.ALL, handle_document))

    print("Final Corrected DOCX Bot is running...")
    application.run_polling()

if __name__ == '__main__':
    main()
                
