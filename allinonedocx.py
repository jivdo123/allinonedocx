import os
import re
import io
import logging
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes
from docx import Document
from docx.shared import Pt

# --- Configuration ---
# --- Paste your bot token here ---
TELEGRAM_BOT_TOKEN = '8112681572:AAHXFkLmUkwsRxcpx8GN0FCvd8gsnxFOk3I' 

# Number of questions to include in each generated DOCX file
QUESTIONS_PER_FILE = 30
# Official MIME type for .docx files, used for filtering
DOCX_MIME_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'

# --- Logging Setup ---
# Enables logging to see errors and bot activity in the console
logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)
logger = logging.getLogger(__name__)


# --- Helper Functions ---

def set_font_for_cell(cell, size_in_pt, is_bold=False):
    """
    Iterates through paragraphs and runs in a cell to set the font properties.
    """
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            font = run.font
            font.size = Pt(size_in_pt)
            font.bold = is_bold

def parse_text_question(block: str):
    """
    Parses a single block of text based on a fixed line structure.
    """
    block = block.strip()
    if not block:
        return None

    lines = [line.strip() for line in block.split('\n') if line.strip()]

    if len(lines) < 5:
        raise ValueError("The question block is incomplete. It must have a question and at least four options.")

    correct_option_match = re.search(r'Correct Option:\s*(\S+)', block, re.IGNORECASE)
    if not correct_option_match:
        raise ValueError("The 'Correct Option: [id]' line is missing.")
    correct_option_id = correct_option_match.group(1).lower()

    question_text = re.sub(r'^(?:\d+\.|Q\.)\s*', '', lines[0])

    parsed_options = []
    option_ids = ['a', 'b', 'c', 'd']
    for i in range(4):
        option_line = lines[i + 1]
        option_text = re.sub(r'^[a-zA-Z\d]+[\.\)]\s*', '', option_line)
        parsed_options.append({'id': option_ids[i], 'text': option_text})
    
    explanation_text = ""
    if len(lines) > 5:
        explanation_lines = []
        for line in lines[5:]:
            if 'Correct Option:' not in line:
                explanation_lines.append(line)
        explanation_text = "\n".join(explanation_lines).strip()

    return {
        'question_text': question_text,
        'options': parsed_options,
        'correct_option_id': correct_option_id,
        'explanation_text': explanation_text
    }


def create_formatted_docx_stream(questions_data) -> io.BytesIO:
    """
    Generates a .docx file in memory with formatted tables for each question
    and returns it as a byte stream.
    """
    doc = Document()
    
    for q_data in questions_data:
        table = doc.add_table(rows=0, cols=3)
        table.style = 'Table Grid'
        
        # --- 1. Question Row ---
        row_cells = table.add_row().cells
        row_cells[0].text = 'Question'
        row_cells[1].merge(row_cells[2]).text = q_data['question_text']
        set_font_for_cell(row_cells[1], 14) # Apply 14pt font

        # --- 2. Type Row ---
        row_cells = table.add_row().cells
        row_cells[0].text = 'Type'
        row_cells[1].merge(row_cells[2]).text = 'multiple_choice'

        # --- 3. Option Rows ---
        correct_id = q_data.get('correct_option_id')
        correct_index = q_data.get('correct_option_index')

        for i, option in enumerate(q_data['options']):
            row_cells = table.add_row().cells
            row_cells[0].text = 'Option'
            row_cells[1].text = option['text']
            set_font_for_cell(row_cells[1], 12) # Apply 12pt font
            
            is_correct = False
            if correct_id is not None:
                if option.get('id', '').lower() == correct_id.lower():
                    is_correct = True
            elif correct_index is not None:
                if i == correct_index:
                    is_correct = True
            
            row_cells[2].text = 'correct' if is_correct else 'incorrect'

        # --- 4. Solution Row ---
        row_cells = table.add_row().cells
        row_cells[0].text = 'Solution'
        row_cells[1].merge(row_cells[2]).text = q_data['explanation_text']
        set_font_for_cell(row_cells[1], 12) # Apply 12pt font

        # --- 5. Marks Row ---
        row_cells = table.add_row().cells
        row_cells[0].text = 'Marks'
        row_cells[1].text = '4'
        row_cells[2].text = '1'

        doc.add_paragraph('') # Add space between tables
        
    # Save the document to an in-memory stream
    output_stream = io.BytesIO()
    doc.save(output_stream)
    output_stream.seek(0) # Rewind the stream to the beginning for reading
    return output_stream


# --- Telegram Bot Handlers ---

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handler for the /start command."""
    await update.message.reply_text(
        "Hello! 👋\n\n"
        "Please send me a .txt or .docx file with your questions, or forward a Telegram Quiz.\n\n"
        "I will convert them into a structured and formatted .docx file for you. "
        f"If a file has more than {QUESTIONS_PER_FILE} questions, I'll create multiple documents."
    )


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handles file uploads (.txt, .docx), extracts text, and creates formatted .docx files."""
    doc_file = update.message.document
    temp_file_path = None # To ensure cleanup happens

    await update.message.reply_text(f"Processing your file: {doc_file.file_name} ... ⏳")
    
    try:
        # 1. Download the file from Telegram
        temp_file_path = f'input_{doc_file.file_id}{os.path.splitext(doc_file.file_name)[1]}'
        file = await doc_file.get_file()
        await file.download_to_drive(temp_file_path)

        # 2. Read the file content based on its type
        content = ""
        if doc_file.mime_type == 'text/plain':
            with open(temp_file_path, 'r', encoding='utf-8') as f:
                content = f.read()
        elif doc_file.mime_type == DOCX_MIME_TYPE:
            doc = Document(temp_file_path)
            content = "\n".join([p.text for p in doc.paragraphs])
        else:
            await update.message.reply_text(f"❌ Unsupported file type: {doc_file.mime_type}")
            return
            
        # 3. Parse the extracted text content
        question_blocks = re.split(r'\n\s*\n', content.strip())
        valid_questions = []
        failed_blocks = []

        for i, block in enumerate(question_blocks):
            if not block.strip(): continue
            try:
                parsed_q = parse_text_question(block)
                if parsed_q:
                    valid_questions.append(parsed_q)
            except ValueError as e:
                failed_blocks.append(f"❗️ ERROR IN QUESTION #{i+1}\nReason: {e}")

        # 4. Report any parsing errors
        if failed_blocks:
            error_summary = "\n\n".join(failed_blocks)
            await update.message.reply_text(f"Found some issues in your file:\n\n{error_summary}")

        # 5. Process valid questions and create batched DOCX files
        if valid_questions:
            total_q = len(valid_questions)
            num_files = (total_q + QUESTIONS_PER_FILE - 1) // QUESTIONS_PER_FILE
            
            await update.message.reply_text(
                f"✅ Successfully parsed {total_q} question(s). "
                f"Generating {num_files} formatted DOCX file(s) for you now..."
            )
            
            # Split valid_questions into chunks
            for i in range(0, total_q, QUESTIONS_PER_FILE):
                chunk = valid_questions[i:i + QUESTIONS_PER_FILE]
                part_num = (i // QUESTIONS_PER_FILE) + 1
                
                # Create formatted docx in memory
                docx_stream = create_formatted_docx_stream(chunk)
                
                output_filename = f"Formatted_Questions_Part_{part_num}.docx"
                await update.message.reply_document(document=docx_stream, filename=output_filename)
                
        elif not failed_blocks:
            await update.message.reply_text("🤔 No valid questions found in the file.")

    except Exception as e:
        logger.error(f"An error occurred in handle_document: {e}")
        await update.message.reply_text(f"🆘 An unexpected error occurred while processing your file.")
        
    finally:
        # 6. Clean up the downloaded file
        if temp_file_path and os.path.exists(temp_file_path):
            os.remove(temp_file_path)


async def handle_quiz(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handles native Telegram quizzes."""
    poll = update.message.poll
    
    if poll.type != 'quiz':
        await update.message.reply_text("This is a regular poll, not a quiz. I can only process quizzes.")
        return

    await update.message.reply_text("Processing quiz... ⏳")

    try:
        quiz_data = {
            'question_text': poll.question,
            'options': [{'text': opt.text} for opt in poll.options],
            'correct_option_index': poll.correct_option_id,
            'explanation_text': poll.explanation or ""
        }

        # Create the formatted docx in memory
        docx_stream = create_formatted_docx_stream([quiz_data])
        
        await update.message.reply_text(f"✅ Quiz processed successfully!")
        await update.message.reply_document(
            document=docx_stream,
            filename="Formatted_Quiz.docx"
        )
    except Exception as e:
        logger.error(f"An error occurred in handle_quiz: {e}")
        await update.message.reply_text("❌ Sorry, something went wrong while processing the quiz.")


def main():
    """Starts the bot and adds all handlers."""
    application = Application.builder().token(TELEGRAM_BOT_TOKEN).build()
    
    application.add_handler(CommandHandler("start", start))
    
    # This single handler now listens for both .txt and .docx files
    combined_filter = filters.Document.TXT | filters.Document.MimeType(DOCX_MIME_TYPE)
    application.add_handler(MessageHandler(combined_filter, handle_document))
    
    # Handler for quizzes
    application.add_handler(MessageHandler(filters.POLL, handle_quiz))
    
    # Optional: Guide users who send plain text instead of files
    async def guide_user(update: Update, context: ContextTypes.DEFAULT_TYPE):
        await update.message.reply_text("Please send your questions as a .txt or .docx file. I don't process plain text messages.")
    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, guide_user))
    
    logger.info("Bot started... (Press Ctrl+C to stop)")
    application.run_polling()


if __name__ == '__main__':
    main()
        
