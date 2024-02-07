import logging
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Application, CommandHandler, CallbackQueryHandler, MessageHandler, filters, ContextTypes
from search import Finder
import os

# Настройка логирования
logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', level=logging.INFO)
logger = logging.getLogger(__name__)

finder = Finder("/app/Data.xlsx")

def generate_preview_texts(texts, name, result_type):
    previews = ["Выберете номер:"]
    for index, text in enumerate(texts, start=1):
        if result_type == 'persona':
            position_line = next((line for line in text.split('\n') if 'Должность:' in line), "Должность: Неизвестна")
            position = position_line.split(': ')[1] if ': ' in position_line else "Неизвестна"
            previews.append(f"{index}. {name.title()} ({position.rstrip(';')})")
        elif result_type == 'department':
            department_line = next((line for line in text.split('\n') if 'Наименование:' in line), None)
            upper_department_line = next((line for line in text.split('\n') if 'Вышестоящее подразделение верхнего уровня:' in line), None)

            department_name = department_line.split(': ')[1] if department_line and ': ' in department_line else name
            upper_department = upper_department_line.split(': ')[1] if upper_department_line and ': ' in upper_department_line else "Неизвестно"
            previews.append(f"{index}. {department_name.title()} ({upper_department.rstrip(';')})")
    return previews

def prepare_message_parts(text, max_length=4096):
    parts = text.split('\n')
    current_message = parts[0]
    result = []

    for part in parts[1:]:
        if len(current_message) + len(part) + 1 <= max_length:
            current_message += '\n' + part
        else:
            result.append(current_message)
            current_message = part

    if current_message:
        result.append(current_message)

    return result

def create_pagination_keyboard(texts, current_page=0, page_size=5):
    keyboard = []
    page_texts = texts[current_page*page_size:(current_page+1)*page_size]
    for i, text in enumerate(page_texts, start=current_page*page_size + 1):
        keyboard.append([InlineKeyboardButton(str(i), callback_data=f"text_{i}")])
    if current_page > 0:
        keyboard.append([InlineKeyboardButton("Назад", callback_data=f"page_{current_page-1}")])
    if len(texts) > (current_page + 1) * page_size:
        keyboard.append([InlineKeyboardButton("Дальше", callback_data=f"page_{current_page+1}")])
    keyboard.append([InlineKeyboardButton("Отмена", callback_data="cancel")])
    return InlineKeyboardMarkup(keyboard)

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text('Привет! Я - чат-бот управления делами.\nВведите ФИО руководителя, '
                                    'наименование подразделения или шифр подразделения для поиска в моей базе знаний!')

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_input = update.message.text
    context.user_data['last_query'] = user_input
    result = finder.search(user_input)

    logger.info(f'Пришел запрос: "{user_input}"')

    if result[0] == "Данные не найдены!":
        logger.info(f"Данные по запросу '{user_input}' не найдены!")
        await update.message.reply_text(result[0])
        await update.message.reply_text('Введите следующее ФИО, наименование подразделения или его шифр для поиска данных:')

    elif result[0] == "Точного совпадения нет.. Ближайшие результаты:":
        keyboard = [[InlineKeyboardButton(result[i][0], callback_data=f"choice_{i}")]
                    for i in range(1, 4)]
        keyboard.append([InlineKeyboardButton("Назад", callback_data='back')])
        reply_markup = InlineKeyboardMarkup(keyboard)
        await update.message.reply_text(result[0], reply_markup=reply_markup)

    else:
        selected_object = result[1]
        name = selected_object[0]
        texts = selected_object[1:]
        result_type = 'persona' if 'Должность:' in texts[0] else 'department'
        find_text = f"Найден сотрудник: __*{name}*__" if result_type == 'persona' else f"Найдено подразделение: __*{name}*__"
        await update.message.reply_text(text=find_text, parse_mode='Markdown')
        await process_search_results(update, context, name, texts, result_type)

async def process_search_results(update, context, name, texts, result_type):
    context.user_data['texts'] = texts
    if len(texts) == 1:
        await send_message_parts(update.message, texts[0], result_type)
        await update.message.reply_text('Введите следующее ФИО, наименование подразделения или его шифр для поиска данных:')
    else:
        # Добавляем превью
        previews = generate_preview_texts(texts, name, result_type)
        preview_text = '\n'.join(previews)
        await update.message.reply_text(preview_text)

        reply_markup = create_pagination_keyboard(texts, current_page=0)
        await update.message.reply_text("Выберите текст:", reply_markup=reply_markup)


async def send_message_parts(message, text, result_type):
    message_parts = prepare_message_parts(text)
    for part in message_parts:
        await message.reply_text(part)
    link_message = "Приказы о полномочиях работников находятся [здесь](https://ud.hse.ru/powers)." if result_type == 'persona' else "Полная организационная структура университета [здесь](https://www.hse.ru/orgstructure/)."
    await message.reply_text(link_message, parse_mode='Markdown')

async def button_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data

    if data.startswith('page_'):
        page_number = int(data.split('_')[1])
        texts = context.user_data['texts']
        reply_markup = create_pagination_keyboard(texts, current_page=page_number)
        await query.message.edit_text("Выберите текст:", reply_markup=reply_markup)

    elif data.startswith('text_'):
        text_index = int(data.split('_')[1]) - 1
        texts = context.user_data['texts']
        if 0 <= text_index < len(texts):
            result_type = 'persona' if 'Должность:' in texts[text_index] else 'department'
            await query.edit_message_text(text=f"Выбранный текст (№{text_index + 1}):")
            await send_message_parts(query.message, texts[text_index], result_type)
            await query.message.reply_text('Введите следующее ФИО, наименование подразделения или его шифр для поиска данных:')
        else:
            await query.message.edit_text("Ошибка: недопустимый индекс текста.", reply_markup=None)

    elif data == 'cancel':
        await query.message.edit_text("Поиск отменен.", reply_markup=None)

    elif data == 'back':
        await query.message.edit_text(text='Введите следующее ФИО, наименование подразделения или его шифр для поиска данных:')

    elif data.startswith('choice_'):
        choice_index = int(data.split('_')[1])
        last_query = context.user_data.get('last_query', '')
        result = finder.search(last_query)
        if result and 1 <= choice_index < len(result):
            selected_object = result[choice_index]
            name = selected_object[0]
            logger.info(f'Отправлен ответ: "{name}"')
            texts = selected_object[1:]
            result_type = 'persona' if 'Должность:' in texts[0] else 'department'
            await query.edit_message_text(text=f"__*{name}*__", parse_mode='Markdown')
            await process_search_results(query, context, name, texts, result_type)
        else:
            await query.message.edit_text("Ошибка: недопустимый выбор.", reply_markup=None)

def main():
    application = Application.builder().token(os.getenv("BOT_TOKEN")).build()
    application.add_handler(CommandHandler("start", start))
    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    application.add_handler(CallbackQueryHandler(button_handler))
    application.run_polling()

if __name__ == "__main__":
    main()
