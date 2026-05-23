#!/usr/bin/env python3
"""
Telegram бот для опросов о тренировках.
Бот отправляет опросы за 4 дня до тренировки (пятница и воскресенье)
и сохраняет результаты в Excel таблицу.
"""

import logging
import asyncio
from datetime import datetime, timedelta
from pathlib import Path
from openpyxl import Workbook, load_workbook
from telegram import Update
from telegram.ext import Application, CommandHandler, PollAnswerHandler, ContextTypes

# Настройка логирования
logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)
logger = logging.getLogger(__name__)

# КОНФИГУРАЦИЯ
# Замените на ваш токен бота от @BotFather
BOT_TOKEN = "YOUR_BOT_TOKEN_HERE"
# ID группы, куда добавлять бота (можно найти через @userinfobot или добавив бота в группу)
GROUP_ID = "YOUR_GROUP_ID_HERE"
# ID ветки (forum topic) для опросов в группе
THREAD_ID = None  # Укажите ID темы "опросы", если группа имеет формат форума
# Путь к файлу Excel для хранения статистики
EXCEL_FILE = "attendance.xlsx"


def get_next_friday():
    """Возвращает дату следующей пятницы."""
    today = datetime.now().date()
    days_until_friday = (4 - today.weekday()) % 7
    if days_until_friday == 0:
        days_until_friday = 7
    return today + timedelta(days=days_until_friday)


def get_next_sunday():
    """Возвращает дату следующего воскресенья."""
    today = datetime.now().date()
    days_until_sunday = (6 - today.weekday()) % 7
    if days_until_sunday == 0:
        days_until_sunday = 7
    return today + timedelta(days=days_until_sunday)


def init_excel_file():
    """Инициализирует Excel файл с заголовками."""
    excel_path = Path(EXCEL_FILE)
    
    if not excel_path.exists():
        wb = Workbook()
        ws = wb.active
        ws.title = "Посещаемость"
        
        # Заголовки
        ws.cell(row=1, column=1, value="Имя")
        ws.cell(row=1, column=2, value="Посещаемость (%)")
        
        wb.save(EXCEL_FILE)
        logger.info(f"Создан новый файл {EXCEL_FILE}")
    
    return excel_path.exists()


def add_user_if_not_exists(user_id: str, username: str):
    """Добавляет пользователя в таблицу, если его там нет."""
    wb = load_workbook(EXCEL_FILE)
    ws = wb.active
    
    # Проверяем, есть ли уже пользователь
    for row in range(2, ws.max_row + 1):
        if ws.cell(row=row, column=1).value == user_id:
            wb.close()
            return
    
    # Добавляем нового пользователя
    new_row = ws.max_row + 1
    ws.cell(row=new_row, column=1, value=user_id)
    ws.cell(row=new_row, column=2, value=0)
    
    wb.save(EXCEL_FILE)
    wb.close()
    logger.info(f"Добавлен пользователь: {username} ({user_id})")


def add_training_date_column(training_date: str):
    """Добавляет колонку для даты тренировки, если её нет."""
    wb = load_workbook(EXCEL_FILE)
    ws = wb.active
    
    # Ищем колонку с этой датой
    for col in range(3, ws.max_column + 1):
        if ws.cell(row=1, column=col).value == training_date:
            wb.close()
            return col
    
    # Добавляем новую колонку
    new_col = ws.max_column + 1
    ws.cell(row=1, column=new_col, value=training_date)
    
    # Заполняем нулями существующих пользователей
    for row in range(2, ws.max_row + 1):
        ws.cell(row=row, column=new_col, value=0)
    
    wb.save(EXCEL_FILE)
    wb.close()
    logger.info(f"Добавлена колонка для даты: {training_date}")
    return new_col


def record_vote(user_id: str, username: str, training_date: str, vote_value: int):
    """Записывает результат голосования в таблицу."""
    add_user_if_not_exists(user_id, username)
    date_col = add_training_date_column(training_date)
    
    wb = load_workbook(EXCEL_FILE)
    ws = wb.active
    
    # Находим строку пользователя
    user_row = None
    for row in range(2, ws.max_row + 1):
        if ws.cell(row=row, column=1).value == user_id:
            user_row = row
            break
    
    if user_row:
        # Записываем результат голосования
        ws.cell(row=user_row, column=date_col, value=vote_value)
        
        # Пересчитываем посещаемость
        total_trainings = 0
        attended_trainings = 0
        for col in range(3, ws.max_column + 1):
            cell_value = ws.cell(row=user_row, column=col).value
            if cell_value is not None and isinstance(cell_value, (int, float)):
                total_trainings += 1
                if cell_value == 1:
                    attended_trainings += 1
        
        attendance_percent = round((attended_trainings / total_trainings * 100), 1) if total_trainings > 0 else 0
        ws.cell(row=user_row, column=2, value=attendance_percent)
        
        wb.save(EXCEL_FILE)
        logger.info(f"Записан голос пользователя {username}: {vote_value} на {training_date}")
    
    wb.close()


async def send_poll(application, title: str, training_date: str):
    """Отправляет опрос в группу."""
    try:
        poll = await application.bot.send_poll(
            chat_id=GROUP_ID,
            question=title,
            options=["иду", "плачу", "не иду"],
            is_anonymous=False,
            allows_multiple_answers=False,
            message_thread_id=THREAD_ID  # Отправка в конкретную ветку (если форум)
        )
        logger.info(f"Опрос отправлен: {title}")
        logger.info(f"ID опроса: {poll.poll.id}")
        return poll.poll.id
    except Exception as e:
        logger.error(f"Ошибка при отправке опроса: {e}")
        return None


async def send_friday_poll(application):
    """Отправляет опрос для пятничной тренировки."""
    friday_date = get_next_friday()
    title = f"Тренировка {friday_date.strftime('%d.%m')} 18:00 БНТУ"
    await send_poll(application, title, friday_date.strftime('%d.%m'))


async def send_sunday_poll(application):
    """Отправляет опрос для воскресной тренировки."""
    sunday_date = get_next_sunday()
    title = f"Тренировка {sunday_date.strftime('%d.%m')} 10:00 РГУОР"
    await send_poll(application, title, sunday_date.strftime('%d.%m'))


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обработчик команды /start."""
    await update.message.reply_text(
        "Привет! Я бот для записи на тренировки.\n"
        "Тренировки проходят по пятницам (18:00 БНТУ) и воскресеньям (10:00 РГУОР).\n"
        "Опросы отправляются за 4 дня до тренировки.\n"
        "Используйте команду /help для получения дополнительной информации."
    )


async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обработчик команды /help."""
    await update.message.reply_text(
        "Доступные команды:\n"
        "/start - Начать работу с ботом\n"
        "/help - Показать эту справку\n"
        "/status - Показать текущую статистику посещаемости\n"
        "/test_poll - Тестовый опрос (только для администраторов)"
    )


async def status_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обработчик команды /status - показывает статистику."""
    try:
        wb = load_workbook(EXCEL_FILE)
        ws = wb.active
        
        if ws.max_row < 2:
            await update.message.reply_text("Пока нет данных о посещаемости.")
            wb.close()
            return
        
        message = "📊 Статистика посещаемости:\n\n"
        for row in range(2, min(ws.max_row + 1, 11)):  # Показываем топ-10
            user_id = ws.cell(row=row, column=1).value
            attendance = ws.cell(row=row, column=2).value
            
            # Получаем имя пользователя (попытка)
            try:
                user = await context.bot.get_chat(user_id)
                username = user.first_name
            except:
                username = str(user_id)
            
            message += f"{username}: {attendance}%\n"
        
        if ws.max_row > 10:
            message += f"\n... и ещё {ws.max_row - 10} участников"
        
        await update.message.reply_text(message)
        wb.close()
    except Exception as e:
        logger.error(f"Ошибка при получении статуса: {e}")
        await update.message.reply_text("Произошла ошибка при получении статистики.")


async def test_poll_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Тестовый опрос для проверки работы бота."""
    # Проверка прав администратора (упрощённая)
    chat_member = await context.bot.get_chat_member(
        chat_id=update.effective_chat.id,
        user_id=update.effective_user.id
    )
    
    if chat_member.status not in ['administrator', 'creator']:
        await update.message.reply_text("Эта команда доступна только администраторам.")
        return
    
    # Создаём тестовую дату
    test_date = datetime.now().strftime('%d.%m')
    title = f"ТЕСТ Тренировка {test_date}"
    
    await send_poll(context.application, title, test_date)
    await update.message.reply_text("Тестовый опрос отправлен!")


async def handle_poll_answer(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обработчик ответов на опросы."""
    poll_answer = update.poll_answer
    
    try:
        # Получаем информацию об опросе
        poll = await context.bot.get_poll(poll_answer.poll_id)
        question = poll.question
        
        # Извлекаем дату из вопроса
        parts = question.split()
        if len(parts) >= 2:
            date_str = parts[1]  # Формат DD.MM
        else:
            date_str = datetime.now().strftime('%d.%m')
        
        # Получаем информацию о пользователе
        user = await context.bot.get_chat(poll_answer.user_id)
        username = user.first_name or user.username or str(poll_answer.user_id)
        
        # Определяем значение голоса
        option_index = poll_answer.option_ids[0] if poll_answer.option_ids else 0
        # "иду" = 1, "плачу" = 0, "не иду" = 0
        vote_value = 1 if option_index == 0 else 0
        
        # Записываем в таблицу
        record_vote(str(poll_answer.user_id), username, date_str, vote_value)
        
        logger.info(f"Голос обработан: {username} -> {'иду' if vote_value == 1 else 'не идёт'} на {date_str}")
        
    except Exception as e:
        logger.error(f"Ошибка при обработке ответа на опрос: {e}")


async def scheduled_task(context: ContextTypes.DEFAULT_TYPE):
    """Ежедневная проверка: нужно ли отправить опрос."""
    today = datetime.now().date()
    
    # Проверяем, какой сегодня день недели
    # Вторник (1) - отправляем опрос на пятницу (через 3 дня)
    # Среда (2) - отправляем опрос на воскресенье (через 4 дня)
    
    weekday = today.weekday()
    
    if weekday == 1:  # Вторник
        logger.info("Вторник - отправляем опрос на пятницу")
        await send_friday_poll(context.application)
    elif weekday == 2:  # Среда
        logger.info("Среда - отправляем опрос на воскресенье")
        await send_sunday_poll(context.application)


async def post_init(application):
    """Инициализация после запуска бота."""
    logger.info("Бот успешно запущен!")
    logger.info(f"Excel файл: {EXCEL_FILE}")
    logger.info(f"Группа: {GROUP_ID}")
    if THREAD_ID:
        logger.info(f"Ветка (topic): {THREAD_ID}")


def main():
    """Основная функция запуска бота."""
    # Инициализация Excel файла
    init_excel_file()
    
    # Создание приложения
    application = Application.builder().token(BOT_TOKEN).post_init(post_init).build()
    
    # Добавление обработчиков команд
    application.add_handler(CommandHandler("start", start))
    application.add_handler(CommandHandler("help", help_command))
    application.add_handler(CommandHandler("status", status_command))
    application.add_handler(CommandHandler("test_poll", test_poll_command))
    
    # Обработчик ответов на опросы
    application.add_handler(PollAnswerHandler(handle_poll_answer))
    
    # Добавляем ежедневную задачу в 10:00
    job_queue = application.job_queue
    job_queue.run_daily(scheduled_task, time=datetime.strptime("10:00", "%H:%M").time())
    logger.info("Запланирована ежедневная задача на 10:00")
    
    # Запуск бота
    logger.info("Бот запущен и ожидает команды...")
    application.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
