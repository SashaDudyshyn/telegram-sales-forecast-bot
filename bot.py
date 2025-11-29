# bot.py ВЕРСІЯ ДЛЯ RENDER (Web Service)
import asyncio
import aiohttp
import os
from aiohttp import web
from aiogram import Bot, Dispatcher, types, F
from aiogram.filters import Command
from aiogram.types import BufferedInputFile, FSInputFile, ReplyKeyboardRemove
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage

API_URL = "https://sales-analysis-forecasting-systems-api.onrender.com/process-excel/"
BOT_TOKEN = os.environ.get("BOT_TOKEN")

bot = Bot(token=BOT_TOKEN)
storage = MemoryStorage()
dp = Dispatcher(storage=storage)


class Form(StatesGroup):
    waiting_file = State()
    column_year = State()
    column_month = State()
    range_data = State()
    row_title = State()
    row_first_data = State()
    row_last_data = State()
    k_value = State()
    sheet_stat = State()
    sheet_factor = State()


TEMPLATE_PATH = "шаблон.xlsx"


def back_cancel_kb():
    keyboard = [[types.KeyboardButton(text="Назад"), types.KeyboardButton(text="Скасувати")]]
    return types.ReplyKeyboardMarkup(keyboard=keyboard, resize_keyboard=True)


@dp.message(Command("start"))
async def start(message: types.Message, state: FSMContext):
    # Перевіряємо, чи є шаблон
    if os.path.exists(TEMPLATE_PATH):
        await message.answer_document(
            FSInputFile(TEMPLATE_PATH),
            caption="Привіт! Я бот для прогнозу продажів\n\n"
                    "Завантаж свій файл, який буди містити два аркуші: статистичні дані та фактори впливу,\n\n"
                    "Або завантаж цей шаблон, заповни його своїми даними та поверни його мені\n\n"
                    "Після цього я зроблю повний аналіз та прогноз на наступний рік"
        )
    else:
        await message.answer(
            "Привіт! Я бот для прогнозу продажів\n\n"
            "Надішли мені файл, який буди містити два аркуші: статистичні дані та фактори впливу (.xlsx) — і я зроблю прогноз\n\n"
            "Шаблон для заповнення можна отримати у адміністратора"
        )

    await message.answer(
        "Тепер надішли свій заповнений файл:",
        reply_markup=ReplyKeyboardRemove()
    )
    await state.set_state(Form.waiting_file)


@dp.message(Form.waiting_file, F.document)
async def file_received(message: types.Message, state: FSMContext):
    if not message.document.file_name.lower().endswith(('.xlsx', '.xls')):
        await message.answer("Потрібен файл Excel (.xlsx або .xls)")
        return

    file = await bot.get_file(message.document.file_id)
    file_bytes = await bot.download_file(file.file_path)

    await state.update_data(
        file_bytes=file_bytes.read(),
        filename=message.document.file_name
    )

    await message.answer(
        "Файл отримано!\n\n"
        "Тепер налаштуємо параметри.\n"
        "Колонка з роками? (наприклад: B)",
        reply_markup=back_cancel_kb()
    )
    await state.set_state(Form.column_year)

async def handle_back_or_cancel(message: types.Message, state: FSMContext):
    if message.text == "Скасувати":
        await state.clear()
        await message.answer("Скасовано. Надішли новий файл або /start", reply_markup=ReplyKeyboardRemove())
        return True
    return False


@dp.message(Form.column_year)
async def set_year_col(message: types.Message, state: FSMContext):
    if await handle_back_or_cancel(message, state): return
    if message.text == "Назад":
        await message.answer("Надішли файл ще раз:", reply_markup=ReplyKeyboardRemove())
        await state.set_state(Form.waiting_file)
        return
    await state.update_data(column_year=message.text.strip().upper())
    await message.answer("Колонка з місяцями? (наприклад: D)", reply_markup=back_cancel_kb())
    await state.set_state(Form.column_month)


@dp.message(Form.column_month)
async def set_month_col(message: types.Message, state: FSMContext):
    if await handle_back_or_cancel(message, state): return
    if message.text == "Назад":
        await message.answer("Колонка з роками?", reply_markup=back_cancel_kb())
        await state.set_state(Form.column_year)
        return
    await state.update_data(column_month=message.text.strip().upper())
    await message.answer("Діапазон даних? (наприклад: G-J)", reply_markup=back_cancel_kb())
    await state.set_state(Form.range_data)


@dp.message(Form.range_data)
async def set_range(message: types.Message, state: FSMContext):
    if await handle_back_or_cancel(message, state): return
    if message.text == "Назад":
        await message.answer("Колонка з місяцями?", reply_markup=back_cancel_kb())
        await state.set_state(Form.column_month)
        return
    await state.update_data(range_data=message.text.strip().upper())
    await message.answer("Рядок з назвами колонок? (наприклад: 3)", reply_markup=back_cancel_kb())
    await state.set_state(Form.row_title)


@dp.message(Form.row_title)
async def set_title_row(message: types.Message, state: FSMContext):
    if await handle_back_or_cancel(message, state): return
    if message.text == "Назад":
        await message.answer("Діапазон даних?", reply_markup=back_cancel_kb())
        await state.set_state(Form.range_data)
        return
    if not message.text.isdigit():
        await message.answer("Введи число")
        return
    await state.update_data(row_title=int(message.text))
    await message.answer("Перший рядок з даними? (наприклад: 4)", reply_markup=back_cancel_kb())
    await state.set_state(Form.row_first_data)


@dp.message(Form.row_first_data)
async def set_first_row(message: types.Message, state: FSMContext):
    if await handle_back_or_cancel(message, state): return
    if message.text == "Назад":
        await message.answer("Рядок з назвами колонок?", reply_markup=back_cancel_kb())
        await state.set_state(Form.row_title)
        return
    if not message.text.isdigit():
        await message.answer("Введи число")
        return
    await state.update_data(row_first_data=int(message.text))
    await message.answer("Останній рядок з даними? (наприклад: 38)", reply_markup=back_cancel_kb())
    await state.set_state(Form.row_last_data)


@dp.message(Form.row_last_data)
async def set_last_row(message: types.Message, state: FSMContext):
    if await handle_back_or_cancel(message, state): return
    if message.text == "Назад":
        await message.answer("Перший рядок з даними?", reply_markup=back_cancel_kb())
        await state.set_state(Form.row_first_data)
        return
    if not message.text.isdigit():
        await message.answer("Введи число")
        return
    await state.update_data(row_last_data=int(message.text))
    await message.answer("Параметр згладжування k? (зазвичай 2)", reply_markup=back_cancel_kb())
    await state.set_state(Form.k_value)


@dp.message(Form.k_value)
async def set_k(message: types.Message, state: FSMContext):
    if await handle_back_or_cancel(message, state): return
    if message.text == "Назад":
        await message.answer("Останній рядок з даними?", reply_markup=back_cancel_kb())
        await state.set_state(Form.row_last_data)
        return
    if not message.text.isdigit():
        await message.answer("Введи число")
        return
    await state.update_data(k=int(message.text))
    await message.answer("Назва аркуша зі статистикою? (наприклад: Статистичні дані)", reply_markup=back_cancel_kb())
    await state.set_state(Form.sheet_stat)


@dp.message(Form.sheet_stat)
async def set_sheet_stat(message: types.Message, state: FSMContext):
    if await handle_back_or_cancel(message, state): return
    if message.text == "Назад":
        await message.answer("Параметр k?", reply_markup=back_cancel_kb())
        await state.set_state(Form.k_value)
        return
    await state.update_data(sheet_stat=message.text.strip())
    await message.answer("Назва аркуша з факторами? (наприклад: Фактори впливу)", reply_markup=back_cancel_kb())
    await state.set_state(Form.sheet_factor)


@dp.message(Form.sheet_factor)
async def final_step(message: types.Message, state: FSMContext):
    if await handle_back_or_cancel(message, state): return
    if message.text == "Назад":
        await message.answer("Назва аркуша зі статистикою?", reply_markup=back_cancel_kb())
        await state.set_state(Form.sheet_stat)
        return

    await state.update_data(sheet_factor=message.text.strip())

    status_msg = await message.answer("Всі параметри отримано!\nВідправляю на обробку... (до 40 сек)", reply_markup=ReplyKeyboardRemove())

    data = await state.get_data()
    file_bytes = data["file_bytes"]
    filename = data["filename"]

    params = {
        "column_year": data["column_year"],
        "column_month": data["column_month"],
        "range_data": data["range_data"],
        "row_title": str(data["row_title"]),
        "row_first_data": str(data["row_first_data"]),
        "row_last_data": str(data["row_last_data"]),
        "k": str(data["k"]),
        "sheet_stat": data["sheet_stat"],
        "sheet_factor": data["sheet_factor"],
    }

    try:
        connector = aiohttp.TCPConnector(ssl=False)
        timeout = aiohttp.ClientTimeout(total=300)

        async with aiohttp.ClientSession(connector=connector, timeout=timeout) as session:
            form = aiohttp.FormData()
            form.add_field('file', file_bytes, filename=filename,
                           content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            for k, v in params.items():
                form.add_field(k, v)

            async with session.post(API_URL, data=form) as resp:
                if resp.status == 200:
                    result = await resp.read()
                    await status_msg.delete()
                    await message.answer_document(
                        BufferedInputFile(result, filename=f"Прогноз_{filename}"),
                        caption="Готово! Ось твій файл з повним аналізом та прогнозом"
                    )
                else:
                    error = await resp.text()
                    await status_msg.edit_text(f"Помилка сервера: {resp.status}\n{error[:800]}")
    except Exception as e:
        await status_msg.edit_text(f"Помилка: {str(e)}")

    await state.clear()
    await message.answer("Готовий до нового файлу! Надішли ще один або /start")

# ===  ЕНДПОІНТ ДЛЯ RENDER ===
async def health(request):
    return web.Response(text="Bot is alive!")


# === ОСНОВНА ФУНКЦІЯ ЗАПУСКУ БОТА ===
async def start_bot():
    print("Запускаю Telegram-бота...")
    await dp.start_polling(bot)


# === ВЕБ-СЕРВЕР ДЛЯ RENDER ===
async def start_web_server():
    app = web.Application()
    app.router.add_get('/', health)
    runner = web.AppRunner(app)
    await runner.setup()
    site = web.TCPSite(runner, '0.0.0.0', int(os.environ.get('PORT', 10000)))
    await site.start()
    print(f"Веб-сервер запущено на порту {os.environ.get('PORT', 10000)}")


# === ГОЛОВНИЙ ЗАПУСК ===
async def main():
    # Запускаємо і бота, і веб-сервер одночасно
    await asyncio.gather(
        start_bot(),
        start_web_server()
    )


if __name__ == "__main__":
    asyncio.run(main())