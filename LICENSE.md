import asyncio
import xlrd
from aiogram import Bot, Dispatcher, types
from aiogram.contrib.fsm_storage.memory import MemoryStorage
from aiogram.dispatcher import FSMContext
from aiogram.dispatcher.filters import Command

# Создаем экземпляр бота с помощью токена
bot_token = 'YOUR_BOT_TOKEN'
bot = Bot(token=bot_token)
storage = MemoryStorage()
dispatcher = Dispatcher(bot, storage=storage)

# Открываем файл Excel с помощью xlrd
excel_file_path = 'YOUR_EXCEL_FILE_PATH'
workbook = xlrd.open_workbook(excel_file_path)
sheet = workbook.sheet_by_index(0)

# Функция для обработки запроса и отправки данных из файла Excel
async def handle_excel_data(message: types.Message, state: FSMContext):
    for row in range(sheet.nrows):
        values = sheet.row_values(row)
        await message.answer(' '.join(str(cell) for cell in values))

# Обрабатываем команду '/excel_data'
@dispatcher.message_handler(Command('excel_data'))
async def handle_excel_data_command(message: types.Message, state: FSMContext):
    await handle_excel_data(message, state)

# Запускаем бота
async def main():
    await dispatcher.start_polling()

if __name__ == '__main__':
    loop = asyncio.get_event_loop()
    loop.run_until_complete(main())
