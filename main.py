import asyncio
import logging
import openpyxl
from config import Token
from aiogram import types
from aiogram.types import Message
from aiogram.filters import Command
from aiogram import Bot, Dispatcher, F

# work = openpyxl.Workbook()
# sheet = work.active
#
# sheet["A1"] = "Ism"
# sheet["B1"] = "Familiya"
# sheet["C1"] = "Yosh"
# sheet["D1"] = "Turar joy"
# sheet["E1"] = "Tel"
#
# sheet["A2"] = "Ahmadjon"  # noqa
# sheet["B2"] = "Abdulfotih"  # noqa
# sheet["C2"] = 16
# sheet["D2"] = "Uzb, Tash"  # noqa
# sheet["E1"] = "+998770333308"
#
# work.save('exel.xlsx')

logging.basicConfig(level=logging.INFO)

myToken = Token
bot = Bot(token=myToken)
dp = Dispatcher()

count = 1


@dp.message(Command("start"))
async def start(message: types.Message):
    await message.reply(f"Salom hurmatli <b>{message.from_user.first_name}</b>.",
                        parse_mode="HTML")
    i = count + 1

    work = openpyxl.Workbook()
    sheet = work.active

    sheet["A1"] = "Ism"
    sheet["B1"] = "Familiya"
    sheet["C1"] = "Username"
    sheet["D1"] = "ID"

    sheet[f"A{i}"] = message.from_user.first_name
    sheet[f"B{i}"] = message.from_user.last_name
    sheet[f"C{i}"] = message.from_user.username
    sheet[f"D{i}"] = message.from_user.id

    work.save('info.xlsx')


async def main():
    await dp.start_polling(bot)


if __name__ == '__main__':
    asyncio.run(main())
