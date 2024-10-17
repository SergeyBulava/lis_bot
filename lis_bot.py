import openpyxl
import time
from telebot import TeleBot, types

book = openpyxl.open('lsx.xlsx', read_only=True)

sheet_6 = book.worksheets[8]

bot = TeleBot(token='7713792741:AAFJ4ISKF6sor5OO2-m0CeXDa3uspgbW1aY')
l = '☀️'
n = '🌙'

def find_day(point):
    schedule = ' '
    for i in range(point, point+9, 2):
        schedule += sheet_6[i][2].value
        schedule += ' '
        if sheet_6[i][13].value is None:
            schedule += '-!'
        else:
            schedule += sheet_6[i][13].value
            schedule += '!'
    schedule = schedule.split('!')
    return schedule

@bot.message_handler(commands=['start'])
def wake_up(message):
    markup = types.ReplyKeyboardMarkup()
    bot1 = types.KeyboardButton(text='Светлая неделя'+l)
    bot2 = types.KeyboardButton(text='Темная неделя'+n)
    markup.add(bot1, bot2)
    bot.send_message(message.chat.id, 'Привет! Выбери текущую неделю.', reply_markup=markup)
    bot.register_next_step_handler(message, txt)

def txt(message):
    if message.text == 'Светлая неделя'+l:
        markup = types.ReplyKeyboardMarkup()
        bot1 = types.KeyboardButton(text='Понедельник' + l)
        bot2 = types.KeyboardButton(text='Вторник' + l)
        bot3 = types.KeyboardButton(text='Среда' + l)
        bot4 = types.KeyboardButton(text='Четверг' + l)
        bot5 = types.KeyboardButton(text='Пятница' + l)
        #bot6 = types.KeyboardButton(text='Суббота' + l)
        #bot7 = types.KeyboardButton(text='Воскресенье' + l)
        bot8 = types.KeyboardButton(text='/start')
        markup.add(bot1, bot2, bot3, bot4, bot5, bot8)
        bot.send_message(message.chat.id, 'Теперь выбери день недели.', reply_markup=markup)
    elif message.text == 'Темная неделя'+n:
        markup = types.ReplyKeyboardMarkup()
        bot1 = types.KeyboardButton(text='Понедельник' + n)
        bot2 = types.KeyboardButton(text='Вторник' + n)
        bot3 = types.KeyboardButton(text='Среда' + n)
        bot4 = types.KeyboardButton(text='Четверг' + n)
        bot5 = types.KeyboardButton(text='Пятница' + n)
        #bot6 = types.KeyboardButton(text='Суббота' + n)
        #bot7 = types.KeyboardButton(text='Воскресенье' + n)
        bot8 = types.KeyboardButton(text='/start')
        markup.add(bot1, bot2, bot3, bot4, bot5, bot8)
        bot.send_message(message.chat.id, 'Теперь выбери день недели.', reply_markup=markup)

def send(txt, message):
    for i in txt:
        bot.send_message(message.chat.id, i)


@bot.message_handler(content_types=['text'])
def day(message):
    t = message.text
    if t == 'Понедельник' + l:
        txt = find_day(3)
        send(txt, message)
    elif t == 'Понедельник' + n:
        txt = find_day(3)
        send(txt, message)
    elif t == 'Вторник' + l:
        txt = find_day(13)
        send(txt, message)
    elif t == 'Вторник' + n:
        txt = find_day(14)
        send(txt, message)
    elif t == 'Среда' + l:
        txt = find_day(23)
        send(txt, message)
    elif t == 'Среда' + n:
        txt = find_day(24)
        send(txt, message)
    elif t == 'Четверг' + l:
        txt = find_day(33)
        send(txt, message)
    elif t == 'Четверг' + n:
        txt = find_day(34)
        send(txt, message)
    elif t == 'Пятница' + l:
        txt = find_day(43)
        send(txt, message)
    elif t == 'Пятница' + n:
        txt = find_day(44)
        send(txt, message)

while True:
    try:
        bot.polling(none_stop=True, interval=0, timeout=20)
    except Exception as E:
        time.sleep(1)
