import re
import openpyxl
from datetime import datetime

import telebot
import config

bot = telebot.TeleBot(config.TOKEN, parse_mode=False)

curr_class = ''
subject_list = []
prev_flag = 1
flag = 0

wb = openpyxl.load_workbook('1511 расписание flat.xlsx')
sheet = wb['Sheet1']


def excel_find(cur_class, day, sheet1):
    global subject_list
    for i in range(1, sheet1.max_row):
        if sheet1[i][0].value == cur_class and sheet1[i][1].value == day:
            subject_list.append(sheet1[i][4].value if sheet1[i][4].value is not None else '')


@bot.message_handler(commands=['start'])
def start(message):
    global subject_list
    bot.send_message(message.chat.id, 'Добро пожаловать! Введите класс, в котором вы учитесь')


@bot.message_handler(commands=['help'])
def bot_help(message):
    bot.send_message(message.chat.id, 'Для того, чтобы сохранить класс, в котором '
                                      'вы учитесь, отправьте его. \nНапример: 9г2, 10И1')
    bot.send_message(message.chat.id, 'Для того, чтобы получить расписание на конкретный день недели, сначала'
                                      ' сохраните класс, затем напишите день недели. \nНапример: Понедельник')
    bot.send_message(message.chat.id, 'Бот каждый день присылает расписание на следующий день, если класс сохранён')


@bot.message_handler(content_types=['text'])
def schedule(message):
    global curr_class
    global subject_list
    global flag
    global prev_flag
    if re.fullmatch('9[А-Га-г][1-2]|8[А-Ва-в][1-2]|1[0-1][А-Га-гИи][1-2]', message.text):
        curr_class = message.text.upper()
        text = 'Класс сохранён. Бот будет отправлять каждый день в 18:00 расписание на завтра, ' \
               'либо вы можете запросить расписание на любой день недели.\nНапример: Понедельник'
        bot.send_message(message.chat.id, text)
        flag += 1

        while flag == prev_flag:
            if datetime.now().strftime('%H:%M:%S') == '18:00:00':
                if datetime.now().weekday() + 1 != 6:
                    bot.send_message(message.chat.id, f'Расписание на завтра для {curr_class}:')
                    tomorrow1 = config.days[(datetime.now().weekday() + 1) % 7][0].upper()
                    tomorrow2 = config.days[(datetime.now().weekday() + 1) % 7][1:]
                    tomorrow = tomorrow1 + tomorrow2
                    excel_find(curr_class, tomorrow, sheet)
                    bot.send_message(message.chat.id, '\n'.join(subject_list))
                    subject_list = []
                else:
                    bot.send_message(message.chat.id, 'Завтра нет уроков')
        else:
            prev_flag += 1

    elif message.text.lower() in config.days and curr_class != '':
        day = message.text[0].upper() + message.text[1:]
        bot.send_message(message.chat.id, f'{day}, расписание для {curr_class}:')
        excel_find(curr_class, day, sheet)
        bot.send_message(message.chat.id, '\n'.join(subject_list))
        subject_list = []
    else:
        bot.send_message(message.chat.id, 'Я не знаю, что ответить, введите /help для помощи')


bot.infinity_polling()
