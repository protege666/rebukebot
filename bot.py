# -*- coding: utf-8 -*-
import telebot
from telebot import types
import openpyxl


bot = telebot.TeleBot("5156138010:AAGaiUjeH39SWDKXCVLbSA7c-Jf9lzmQCDc")
wb = openpyxl.reader.excel.load_workbook(filename='rebuke.xlsx')

comandirs = {"Командир 691/11 уч.группы": "1501065385",
             "Командир 691/12 уч.группы": "668569146",
             "Командир 691/2 уч.группы": "1272155176",
             "Командир 691/3 уч.группы": "776868996",  # my id
             "Командир 692/11 уч.группы": "853467034",
             "Командир 692/12 уч.группы": "1047619270",
             "Командир 695/11 уч.группы": "1366455616",
             "Командир 695/12 уч.группы": "553719065"}


def get_key(d, value):
    for k, v in d.items():
        if v == value:
            return k


@bot.message_handler(commands=['start'])
def welcome_start(message):
    bot.send_message(message.chat.id, f'Приветствую тебя, {message.chat.first_name}')

    # telegram_id = message.from_user.id
    # global telegram_id
    # print(telegram_id)

@bot.message_handler(commands=['prem'])
def prem_change_group(message):
    telegram_id = message.from_user.id
    print(telegram_id)
    if telegram_id == 776868996:
        markup = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=1)
        btn1 = types.KeyboardButton("Въебать выговор")
        btn2 = types.KeyboardButton("Посмотреть выговоры")
        btn3 = types.KeyboardButton("Снять выговор")
        markup.add(btn1, btn2, btn3)
        bot.send_message(message.chat.id, f'Приветствую тебя, Старшина. Что хочешь сделать?', reply_markup=markup)
        bot.register_next_step_handler(message, prem_check)
    else:
        bot.send_message(message.chat.id, f'У вас нет доступа к этой коменде')

def prem_check(message):
    try:
        text = message.text
        if text == "Въебать выговор":
            a = telebot.types.ReplyKeyboardRemove()
            bot.send_message(message.chat.id, f'Выберите группу?(Введите номер по порядку)', reply_markup=a)
            bot.register_next_step_handler(message, add_prem_from_group)
        elif text == "Посмотреть выговоры":
            a = telebot.types.ReplyKeyboardRemove()
            bot.send_message(message.chat.id, f'Выберите группу?(Введите номер по порядку)', reply_markup=a, parse_mode="html")
            bot.register_next_step_handler(message, prem_check_from_group)
        elif text == "Снять выговор":
            a = telebot.types.ReplyKeyboardRemove()
            bot.send_message(message.chat.id, f'Выберите группу?(Введите номер по порядку)', reply_markup=a)
            bot.register_next_step_handler(message, del_prem_from_group)
        else:
            a = telebot.types.ReplyKeyboardRemove()
            bot.send_message(message.chat.id, f'Error1111!! !   Try again!', reply_markup=a)
    except:
        a = telebot.types.ReplyKeyboardRemove()
        bot.send_message(message.chat.id, f'Try again!', reply_markup=a)

def add_prem_from_group(message):
    try:
        text = message.text
        global sheet
        if text == "1":

            wb.active = 0
            sheet = wb.active
            fullstrk = ''
            for i in range(2, 16):
                strk = str(i-1)+") " + str(sheet[i][0].value)+"\n"
                fullstrk += strk
            bot.send_message(message.chat.id, f'{fullstrk}\n\nКому?(Введите номер по порядку)')
            bot.register_next_step_handler(message, prem_cadet)
        elif text == "2":

            wb.active = 1
            sheet = wb.active
            fullstrk = ''
            for i in range(2, 15):
                strk = str(i-1)+") " + str(sheet[i][0].value)+"\n"
                fullstrk += strk
            bot.send_message(message.chat.id, f'{fullstrk}\n\nКому?(Введите номер по порядку)')
            bot.register_next_step_handler(message, prem_cadet)
        elif text == "3":

            wb.active = 2
            sheet = wb.active
            fullstrk = ''
            for i in range(2, 16):
                strk = str(i-1)+") " + str(sheet[i][0].value)+"\n"
                fullstrk += strk
            bot.send_message(message.chat.id, f'{fullstrk}\n\nКому?(Введите номер по порядку)')
            bot.register_next_step_handler(message, prem_cadet)
        elif text == "4":
            wb.active = 3
            sheet = wb.active
            fullstrk = ''
            for i in range(2, 17):
                strk = str(i-1)+") " + str(sheet[i][0].value)+"\n"
                fullstrk += strk
            bot.send_message(message.chat.id, f'{fullstrk}\n\nКому?(Введите номер по порядку)')
            bot.register_next_step_handler(message, prem_cadet)
        elif text == "5":
            wb.active = 4
            sheet = wb.active
            fullstrk = ''
            for i in range(2, 19):
                strk = str(i-1)+") " + str(sheet[i][0].value)+"\n"
                fullstrk += strk
            bot.send_message(message.chat.id, f'{fullstrk}\n\nКому?(Введите номер по порядку)')
            bot.register_next_step_handler(message, prem_cadet)
        elif text == "6":
            wb.active = 5
            sheet = wb.active
            fullstrk = ''
            for i in range(2, 20):
                strk = str(i-1)+") " + str(sheet[i][0].value)+"\n"
                fullstrk += strk
            bot.send_message(message.chat.id, f'{fullstrk}\n\nКому?(Введите номер по порядку)')
            bot.register_next_step_handler(message, prem_cadet)
        elif text == "7":
            wb.active = 6
            sheet = wb.active
            fullstrk = ''
            for i in range(2, 21):
                strk = str(i-1)+") " + str(sheet[i][0].value)+"\n"
                fullstrk += strk
            bot.send_message(message.chat.id, f'{fullstrk}\n\nКому?(Введите номер по порядку)')
            bot.register_next_step_handler(message, prem_cadet)
        elif text == "8":
            wb.active = 7
            sheet = wb.active
            fullstrk = ''
            for i in range(2, 24):
                strk = str(i-1)+") " + str(sheet[i][0].value)+"\n"
                fullstrk += strk
            bot.send_message(message.chat.id, f'{fullstrk}\n\nКому?(Введите номер по порядку)')
            bot.register_next_step_handler(message, prem_cadet)
    except:
        a = telebot.types.ReplyKeyboardRemove()
        bot.send_message(message.chat.id, f'Error1111!! !   Try again!', reply_markup=a)

def prem_cadet(message):
    try:
        global prem_cadet_num
        prem_cadet_num = message.text

        markup = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=1)
        NPH = types.KeyboardButton("НФ")
        NK = types.KeyboardButton("НК")
        KO = types.KeyboardButton("КО")
        Star = types.KeyboardButton("Старшина")
        KUG = types.KeyboardButton("КУГ")
        KomO = types.KeyboardButton("КомО")
        markup.add(NPH, NK, KO, Star, KUG,  KomO)
        bot.send_message(message.chat.id, 'От кого?', reply_markup=markup)
        bot.register_next_step_handler(message, prem_add_rebuke)
    except:
        a = telebot.types.ReplyKeyboardRemove()
        bot.send_message(message.chat.id, f'Error1111!! !   Try again!', reply_markup=a)

def prem_add_rebuke(message):
    whom = message.text
    nullstr = ''
    if whom == "НФ":
        nullstr += 'B'
    elif whom == "НК":
        nullstr += 'C'
    elif whom == "КО":
        nullstr += 'D'
    elif whom == "Старшина":
        nullstr += 'E'
    elif whom == "КУГ":
        nullstr += 'F'
    elif whom == "КомО":
        nullstr += 'G'

    try:
        count = sheet[nullstr + str(int(prem_cadet_num) + 1)].value
        ncount = int(count)
        ncount += 1
        sheet[nullstr + str(int(prem_cadet_num) + 1)].value = ncount
        wb.save("rebuke.xlsx")
        a = telebot.types.ReplyKeyboardRemove()
        bot.send_message(message.chat.id, 'Въебали)', reply_markup=a)
    except:
        a = telebot.types.ReplyKeyboardRemove()
        bot.send_message(message.chat.id, 'Что то пошло не так...', reply_markup=a)

def prem_check_from_group(message):

    text = message.text
    a = telebot.types.ReplyKeyboardRemove()
    telegram_id = message.from_user.id
    if text == "1":
        wb.active = 0
        sheet = wb.active
        string = ''
        for i in range(2, 16):
            if int(sheet[i][1].value) > 0 or int(sheet[i][2].value) > 0 or int(sheet[i][3].value) > 0 or int(
                    sheet[i][4].value) > 0 or int(sheet[i][5].value) > 0 or int(sheet[i][6].value) > 0:
                string += f'{i - 1}) <b>{str(sheet[i][0].value)}</b>' + '\n' + str(sheet[1][1].value) + ':' + str(
                    sheet[i][1].value) + '\n' + str(sheet[1][2].value) + ':' + str(sheet[i][2].value) + '\n' + str(
                    sheet[1][3].value) + ':' + str(sheet[i][3].value) + '\n' + str(sheet[1][4].value) + ':' + str(
                    sheet[i][4].value) + '\n' + str(sheet[1][5].value) + ':' + str(sheet[i][5].value) + '\n' + str(
                    sheet[1][6].value) + ':' + str(sheet[i][6].value) + '\n\n'
            else:
                pass
        bot.send_message(message.chat.id, f'{string}', reply_markup=a, parse_mode="html")
    elif text == "2":
        wb.active = 1
        sheet = wb.active
        string = ''
        for i in range(2, 15):
            if int(sheet[i][1].value) > 0 or int(sheet[i][2].value) > 0 or int(sheet[i][3].value) > 0 or int(
                    sheet[i][4].value) > 0 or int(sheet[i][5].value) > 0 or int(sheet[i][6].value) > 0:
                string += f'{i - 1}) <b>{str(sheet[i][0].value)}</b>' + '\n' + str(sheet[1][1].value) + ':' + str(
                    sheet[i][1].value) + '\n' + str(sheet[1][2].value) + ':' + str(sheet[i][2].value) + '\n' + str(
                    sheet[1][3].value) + ':' + str(sheet[i][3].value) + '\n' + str(sheet[1][4].value) + ':' + str(
                    sheet[i][4].value) + '\n' + str(sheet[1][5].value) + ':' + str(sheet[i][5].value) + '\n' + str(
                    sheet[1][6].value) + ':' + str(sheet[i][6].value) + '\n\n'
            else:
                pass
        bot.send_message(message.chat.id, f'{string}', reply_markup=a, parse_mode="html")
    elif text == "3":
        wb.active = 2
        sheet = wb.active
        string = ''
        for i in range(2, 16):
            if int(sheet[i][1].value) > 0 or int(sheet[i][2].value) > 0 or int(sheet[i][3].value) > 0 or int(
                    sheet[i][4].value) > 0 or int(sheet[i][5].value) > 0 or int(sheet[i][6].value) > 0:
                string += f'{i - 1}) <b>{str(sheet[i][0].value)}</b>' + '\n' + str(sheet[1][1].value) + ':' + str(
                    sheet[i][1].value) + '\n' + str(sheet[1][2].value) + ':' + str(sheet[i][2].value) + '\n' + str(
                    sheet[1][3].value) + ':' + str(sheet[i][3].value) + '\n' + str(sheet[1][4].value) + ':' + str(
                    sheet[i][4].value) + '\n' + str(sheet[1][5].value) + ':' + str(sheet[i][5].value) + '\n' + str(
                    sheet[1][6].value) + ':' + str(sheet[i][6].value) + '\n\n'
            else:
                pass
        bot.send_message(message.chat.id, f'{string}', reply_markup=a, parse_mode="html")
    elif text == "4":
        wb.active = 3
        sheet = wb.active
        string = ''
        for i in range(2, 17):
            if int(sheet[i][1].value) > 0 or int(sheet[i][2].value) > 0 or int(sheet[i][3].value) > 0 or int(
                    sheet[i][4].value) > 0 or int(sheet[i][5].value) > 0 or int(sheet[i][6].value) > 0:
                string += f'{i - 1}) <b>{str(sheet[i][0].value)}</b>' + '\n' + str(sheet[1][1].value) + ':' + str(
                    sheet[i][1].value) + '\n' + str(sheet[1][2].value) + ':' + str(sheet[i][2].value) + '\n' + str(
                    sheet[1][3].value) + ':' + str(sheet[i][3].value) + '\n' + str(sheet[1][4].value) + ':' + str(
                    sheet[i][4].value) + '\n' + str(sheet[1][5].value) + ':' + str(sheet[i][5].value) + '\n' + str(
                    sheet[1][6].value) + ':' + str(sheet[i][6].value) + '\n\n'
            else:
                pass
        bot.send_message(message.chat.id, f'{string}', reply_markup=a, parse_mode="html")
    elif text == "5":
        wb.active = 4
        sheet = wb.active
        string = ''
        for i in range(2, 19):
            if int(sheet[i][1].value) > 0 or int(sheet[i][2].value) > 0 or int(sheet[i][3].value) > 0 or int(
                    sheet[i][4].value) > 0 or int(sheet[i][5].value) > 0 or int(sheet[i][6].value) > 0:
                string += f'{i - 1}) <b>{str(sheet[i][0].value)}</b>' + '\n' + str(sheet[1][1].value) + ':' + str(
                    sheet[i][1].value) + '\n' + str(sheet[1][2].value) + ':' + str(sheet[i][2].value) + '\n' + str(
                    sheet[1][3].value) + ':' + str(sheet[i][3].value) + '\n' + str(sheet[1][4].value) + ':' + str(
                    sheet[i][4].value) + '\n' + str(sheet[1][5].value) + ':' + str(sheet[i][5].value) + '\n' + str(
                    sheet[1][6].value) + ':' + str(sheet[i][6].value) + '\n\n'
            else:
                pass
        bot.send_message(message.chat.id, f'{string}', reply_markup=a, parse_mode="html")
    elif text == "6":
        wb.active = 5
        sheet = wb.active
        string = ''
        for i in range(2, 20):
            if int(sheet[i][1].value) > 0 or int(sheet[i][2].value) > 0 or int(sheet[i][3].value) > 0 or int(
                    sheet[i][4].value) > 0 or int(sheet[i][5].value) > 0 or int(sheet[i][6].value) > 0:
                string += f'{i - 1}) <b>{str(sheet[i][0].value)}</b>' + '\n' + str(sheet[1][1].value) + ':' + str(
                    sheet[i][1].value) + '\n' + str(sheet[1][2].value) + ':' + str(sheet[i][2].value) + '\n' + str(
                    sheet[1][3].value) + ':' + str(sheet[i][3].value) + '\n' + str(sheet[1][4].value) + ':' + str(
                    sheet[i][4].value) + '\n' + str(sheet[1][5].value) + ':' + str(sheet[i][5].value) + '\n' + str(
                    sheet[1][6].value) + ':' + str(sheet[i][6].value) + '\n\n'
            else:
                pass
        bot.send_message(message.chat.id, f'{string}', reply_markup=a, parse_mode="html")
    elif text == "7":
        wb.active = 6
        sheet = wb.active
        string = ''
        for i in range(2, 21):
            if int(sheet[i][1].value) > 0 or int(sheet[i][2].value) > 0 or int(sheet[i][3].value) > 0 or int(
                    sheet[i][4].value) > 0 or int(sheet[i][5].value) > 0 or int(sheet[i][6].value) > 0:
                string += f'{i - 1}) <b>{str(sheet[i][0].value)}</b>' + '\n' + str(sheet[1][1].value) + ':' + str(
                    sheet[i][1].value) + '\n' + str(sheet[1][2].value) + ':' + str(sheet[i][2].value) + '\n' + str(
                    sheet[1][3].value) + ':' + str(sheet[i][3].value) + '\n' + str(sheet[1][4].value) + ':' + str(
                    sheet[i][4].value) + '\n' + str(sheet[1][5].value) + ':' + str(sheet[i][5].value) + '\n' + str(
                    sheet[1][6].value) + ':' + str(sheet[i][6].value) + '\n\n'
            else:
                pass
        bot.send_message(message.chat.id, f'{string}', reply_markup=a, parse_mode="html")
    elif text == "8":
        wb.active = 7
        sheet = wb.active
        string = ''
        for i in range(2, 24):
            if int(sheet[i][1].value) > 0 or int(sheet[i][2].value) > 0 or int(sheet[i][3].value) > 0 or int(
                    sheet[i][4].value) > 0 or int(sheet[i][5].value) > 0 or int(sheet[i][6].value) > 0:
                string += f'{i - 1}) <b>{str(sheet[i][0].value)}</b>' + '\n' + str(sheet[1][1].value) + ':' + str(
                    sheet[i][1].value) + '\n' + str(sheet[1][2].value) + ':' + str(sheet[i][2].value) + '\n' + str(
                    sheet[1][3].value) + ':' + str(sheet[i][3].value) + '\n' + str(sheet[1][4].value) + ':' + str(
                    sheet[i][4].value) + '\n' + str(sheet[1][5].value) + ':' + str(sheet[i][5].value) + '\n' + str(
                    sheet[1][6].value) + ':' + str(sheet[i][6].value) + '\n\n'
            else:
                pass
        bot.send_message(message.chat.id, f'{string}', reply_markup=a, parse_mode="html")
    else:
        a = telebot.types.ReplyKeyboardRemove()
        bot.send_message(message.chat.id, f'Такой группы нет!', reply_markup=a)

def del_prem_from_group(message):
    try:
        text = message.text
        global sheet
        if text == "1":

            wb.active = 0
            sheet = wb.active
            fullstrk = ''
            for i in range(2, 16):
                strk = str(i-1)+") " + str(sheet[i][0].value)+"\n"
                fullstrk += strk
            bot.send_message(message.chat.id, f'{fullstrk}\n\nКому?(Введите номер по порядку)')
            bot.register_next_step_handler(message, del_prem_cadet)
        elif text == "2":

            wb.active = 1
            sheet = wb.active
            fullstrk = ''
            for i in range(2, 15):
                strk = str(i-1)+") " + str(sheet[i][0].value)+"\n"
                fullstrk += strk
            bot.send_message(message.chat.id, f'{fullstrk}\n\nКому?(Введите номер по порядку)')
            bot.register_next_step_handler(message, del_prem_cadet)
        elif text == "3":

            wb.active = 2
            sheet = wb.active
            fullstrk = ''
            for i in range(2, 16):
                strk = str(i-1)+") " + str(sheet[i][0].value)+"\n"
                fullstrk += strk
            bot.send_message(message.chat.id, f'{fullstrk}\n\nКому?(Введите номер по порядку)')
            bot.register_next_step_handler(message, del_prem_cadet)
        elif text == "4":
            wb.active = 3
            sheet = wb.active
            fullstrk = ''
            for i in range(2, 17):
                strk = str(i-1)+") " + str(sheet[i][0].value)+"\n"
                fullstrk += strk
            bot.send_message(message.chat.id, f'{fullstrk}\n\nКому?(Введите номер по порядку)')
            bot.register_next_step_handler(message, del_prem_cadet)
        elif text == "5":
            wb.active = 4
            sheet = wb.active
            fullstrk = ''
            for i in range(2, 19):
                strk = str(i-1)+") " + str(sheet[i][0].value)+"\n"
                fullstrk += strk
            bot.send_message(message.chat.id, f'{fullstrk}\n\nКому?(Введите номер по порядку)')
            bot.register_next_step_handler(message, del_prem_cadet)
        elif text == "6":
            wb.active = 5
            sheet = wb.active
            fullstrk = ''
            for i in range(2, 20):
                strk = str(i-1)+") " + str(sheet[i][0].value)+"\n"
                fullstrk += strk
            bot.send_message(message.chat.id, f'{fullstrk}\n\nКому?(Введите номер по порядку)')
            bot.register_next_step_handler(message, del_prem_cadet)
        elif text == "7":
            wb.active = 6
            sheet = wb.active
            fullstrk = ''
            for i in range(2, 21):
                strk = str(i-1)+") " + str(sheet[i][0].value)+"\n"
                fullstrk += strk
            bot.send_message(message.chat.id, f'{fullstrk}\n\nКому?(Введите номер по порядку)')
            bot.register_next_step_handler(message, del_prem_cadet)
        elif text == "8":
            wb.active = 7
            sheet = wb.active
            fullstrk = ''
            for i in range(2, 24):
                strk = str(i-1)+") " + str(sheet[i][0].value)+"\n"
                fullstrk += strk
            bot.send_message(message.chat.id, f'{fullstrk}\n\nКому?(Введите номер по порядку)')
            bot.register_next_step_handler(message, del_prem_cadet)
    except:
        a = telebot.types.ReplyKeyboardRemove()
        bot.send_message(message.chat.id, f'Error1111!! !   Try again!', reply_markup=a)

def del_prem_cadet(message):
    try:
        global del_prem_cadet_num
        del_prem_cadet_num = message.text

        markup = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=1)
        NPH = types.KeyboardButton("НФ")
        NK = types.KeyboardButton("НК")
        KO = types.KeyboardButton("КО")
        Star = types.KeyboardButton("Старшина")
        KUG = types.KeyboardButton("КУГ")
        KomO = types.KeyboardButton("КомО")
        markup.add(NPH, NK, KO, Star, KUG,  KomO)
        bot.send_message(message.chat.id, 'От кого?', reply_markup=markup)
        bot.register_next_step_handler(message, prem_del_rebuke)
    except:
        a = telebot.types.ReplyKeyboardRemove()
        bot.send_message(message.chat.id, f'Error1111!! !   Try again!', reply_markup=a)

def prem_del_rebuke(message):
    whom = message.text
    nullstr = ''
    if whom == "НФ":
        nullstr += 'B'
    elif whom == "НК":
        nullstr += 'C'
    elif whom == "КО":
        nullstr += 'D'
    elif whom == "Старшина":
        nullstr += 'E'
    elif whom == "КУГ":
        nullstr += 'F'
    elif whom == "КомО":
        nullstr += 'G'

    try:
        count = sheet[nullstr + str(int(del_prem_cadet_num) + 1)].value
        ncount = int(count)
        ncount -= 1
        sheet[nullstr + str(int(del_prem_cadet_num) + 1)].value = ncount
        wb.save("rebuke.xlsx")
        a = telebot.types.ReplyKeyboardRemove()
        bot.send_message(message.chat.id, 'Сняли)', reply_markup=a)
    except:
        a = telebot.types.ReplyKeyboardRemove()
        bot.send_message(message.chat.id, 'Что то пошло не так...', reply_markup=a)







@bot.message_handler(commands=['doc'])
def send_doc(message):
    telegram_id = message.from_user.id
    if telegram_id == 776868996:
        bot.send_document(message.chat.id, open(r'rebuke.xlsx', 'rb'))
    else:
        bot.send_message(message.chat.id, f'Вам не доступна данная команда! Извините:(')



@bot.message_handler(commands=['rebuke'])
def change_group(message):
    try:
        telegram_id = message.from_user.id
        print(telegram_id)
        if telegram_id == 1501065385:
            markup = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=1)
            btn1 = types.KeyboardButton("Въебать выговор")
            btn2 = types.KeyboardButton("Посмотреть выговоры")
            btn3 = types.KeyboardButton("Снять выговор")
            markup.add(btn1, btn2, btn3)
            bot.send_message(message.chat.id, f'Приветствую тебя, {get_key(comandirs, "1501065385")}. Что хочешь сделать?', reply_markup=markup)
            bot.register_next_step_handler(message, check)
        elif telegram_id == 668569146:
            markup = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=1)
            btn1 = types.KeyboardButton("Въебать выговор")
            btn2 = types.KeyboardButton("Посмотреть выговоры")
            btn3 = types.KeyboardButton("Снять выговор")
            markup.add(btn1, btn2, btn3)
            bot.send_message(message.chat.id, f'Приветствую тебя, {get_key(comandirs, "668569146")}. Что хочешь сделать?', reply_markup=markup)
            bot.register_next_step_handler(message, check)
        elif telegram_id == 1272155176:
            markup = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=1)
            btn1 = types.KeyboardButton("Въебать выговор")
            btn2 = types.KeyboardButton("Посмотреть выговоры")
            btn3 = types.KeyboardButton("Снять выговор")
            markup.add(btn1, btn2, btn3)
            bot.send_message(message.chat.id, f'Приветствую тебя, {get_key(comandirs, "1272155176")}. Что хочешь сделать?', reply_markup=markup)
            bot.register_next_step_handler(message, check)
        elif telegram_id == 776868996:
            markup = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=1)
            btn1 = types.KeyboardButton("Въебать выговор")
            btn2 = types.KeyboardButton("Посмотреть выговоры")
            btn3 = types.KeyboardButton("Снять выговор")
            markup.add(btn1, btn2, btn3)
            bot.send_message(message.chat.id, f'Приветствую тебя, {get_key(comandirs, "776868996")}. Что хочешь сделать?', reply_markup=markup)
            bot.register_next_step_handler(message, check)
        elif telegram_id == 853467034:
            markup = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=1)
            btn1 = types.KeyboardButton("Въебать выговор")
            btn2 = types.KeyboardButton("Посмотреть выговоры")
            btn3 = types.KeyboardButton("Снять выговор")
            markup.add(btn1, btn2, btn3)
            bot.send_message(message.chat.id, f'Приветствую тебя, {get_key(comandirs, "853467034")}. Что хочешь сделать?', reply_markup=markup)
            bot.register_next_step_handler(message, check)
        elif telegram_id == 1047619270:
            markup = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=1)
            btn1 = types.KeyboardButton("Въебать выговор")
            btn2 = types.KeyboardButton("Посмотреть выговоры")
            btn3 = types.KeyboardButton("Снять выговор")
            markup.add(btn1, btn2, btn3)
            bot.send_message(message.chat.id, f'Приветствую тебя, {get_key(comandirs, "1047619270")}. Что хочешь сделать?', reply_markup=markup)
            bot.register_next_step_handler(message, check)
        elif telegram_id == 1366455616:
            markup = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=1)
            btn1 = types.KeyboardButton("Въебать выговор")
            btn2 = types.KeyboardButton("Посмотреть выговоры")
            btn3 = types.KeyboardButton("Снять выговор")
            markup.add(btn1, btn2, btn3)
            bot.send_message(message.chat.id, f'Приветствую тебя, {get_key(comandirs, "1366455616")}. Что хочешь сделать?', reply_markup=markup)
            bot.register_next_step_handler(message, check)
        elif telegram_id == 553719065:
            markup = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=1)
            btn1 = types.KeyboardButton("Въебать выговор")
            btn2 = types.KeyboardButton("Посмотреть выговоры")
            btn3 = types.KeyboardButton("Снять выговор")
            markup.add(btn1, btn2, btn3)
            bot.send_message(message.chat.id, f'Приветствую тебя, {get_key(comandirs, "553719065")}. Что хочешь сделать?', reply_markup=markup)
            bot.register_next_step_handler(message, check)
    except:
        a = telebot.types.ReplyKeyboardRemove()
        bot.send_message(message.chat.id, f'Error1111!! !   Try again!', reply_markup=a)


def check(message):
    try:
        telegram_id = message.from_user.id
        text = message.text
        if text == "Въебать выговор":
            a = telebot.types.ReplyKeyboardRemove()
            if telegram_id == 553719065:
                wb.active = 7
                sheet = wb.active
                fullstrk = ''
                for i in range(2, 24):
                    strk = str(i - 1) + ") " + str(sheet[i][0].value) + "\n"
                    fullstrk += strk
                bot.send_message(message.chat.id, f'{fullstrk}')

            bot.send_message(message.chat.id, f'Кому?(Введите номер по списку)', reply_markup=a)
            bot.register_next_step_handler(message, from_whom)
        elif text == "Посмотреть выговоры":
            a = telebot.types.ReplyKeyboardRemove()
            telegram_id = message.from_user.id
            if telegram_id == 1501065385:
                wb.active = 0
                sheet = wb.active
                string = ''
                for i in range(2, 16):
                    if int(sheet[i][1].value) > 0 or int(sheet[i][2].value) > 0 or int(sheet[i][3].value) > 0 or int(sheet[i][4].value) > 0 or int(sheet[i][5].value) > 0 or int(sheet[i][6].value) > 0 :
                        string += f'{i-1}) <b>{str(sheet[i][0].value)}</b>' + '\n' + str(sheet[1][1].value) + ':' + str(sheet[i][1].value) + '\n' + str(sheet[1][2].value) + ':' + str(sheet[i][2].value) + '\n' + str(sheet[1][3].value) + ':' + str(sheet[i][3].value) + '\n' + str(sheet[1][4].value) + ':' + str(sheet[i][4].value) + '\n' + str(sheet[1][5].value) + ':' + str(sheet[i][5].value) + '\n' + str(sheet[1][6].value) + ':' + str(sheet[i][6].value) + '\n\n'
                    else:
                        pass
                bot.send_message(message.chat.id, f'{string}', reply_markup=a, parse_mode="html")
            elif telegram_id == 668569146:
                wb.active = 1
                sheet = wb.active
                string = ''
                for i in range(2, 15):
                    if int(sheet[i][1].value) > 0 or int(sheet[i][2].value) > 0 or int(sheet[i][3].value) > 0 or int(sheet[i][4].value) > 0 or int(sheet[i][5].value) > 0 or int(sheet[i][6].value) > 0 :
                        string += f'{i-1}) <b>{str(sheet[i][0].value)}</b>' + '\n' + str(sheet[1][1].value) + ':' + str(sheet[i][1].value) + '\n' + str(sheet[1][2].value) + ':' + str(sheet[i][2].value) + '\n' + str(sheet[1][3].value) + ':' + str(sheet[i][3].value) + '\n' + str(sheet[1][4].value) + ':' + str(sheet[i][4].value) + '\n' + str(sheet[1][5].value) + ':' + str(sheet[i][5].value) + '\n' + str(sheet[1][6].value) + ':' + str(sheet[i][6].value) + '\n\n'
                    else:
                        pass
                bot.send_message(message.chat.id, f'{string}', reply_markup=a, parse_mode="html")
            elif telegram_id == 1272155176:
                wb.active = 2
                sheet = wb.active
                string = ''
                for i in range(2, 16):
                    if int(sheet[i][1].value) > 0 or int(sheet[i][2].value) > 0 or int(sheet[i][3].value) > 0 or int(sheet[i][4].value) > 0 or int(sheet[i][5].value) > 0 or int(sheet[i][6].value) > 0 :
                        string += f'{i-1}) <b>{str(sheet[i][0].value)}</b>' + '\n' + str(sheet[1][1].value) + ':' + str(sheet[i][1].value) + '\n' + str(sheet[1][2].value) + ':' + str(sheet[i][2].value) + '\n' + str(sheet[1][3].value) + ':' + str(sheet[i][3].value) + '\n' + str(sheet[1][4].value) + ':' + str(sheet[i][4].value) + '\n' + str(sheet[1][5].value) + ':' + str(sheet[i][5].value) + '\n' + str(sheet[1][6].value) + ':' + str(sheet[i][6].value) + '\n\n'
                    else:
                        pass
                bot.send_message(message.chat.id, f'{string}', reply_markup=a, parse_mode="html")
            elif telegram_id == 776868996:
                wb.active = 3
                sheet = wb.active
                string = ''
                for i in range(2, 17):
                    if int(sheet[i][1].value) > 0 or int(sheet[i][2].value) > 0 or int(sheet[i][3].value) > 0 or int(sheet[i][4].value) > 0 or int(sheet[i][5].value) > 0 or int(sheet[i][6].value) > 0 :
                        string += f'{i-1}) <b>{str(sheet[i][0].value)}</b>' + '\n' + str(sheet[1][1].value) + ':' + str(sheet[i][1].value) + '\n' + str(sheet[1][2].value) + ':' + str(sheet[i][2].value) + '\n' + str(sheet[1][3].value) + ':' + str(sheet[i][3].value) + '\n' + str(sheet[1][4].value) + ':' + str(sheet[i][4].value) + '\n' + str(sheet[1][5].value) + ':' + str(sheet[i][5].value) + '\n' + str(sheet[1][6].value) + ':' + str(sheet[i][6].value) + '\n\n'
                    else:
                        pass
                bot.send_message(message.chat.id, f'{string}', reply_markup=a, parse_mode="html")
            elif telegram_id == 853467034:
                wb.active = 4
                sheet = wb.active
                string = ''
                for i in range(2, 19):
                    if int(sheet[i][1].value) > 0 or int(sheet[i][2].value) > 0 or int(sheet[i][3].value) > 0 or int(sheet[i][4].value) > 0 or int(sheet[i][5].value) > 0 or int(sheet[i][6].value) > 0 :
                        string += f'{i-1}) <b>{str(sheet[i][0].value)}</b>' + '\n' + str(sheet[1][1].value) + ':' + str(sheet[i][1].value) + '\n' + str(sheet[1][2].value) + ':' + str(sheet[i][2].value) + '\n' + str(sheet[1][3].value) + ':' + str(sheet[i][3].value) + '\n' + str(sheet[1][4].value) + ':' + str(sheet[i][4].value) + '\n' + str(sheet[1][5].value) + ':' + str(sheet[i][5].value) + '\n' + str(sheet[1][6].value) + ':' + str(sheet[i][6].value) + '\n\n'
                    else:
                        pass
                bot.send_message(message.chat.id, f'{string}', reply_markup=a, parse_mode="html")
            elif telegram_id == 1047619270:
                wb.active = 5
                sheet = wb.active
                string = ''
                for i in range(2, 20):
                    if int(sheet[i][1].value) > 0 or int(sheet[i][2].value) > 0 or int(sheet[i][3].value) > 0 or int(sheet[i][4].value) > 0 or int(sheet[i][5].value) > 0 or int(sheet[i][6].value) > 0 :
                        string += f'{i-1}) <b>{str(sheet[i][0].value)}</b>' + '\n' + str(sheet[1][1].value) + ':' + str(sheet[i][1].value) + '\n' + str(sheet[1][2].value) + ':' + str(sheet[i][2].value) + '\n' + str(sheet[1][3].value) + ':' + str(sheet[i][3].value) + '\n' + str(sheet[1][4].value) + ':' + str(sheet[i][4].value) + '\n' + str(sheet[1][5].value) + ':' + str(sheet[i][5].value) + '\n' + str(sheet[1][6].value) + ':' + str(sheet[i][6].value) + '\n\n'
                    else:
                        pass
                bot.send_message(message.chat.id, f'{string}', reply_markup=a, parse_mode="html")
            elif telegram_id == 1366455616:
                wb.active = 6
                sheet = wb.active
                string = ''
                for i in range(2, 21):
                    if int(sheet[i][1].value) > 0 or int(sheet[i][2].value) > 0 or int(sheet[i][3].value) > 0 or int(sheet[i][4].value) > 0 or int(sheet[i][5].value) > 0 or int(sheet[i][6].value) > 0 :
                        string += f'{i-1}) <b>{str(sheet[i][0].value)}</b>' + '\n' + str(sheet[1][1].value) + ':' + str(sheet[i][1].value) + '\n' + str(sheet[1][2].value) + ':' + str(sheet[i][2].value) + '\n' + str(sheet[1][3].value) + ':' + str(sheet[i][3].value) + '\n' + str(sheet[1][4].value) + ':' + str(sheet[i][4].value) + '\n' + str(sheet[1][5].value) + ':' + str(sheet[i][5].value) + '\n' + str(sheet[1][6].value) + ':' + str(sheet[i][6].value) + '\n\n'
                    else:
                        pass
                bot.send_message(message.chat.id, f'{string}', reply_markup=a, parse_mode="html")
            elif telegram_id == 553719065:
                wb.active = 7
                sheet = wb.active
                string = ''
                for i in range(2, 24):
                    if int(sheet[i][1].value) > 0 or int(sheet[i][2].value) > 0 or int(sheet[i][3].value) > 0 or int(sheet[i][4].value) > 0 or int(sheet[i][5].value) > 0 or int(sheet[i][6].value) > 0 :
                        string += f'{i-1}) <b>{str(sheet[i][0].value)}</b>' + '\n' + str(sheet[1][1].value) + ':' + str(sheet[i][1].value) + '\n' + str(sheet[1][2].value) + ':' + str(sheet[i][2].value) + '\n' + str(sheet[1][3].value) + ':' + str(sheet[i][3].value) + '\n' + str(sheet[1][4].value) + ':' + str(sheet[i][4].value) + '\n' + str(sheet[1][5].value) + ':' + str(sheet[i][5].value) + '\n' + str(sheet[1][6].value) + ':' + str(sheet[i][6].value) + '\n\n'
                    else:
                        pass
                bot.send_message(message.chat.id, f'{string}', reply_markup=a, parse_mode="html")
        elif text == "Снять выговор":
            a = telebot.types.ReplyKeyboardRemove()
            if telegram_id == 553719065:
                wb.active = 7
                sheet = wb.active
                fullstrk = ''
                for i in range(2, 24):
                    strk = str(i - 1) + ") " + str(sheet[i][0].value) + "\n"
                    fullstrk += strk
                bot.send_message(message.chat.id, f'{fullstrk}')
            bot.send_message(message.chat.id, f'Кому?(Введите номер по списку)', reply_markup=a)
            bot.register_next_step_handler(message, from_whom_del)
    except:
        a = telebot.types.ReplyKeyboardRemove()
        bot.send_message(message.chat.id, f'Error1111!! !   Try again!', reply_markup=a)


def from_whom(message):
    try:
        global cadet_num
        cadet_num = message.text

        markup = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=1)
        # NPH = types.KeyboardButton("НФ")
        # NK = types.KeyboardButton("НК")
        # KO = types.KeyboardButton("КО")
        # Star = types.KeyboardButton("Старшина")
        KUG = types.KeyboardButton("КУГ")
        KomO = types.KeyboardButton("КомО")
        markup.add(KUG, KomO)
        bot.send_message(message.chat.id, 'От кого?', reply_markup=markup)
        bot.register_next_step_handler(message, add_rebuke)
    except:
        a = telebot.types.ReplyKeyboardRemove()
        bot.send_message(message.chat.id, f'Error1111!! !   Try again!', reply_markup=a)


def add_rebuke(message):
    whom = message.text
    telegram_id = message.from_user.id
    nullstr = ''
    if whom == "НФ":
        nullstr += 'B'
    elif whom == "НК":
        nullstr += 'C'
    elif whom == "КО":
        nullstr += 'D'
    elif whom == "Старшина":
        nullstr += 'E'
    elif whom == "КУГ":
        nullstr += 'F'
    elif whom == "КомО":
        nullstr += 'G'

    if telegram_id == 1501065385:
        try:
            wb.active = 0
            sheet = wb.active
            count = sheet[nullstr + str(int(cadet_num) + 1)].value
            ncount = int(count)
            ncount += 1
            sheet[nullstr + str(int(cadet_num) + 1)].value = ncount
            wb.save("rebuke.xlsx")
            a = telebot.types.ReplyKeyboardRemove()
            bot.send_message(message.chat.id, 'Въебали)', reply_markup=a)
        except:
            a = telebot.types.ReplyKeyboardRemove()
            bot.send_message(message.chat.id, 'Что то пошло не так...', reply_markup=a)
    elif telegram_id == 668569146:
        try:
            wb.active = 1
            sheet = wb.active
            count = sheet[nullstr + str(int(cadet_num) + 1)].value
            ncount = int(count)
            ncount += 1
            sheet[nullstr + str(int(cadet_num) + 1)].value = ncount
            wb.save("rebuke.xlsx")
            a = telebot.types.ReplyKeyboardRemove()
            bot.send_message(message.chat.id, 'Въебали)', reply_markup=a)
        except:
            a = telebot.types.ReplyKeyboardRemove()
            bot.send_message(message.chat.id, 'Что то пошло не так...', reply_markup=a)
    elif telegram_id == 1272155176:
        try:
            wb.active = 2
            sheet = wb.active
            count = sheet[nullstr + str(int(cadet_num) + 1)].value
            ncount = int(count)
            ncount += 1
            sheet[nullstr + str(int(cadet_num) + 1)].value = ncount
            wb.save("rebuke.xlsx")
            a = telebot.types.ReplyKeyboardRemove()
            bot.send_message(message.chat.id, 'Въебали)', reply_markup=a)
        except:
            a = telebot.types.ReplyKeyboardRemove()
            bot.send_message(message.chat.id, 'Что то пошло не так...', reply_markup=a)
    elif telegram_id == 776868996:
        try:
            wb.active = 3
            sheet = wb.active
            count = sheet[nullstr + str(int(cadet_num) + 1)].value
            ncount = int(count)
            ncount += 1
            sheet[nullstr + str(int(cadet_num) + 1)].value = ncount
            wb.save("rebuke.xlsx")
            a = telebot.types.ReplyKeyboardRemove()
            bot.send_message(message.chat.id, 'Въебали)', reply_markup=a)
        except:
            a = telebot.types.ReplyKeyboardRemove()
            bot.send_message(message.chat.id, 'Что то пошло не так...', reply_markup=a)
    elif telegram_id == 853467034:
        try:
            wb.active = 4
            sheet = wb.active
            count = sheet[nullstr + str(int(cadet_num) + 1)].value
            ncount = int(count)
            ncount += 1
            sheet[nullstr + str(int(cadet_num) + 1)].value = ncount
            wb.save("rebuke.xlsx")
            a = telebot.types.ReplyKeyboardRemove()
            bot.send_message(message.chat.id, 'Въебали)', reply_markup=a)
        except:
            a = telebot.types.ReplyKeyboardRemove()
            bot.send_message(message.chat.id, 'Что то пошло не так...', reply_markup=a)
    elif telegram_id == 1047619270:
        try:
            wb.active = 5
            sheet = wb.active
            count = sheet[nullstr + str(int(cadet_num) + 1)].value
            ncount = int(count)
            ncount += 1
            sheet[nullstr + str(int(cadet_num) + 1)].value = ncount
            wb.save("rebuke.xlsx")
            a = telebot.types.ReplyKeyboardRemove()
            bot.send_message(message.chat.id, 'Въебали)', reply_markup=a)
        except:
            a = telebot.types.ReplyKeyboardRemove()
            bot.send_message(message.chat.id, 'Что то пошло не так...', reply_markup=a)
    elif telegram_id == 1366455616:
        try:
            wb.active = 6
            sheet = wb.active
            count = sheet[nullstr + str(int(cadet_num) + 1)].value
            ncount = int(count)
            ncount += 1
            sheet[nullstr + str(int(cadet_num) + 1)].value = ncount
            wb.save("rebuke.xlsx")
            a = telebot.types.ReplyKeyboardRemove()
            bot.send_message(message.chat.id, 'Въебали)', reply_markup=a)
        except:
            a = telebot.types.ReplyKeyboardRemove()
            bot.send_message(message.chat.id, 'Что то пошло не так...', reply_markup=a)
    elif telegram_id == 553719065:
        try:
            wb.active = 7
            sheet = wb.active
            count = sheet[nullstr + str(int(cadet_num) + 1)].value
            ncount = int(count)
            ncount += 1
            sheet[nullstr + str(int(cadet_num) + 1)].value = ncount
            wb.save("rebuke.xlsx")
            a = telebot.types.ReplyKeyboardRemove()
            bot.send_message(message.chat.id, 'Въебали)', reply_markup=a)
        except:
            a = telebot.types.ReplyKeyboardRemove()
            bot.send_message(message.chat.id, 'Что то пошло не так...', reply_markup=a)


def from_whom_del(message):
    try:
        global cadet_num_del
        cadet_num_del = message.text

        markup = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=1)
        # NPH = types.KeyboardButton("НФ")
        # NK = types.KeyboardButton("НК")
        # KO = types.KeyboardButton("КО")
        # Star = types.KeyboardButton("Старшина")
        KUG = types.KeyboardButton("КУГ")
        KomO = types.KeyboardButton("КомО")
        markup.add(KUG, KomO)
        bot.send_message(message.chat.id, 'От кого?', reply_markup=markup)
        bot.register_next_step_handler(message, del_rebuke)
    except:
        a = telebot.types.ReplyKeyboardRemove()
        bot.send_message(message.chat.id, f'Error1111!! !   Try again!', reply_markup=a)


def del_rebuke(message):
    whom = message.text
    telegram_id = message.from_user.id
    nullstr = ''
    if whom == "НФ":
        nullstr += 'B'
    elif whom == "НК":
        nullstr += 'C'
    elif whom == "КО":
        nullstr += 'D'
    elif whom == "Старшина":
        nullstr += 'E'
    elif whom == "КУГ":
        nullstr += 'F'
    elif whom == "КомО":
        nullstr += 'G'

    if telegram_id == 1501065385:
        try:
            wb.active = 0
            sheet = wb.active
            count = sheet[nullstr + str(int(cadet_num_del) + 1)].value
            ncount = int(count)
            ncount -= 1
            sheet[nullstr + str(int(cadet_num_del) + 1)].value = ncount
            wb.save("rebuke.xlsx")
            a = telebot.types.ReplyKeyboardRemove()
            bot.send_message(message.chat.id, 'Сняли)', reply_markup=a)
        except:
            a = telebot.types.ReplyKeyboardRemove()
            bot.send_message(message.chat.id, 'Что то пошло не так...', reply_markup=a)
    elif telegram_id == 668569146:
        try:
            wb.active = 1
            sheet = wb.active
            count = sheet[nullstr + str(int(cadet_num_del) + 1)].value
            ncount = int(count)
            ncount -= 1
            sheet[nullstr + str(int(cadet_num_del) + 1)].value = ncount
            wb.save("rebuke.xlsx")
            a = telebot.types.ReplyKeyboardRemove()
            bot.send_message(message.chat.id, 'Сняли)', reply_markup=a)
        except:
            a = telebot.types.ReplyKeyboardRemove()
            bot.send_message(message.chat.id, 'Что то пошло не так...', reply_markup=a)
    elif telegram_id == 1272155176:
        try:
            wb.active = 2
            sheet = wb.active
            count = sheet[nullstr + str(int(cadet_num_del) + 1)].value
            ncount = int(count)
            ncount -= 1
            sheet[nullstr + str(int(cadet_num_del) + 1)].value = ncount
            wb.save("rebuke.xlsx")
            a = telebot.types.ReplyKeyboardRemove()
            bot.send_message(message.chat.id, 'Сняли)', reply_markup=a)
        except:
            a = telebot.types.ReplyKeyboardRemove()
            bot.send_message(message.chat.id, 'Что то пошло не так...', reply_markup=a)
    elif telegram_id == 776868996:
        try:
            wb.active = 3
            sheet = wb.active
            count = sheet[nullstr + str(int(cadet_num_del) + 1)].value
            ncount = int(count)
            ncount -= 1
            sheet[nullstr + str(int(cadet_num_del) + 1)].value = ncount
            wb.save("rebuke.xlsx")
            a = telebot.types.ReplyKeyboardRemove()
            bot.send_message(message.chat.id, 'Сняли)', reply_markup=a)
        except:
            a = telebot.types.ReplyKeyboardRemove()
            bot.send_message(message.chat.id, 'Что то пошло не так...', reply_markup=a)
    elif telegram_id == 853467034:
        try:
            wb.active = 4
            sheet = wb.active
            count = sheet[nullstr + str(int(cadet_num_del) + 1)].value
            ncount = int(count)
            ncount -= 1
            sheet[nullstr + str(int(cadet_num_del) + 1)].value = ncount
            wb.save("rebuke.xlsx")
            a = telebot.types.ReplyKeyboardRemove()
            bot.send_message(message.chat.id, 'Сняли)', reply_markup=a)
        except:
            a = telebot.types.ReplyKeyboardRemove()
            bot.send_message(message.chat.id, 'Что то пошло не так...', reply_markup=a)
    elif telegram_id == 1047619270:
        try:
            wb.active = 5
            sheet = wb.active
            count = sheet[nullstr + str(int(cadet_num_del) + 1)].value
            ncount = int(count)
            ncount -= 1
            sheet[nullstr + str(int(cadet_num_del) + 1)].value = ncount
            wb.save("rebuke.xlsx")
            a = telebot.types.ReplyKeyboardRemove()
            bot.send_message(message.chat.id, 'Сняли)', reply_markup=a)
        except:
            a = telebot.types.ReplyKeyboardRemove()
            bot.send_message(message.chat.id, 'Что то пошло не так...', reply_markup=a)
    elif telegram_id == 1366455616:
        try:
            wb.active = 6
            sheet = wb.active
            count = sheet[nullstr + str(int(cadet_num_del) + 1)].value
            ncount = int(count)
            ncount -= 1
            sheet[nullstr + str(int(cadet_num_del) + 1)].value = ncount
            wb.save("rebuke.xlsx")
            a = telebot.types.ReplyKeyboardRemove()
            bot.send_message(message.chat.id, 'Сняли)', reply_markup=a)
        except:
            a = telebot.types.ReplyKeyboardRemove()
            bot.send_message(message.chat.id, 'Что то пошло не так...', reply_markup=a)
    elif telegram_id == 553719065:
        try:
            wb.active = 7
            sheet = wb.active
            count = sheet[nullstr + str(int(cadet_num_del) + 1)].value
            ncount = int(count)
            ncount -= 1
            sheet[nullstr + str(int(cadet_num_del) + 1)].value = ncount
            wb.save("rebuke.xlsx")
            a = telebot.types.ReplyKeyboardRemove()
            bot.send_message(message.chat.id, 'Сняли)', reply_markup=a)
        except:
            a = telebot.types.ReplyKeyboardRemove()
            bot.send_message(message.chat.id, 'Что то пошло не так...', reply_markup=a)


# def check_rebuke(message):
#     telegram_id = message.from_user.id
#     if telegram_id == 776868996:
#         wb.active = 0
#         sheet = wb.active
#         cells = sheet['A1':'G6']
#         for i in cells:
#             for j in i:
#
#                 print(j.value)






# global telegram_id
    # telegram_id = message.from_user.id

"""@bot.message_handler(commands=['profile'])
def check_profilee(message):
    global telegram_id
    telegram_id = message.from_user.id
    result = check_profile(telegram_id)
    name = result[2]
    surname = result[1]
    pat = result[3]
    post = result[4]

    print(result)
    markup_inline = types.InlineKeyboardMarkup()
    firsttest = types.InlineKeyboardButton(text='Пройти начальное тестирование', callback_data='first_test')
    markup_inline.add(firsttest)
    bot.send_message(message.chat.id,
                     f"{smile[1]} Ваш профиль {smile[1]}\n\nФамилия: {surname}\nИмя: {name}\nОтчество: {pat}\nДолжность: {post}\nУникальный идентификатор: {telegram_id}\n\n{smile[1]} Успеваемость {smile[1]}\n\nМодуль 1: {modlist[0]}\nМодуль 2: {modlist[1]}\nМодуль 3: Не пройден\nМодуль 4: Не пройден\nМодуль 5: Не пройден\nМодуль 6: Не пройден\nМодуль 7: Не пройден\n", reply_markup=markup_inline)
"""


bot.polling(none_stop=True)

