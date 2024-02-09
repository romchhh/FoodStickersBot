import telebot
from telebot import types
from config import TOKEN
import os
import datetime  

from stickers.orders import generate_word_from_excel_orders
from stickers.names import read_names_from_excel, generate_word_from_excel_names
from stickers.fridge import read_fridge_from_excel, generate_word_from_excel_fridge
from stickers.dishes import read_dishes_from_excel, generate_word_from_excel_dishes
from stickers.complex import generate_word_from_excel_complex
from keyboards import main_keyboard, dishes_keyboard, fridge_keyboard

bot = telebot.TeleBot(TOKEN)

waiting_for_document_names = False 
waiting_for_document_dishes = False
waiting_for_document_fridge = False
waiting_for_document_complex = False
waiting_for_document_orders = False


@bot.message_handler(commands=['start'])
def start(message):
    global waiting_for_document_names
    global waiting_for_document_dishes
    global waiting_for_document_fridge
    global waiting_for_document_orders
    global waiting_for_document_complex
    waiting_for_document_names = False  
    waiting_for_document_dishes = False
    waiting_for_document_fridge = False
    waiting_for_document_orders = False
    waiting_for_document_complex = False
    
    markup = main_keyboard()
    bot.send_message(message.chat.id, "Виберіть категорію наліпок:", reply_markup=markup)

@bot.message_handler(func=lambda message: message.text == "Наліпки замовлень")
def handle_name_stickers_orders(message):
    global waiting_for_document_orders, waiting_for_document_complex, waiting_for_document_dishes, waiting_for_document_fridge, waiting_for_document_names
    waiting_for_document_names = False  
    waiting_for_document_dishes = False
    waiting_for_document_fridge = False
    waiting_for_document_orders = True
    waiting_for_document_complex = False  # Зміна статусу при натисканні іншої кнопки
    bot.send_message(message.chat.id, "Будь ласка, надішліть файл Excel з замовленнями.")

@bot.message_handler(func=lambda message: message.text == "Наліпки комплексів")
def handle_name_stickers_complex(message):
    global waiting_for_document_orders, waiting_for_document_complex, waiting_for_document_dishes, waiting_for_document_fridge, waiting_for_document_names
    waiting_for_document_names = False  
    waiting_for_document_dishes = False
    waiting_for_document_fridge = False
    waiting_for_document_orders = False
    waiting_for_document_complex = True
    bot.send_message(message.chat.id, "Будь ласка, надішліть файл Excel з комплексами.")

@bot.message_handler(func=lambda message: message.text == "Наліпки імен")
def handle_name_stickers(message):
    global waiting_for_document_orders, waiting_for_document_complex, waiting_for_document_dishes, waiting_for_document_fridge, waiting_for_document_names
    waiting_for_document_names = True  
    waiting_for_document_dishes = False
    waiting_for_document_fridge = False
    waiting_for_document_orders = False
    waiting_for_document_complex = False 
    bot.send_message(message.chat.id, "Будь ласка, надішліть файл Excel з іменами.")

@bot.message_handler(func=lambda message: message.text == "Наліпки страв")
def handle_name_stickers(message):
    global waiting_for_document_orders, waiting_for_document_complex, waiting_for_document_dishes, waiting_for_document_fridge, waiting_for_document_names
    waiting_for_document_names = False  
    waiting_for_document_dishes = True
    waiting_for_document_fridge = False
    waiting_for_document_orders = False
    waiting_for_document_complex = False 
    bot.send_message(message.chat.id, "Будь ласка, надішліть файл Excel зі стравами.")

@bot.message_handler(func=lambda message: message.text == "Наліпки для холодильників")
def handle_name_stickers(message):
    global waiting_for_document_orders, waiting_for_document_complex, waiting_for_document_dishes, waiting_for_document_fridge, waiting_for_document_names
    waiting_for_document_names = False  
    waiting_for_document_dishes = False
    waiting_for_document_fridge = True
    waiting_for_document_orders = False
    waiting_for_document_complex = False  
    bot.send_message(message.chat.id, "Будь ласка, надішліть файл Excel для холодильників.")

@bot.message_handler(content_types=['document'])
def handle_document(message):
    global waiting_for_document_orders, waiting_for_document_complex, waiting_for_document_dishes, waiting_for_document_fridge, waiting_for_document_names

    if waiting_for_document_names and message.document.file_name.endswith('.xlsx'):
        try:
            file_info = bot.get_file(message.document.file_id)
            downloaded_file = bot.download_file(file_info.file_path)
            with open('file.xlsx', 'wb') as new_file:
                new_file.write(downloaded_file)

            global selected_names
            selected_names = read_names_from_excel('file.xlsx')

            markup = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=1)
            generate_button = types.KeyboardButton("Згенерувати наліпки імен")
            cancel_button = types.KeyboardButton("❌Скасувати")
            markup.add(generate_button, cancel_button)
            bot.send_message(message.chat.id, "Файл успішно завантажено. Оберіть дію:", reply_markup=markup)
            waiting_for_document_names = False

        except Exception as e:
            bot.send_message(message.chat.id, f"Помилка обробки файлу, спробуйте будь ласка ще раз.")
        waiting_for_document_names = False

    elif waiting_for_document_dishes and message.document.file_name.endswith('.xlsx'):
        try:
            file_info = bot.get_file(message.document.file_id)
            downloaded_file = bot.download_file(file_info.file_path)
            with open('file.xlsx', 'wb') as new_file:
                new_file.write(downloaded_file)

            global selected_dishes
            selected_dishes = read_dishes_from_excel('file.xlsx')

            markup = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=1)
            generate_dishes_button = types.KeyboardButton("Згенерувати наліпки страв")
            cancel_button = types.KeyboardButton("❌Скасувати")
            markup.add(generate_dishes_button, cancel_button)
            bot.send_message(message.chat.id, "Файл успішно завантажено. Оберіть дію:", reply_markup=markup)

        except Exception as e:
            bot.send_message(message.chat.id, f"Помилка обробки файлу, спробуйте будь ласка ще раз.")
        waiting_for_document_dishes = False
        
    elif waiting_for_document_fridge and message.document.file_name.endswith('.xlsx'):
        try:
            file_info = bot.get_file(message.document.file_id)
            downloaded_file = bot.download_file(file_info.file_path)
            with open('file.xlsx', 'wb') as new_file:
                new_file.write(downloaded_file)

            global selected_fridge
            selected_fridge = read_fridge_from_excel('file.xlsx')

            markup = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=1)
            generate_fridge_button = types.KeyboardButton("Згенерувати наліпки для холодильника")
            cancel_button = types.KeyboardButton("❌Скасувати")
            markup.add(generate_fridge_button, cancel_button)
            bot.send_message(message.chat.id, "Файл успішно завантажено. Оберіть дію:", reply_markup=markup)

        except Exception as e:
            bot.send_message(message.chat.id, f"Помилка обробки файлу, спробуйте будь ласка ще раз.")
        waiting_for_document_fridge = False
        
    elif waiting_for_document_complex and message.document.file_name.endswith('.xlsx'):
        try:
            file_info = bot.get_file(message.document.file_id)
            downloaded_file = bot.download_file(file_info.file_path)
            with open('file.xlsx', 'wb') as new_file:
                new_file.write(downloaded_file)

            global selected_complex
            selected_complex = read_fridge_from_excel('file.xlsx')

            markup = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=1)
            generate_complex_button = types.KeyboardButton("Згенерувати наліпки комплексів")
            cancel_button = types.KeyboardButton("❌Скасувати")
            markup.add(generate_complex_button, cancel_button)
            bot.send_message(message.chat.id, "Файл успішно завантажено. Оберіть дію:", reply_markup=markup)

        except Exception as e:
            bot.send_message(message.chat.id, f"Помилка обробки файлу, спробуйте будь ласка ще раз.")
        waiting_for_document_fridge = False
        
    elif waiting_for_document_orders and message.document.file_name.endswith('.xlsx'):
        try:
            file_info = bot.get_file(message.document.file_id)
            downloaded_file = bot.download_file(file_info.file_path)
            with open('file.xlsx', 'wb') as new_file:
                new_file.write(downloaded_file)

            global selected_orders
            selected_orders = read_names_from_excel('file.xlsx')

            markup = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=1)
            generate_orders_button = types.KeyboardButton("Згенерувати наліпки замовлень")
            cancel_button = types.KeyboardButton("❌Скасувати")
            markup.add(generate_orders_button, cancel_button)
            bot.send_message(message.chat.id, "Файл успішно завантажено. Оберіть дію:", reply_markup=markup)

        except Exception as e:
            bot.send_message(message.chat.id, f"Помилка обробки файлу, спробуйте будь ласка ще раз.")
        waiting_for_document_orders = False

    else:
        bot.send_message(message.chat.id, "Будь ласка, надішліть файл Excel з розширенням .xlsx або оберіть тип наліпок.")

waiting_for_document_complex = False
complex_values = ["complex1", "complex2", "mini1", "mini2"]
current_complex_index = 0
message_to_edit = None
complex_info = {}
user_values = []  

@bot.message_handler(func=lambda message: message.text == "Згенерувати наліпки комплексів")
def handle_name_stickers(message):
    global waiting_for_document_complex
    global current_complex_index
    global message_to_edit
    global complex_info
    global user_values  
    waiting_for_document_complex = True 
    current_complex_index = 0
    complex_info = {}
    user_values = [] 
    message_text = "Будь ласка, введіть ціле числове значення для {}".format(complex_values[current_complex_index])
    message_to_edit = bot.send_message(message.chat.id, message_text)

@bot.message_handler(func=lambda message: waiting_for_document_complex)
def handle_complex_value(message):
    global waiting_for_document_complex
    global current_complex_index
    global message_to_edit
    global complex_info
    global user_values 
    global complex1_value, complex2_value, mini1_value, mini2_value  
    try:
        complex_value = int(message.text)
        complex_info[complex_values[current_complex_index]] = complex_value
        user_values.append(complex_value) 

        if complex_values[current_complex_index] == "complex1":
            complex1_value = complex_value
        elif complex_values[current_complex_index] == "complex2":
            complex2_value = complex_value
        elif complex_values[current_complex_index] == "mini1":
            mini1_value = complex_value
        elif complex_values[current_complex_index] == "mini2":
            mini2_value = complex_value

        message_text = "Кількість комплексів:"
        for key, value in complex_info.items():
            message_text += "\n{} - {}".format(key, value)

        current_complex_index += 1

        if current_complex_index < len(complex_values):
            message_text += "\nБудь ласка, введіть ціле числове значення для {}".format(complex_values[current_complex_index])
        else:
            waiting_for_document_complex = False
            message_text += "\nГенеруємо наліпки...."

        bot.edit_message_text(chat_id=message_to_edit.chat.id, message_id=message_to_edit.message_id, text=message_text)
        
        if not waiting_for_document_complex:
            generate_word_from_excel_complex('file.xlsx', 'stickers.docx', complex1_value, complex2_value, mini1_value, mini2_value)
                    
            with open('stickers.docx', 'rb') as docx_file:
                bot.send_document(message.chat.id, docx_file)

            return_to_main_menu(message.chat.id)

            os.remove('stickers.docx')
            os.remove('file.xlsx')
    
    except ValueError:
        bot.send_message(message.chat.id, "Будь ласка, введіть ціле числове значення.")

@bot.message_handler(func=lambda message: message.text == "Згенерувати наліпки замовлень")
def generate_stickers_orders(message):
    try:   
        generate_word_from_excel_orders('file.xlsx', 'stickers.docx')
        
        with open('stickers.docx', 'rb') as docx_file:
            bot.send_document(message.chat.id, docx_file)

        return_to_main_menu(message.chat.id)

        os.remove('stickers.docx')
        os.remove('file.xlsx')

    except Exception as e:
        bot.send_message(message.chat.id, f"Помилка генерації наліпок, спробуйте будь ласка ще раз.")
        os.remove('file.xlsx')
        
@bot.message_handler(func=lambda message: message.text == "Згенерувати наліпки імен")
def generate_stickers(message):
    try:
        global selected_names          
        selected_names = []  

        generate_word_from_excel_names('file.xlsx', 'stickers.docx')
        
        with open('stickers.docx', 'rb') as docx_file:
            bot.send_document(message.chat.id, docx_file)

        return_to_main_menu(message.chat.id)

        os.remove('stickers.docx')
        os.remove('file.xlsx')

    except Exception as e:
        bot.send_message(message.chat.id, f"Помилка генерації наліпок, спробуйте будь ласка ще раз.")
        os.remove('file.xlsx')

@bot.message_handler(func=lambda message: message.text == "Згенерувати наліпки страв")
def generate_stickers_dishes(message):
    try:
        global selected_names 
        selected_names = []  

        markup = dishes_keyboard()

        bot.send_message(message.chat.id, "Виберіть дату виготовлення:", reply_markup=markup)

    except Exception as e:
        bot.send_message(message.chat.id, f"Помилка генерації наліпок, спробуйте будь ласка ще раз.")
        os.remove('file.xlsx')
        
def handle_user_entered_date_dishes(message):
    try:
        user_entered_date_str = message.text.strip()
        selected_date_dishes = datetime.datetime.strptime(user_entered_date_str, '%d.%m').date()

        bot.send_message(message.chat.id, f"Ви ввели дату виготовлення: {selected_date_dishes}")

        generate_word_from_excel_dishes('file.xlsx', 'stickers.docx', selected_date_dishes)
        with open('stickers.docx', 'rb') as docx_file:
            bot.send_document(message.chat.id, docx_file)

        return_to_main_menu(message.chat.id)

        os.remove('stickers.docx')
        os.remove('file.xlsx')

    except ValueError:
        bot.send_message(message.chat.id, "Невірний формат дати. Будь ласка, введіть дату в форматі DD-MM.")
        bot.register_next_step_handler(message, handle_user_entered_date_dishes)

@bot.message_handler(func=lambda message: message.text == "Згенерувати наліпки для холодильника")
def generate_stickers_fridge(message):
    try:
        markup_fridge = fridge_keyboard()
        bot.send_message(message.chat.id, "Виберіть термін придатності:", reply_markup=markup_fridge)

    except Exception as e:
        bot.send_message(message.chat.id, f"Помилка генерації наліпок, спробуйте будь ласка ще раз.")

@bot.callback_query_handler(func=lambda call: True)
def handle_inline_keyboard(call):
    try:
        if call.data == "day_before_yesterday" or call.data == "enter_date":
            handle_inline_keyboard_dishes(call)
        elif call.data == "tomorrow" or call.data == "after_tomorrow" or call.data == "enter_date_fridge":
            handle_inline_fridge_keyboard(call)
    except Exception as e:
        bot.send_message(call.message.chat.id, f"Помилка генерації наліпок, спробуйте будь ласка ще раз.")
  
@bot.callback_query_handler(func=lambda call: True)
def handle_inline_keyboard_dishes(call):
    try:
        if call.data == "day_before_yesterday":
            selected_date_dishes = datetime.date.today() + datetime.timedelta(days=1)
        elif call.data == "enter_date":
            bot.send_message(call.message.chat.id, "Введіть дату в форматі DD-MM (наприклад, 07.12):")

            bot.register_next_step_handler(call.message, handle_user_entered_date_dishes)
            return
        
        bot.send_message(call.message.chat.id, f"Ви ввели дату виготовлення: {selected_date_dishes.strftime('%d.%m')}")

        generate_word_from_excel_dishes('file.xlsx', 'stickers.docx', selected_date_dishes)
        with open('stickers.docx', 'rb') as docx_file:
            bot.send_document(call.message.chat.id, docx_file)
        return_to_main_menu(call.message.chat.id)
        os.remove('stickers.docx')
        os.remove('file.xlsx')

    except Exception as e:
        bot.send_message(call.message.chat.id, f"Помилка генерації наліпок, спробуйте будь ласка ще раз.")

@bot.callback_query_handler(func=lambda call: True)
def handle_inline_fridge_keyboard(call):
    try:
        if call.data == "tomorrow":
            selected_date_fridge = datetime.date.today() + datetime.timedelta(days=1)
        elif call.data == "after_tomorrow":
            selected_date_fridge = datetime.date.today() + datetime.timedelta(days=2)
        elif call.data == "enter_date_fridge":
            
            bot.send_message(call.message.chat.id, "Введіть дату в форматі DD-MM (наприклад, 07.12):")
            bot.register_next_step_handler(call.message, handle_user_entered_date_fridge)
            return  

        bot.send_message(call.message.chat.id, f"Ви ввели термін придатності: {selected_date_fridge.strftime('%d.%m')}")
        generate_word_from_excel_fridge('file.xlsx', 'stickers.docx', selected_date_fridge)

        with open('stickers.docx', 'rb') as docx_file:
            bot.send_document(call.message.chat.id, docx_file)
        return_to_main_menu(call.message.chat.id)
        os.remove('stickers.docx')
        os.remove('file.xlsx')

    except Exception as e:
        bot.send_message(call.message.chat.id, f"Помилка генерації наліпок, спробуйте будь ласка ще раз.")

def handle_user_entered_date_fridge(message):
    try:
        user_entered_date_fridge_str = message.text.strip()
        selected_date_fridge = datetime.datetime.strptime(user_entered_date_fridge_str, '%d.%m').date()

        bot.send_message(message.chat.id, f"Ви ввели термін придатності: {selected_date_fridge}")

        generate_word_from_excel_fridge('file.xlsx', 'stickers.docx', selected_date_fridge)
        with open('stickers.docx', 'rb') as docx_file:
            bot.send_document(message.chat.id, docx_file)
        return_to_main_menu(message.chat.id)
        os.remove('stickers.docx')
        os.remove('file.xlsx')

    except ValueError:
        bot.send_message(message.chat.id, "Невірний формат дати. Будь ласка, введіть дату в форматі DD-MM.")
        bot.register_next_step_handler(message, handle_user_entered_date_fridge)

@bot.message_handler(func=lambda message: message.text == "❌Скасувати")
def cancel_generation(message):
    global selected_names
    selected_names = []
    os.remove('file.xlsx')
    return_to_main_menu(message.chat.id)

def return_to_main_menu(chat_id):
    markup = main_keyboard()
    bot.send_message(chat_id, "Виберіть категорію наліпок:", reply_markup=markup)

if __name__ == '__main__':
    bot.polling(none_stop=True)
