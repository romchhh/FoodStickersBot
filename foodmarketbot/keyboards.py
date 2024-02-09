from telebot import types

def main_keyboard():
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=2)
    food_stickers_button = types.KeyboardButton("Наліпки страв")
    fridge_stickers_button = types.KeyboardButton("Наліпки для холодильників")
    name_stickers_button = types.KeyboardButton("Наліпки імен")
    complex_stickers_button = types.KeyboardButton("Наліпки комплексів")
    orders_button = types.KeyboardButton("Наліпки замовлень")

    markup.add(food_stickers_button, orders_button, name_stickers_button, complex_stickers_button, fridge_stickers_button)
    
    return markup


def dishes_keyboard():
    markup_dishes = types.InlineKeyboardMarkup()
    day_before_yesterday_button = types.InlineKeyboardButton("Завтра", callback_data="day_before_yesterday")
    enter_date_button = types.InlineKeyboardButton("Ввести власну дату", callback_data="enter_date")
    markup_dishes.row(day_before_yesterday_button, enter_date_button)

    return markup_dishes

def fridge_keyboard():
    markup_fridge = types.InlineKeyboardMarkup()
    tomorrow_button = types.InlineKeyboardButton("Завтра", callback_data="tomorrow")
    after_tomorrow_button = types.InlineKeyboardButton("Післязавтра", callback_data="after_tomorrow")
    enter_date_fridge_button = types.InlineKeyboardButton("Ввести власну дату", callback_data="enter_date_fridge")
    markup_fridge.row(tomorrow_button, after_tomorrow_button)
    markup_fridge.row( enter_date_fridge_button)
    
    return markup_fridge