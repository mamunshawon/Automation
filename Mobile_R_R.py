import requests
import time

# Define your bot token and chat ID

bot_token = '7060819285:AAFmCNBitSGjkUFCaznIZU1MGJcx4hog1fc'
group_chat_id = '-4192540905'


def send_message_telegram(bot_token, chat_id, message):
    url = f'https://api.telegram.org/bot{bot_token}/sendMessage'
    data = {
        'chat_id': chat_id,
        'text': message
    }

    response = requests.post(url, data = data)
    return response

