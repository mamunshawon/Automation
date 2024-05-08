import requests

# Replace 'YOUR_BOT_TOKEN' with your actual bot token
bot_token = '7060819285:AAFmCNBitSGjkUFCaznIZU1MGJcx4hog1fc'

# Make a request to the Telegram Bot API to get recent updates
response = requests.get(f'https://api.telegram.org/bot{bot_token}/getUpdates')

# Check if the request was successful
if response.status_code == 200:
    # Parse the JSON response
    data = response.json()
    # Check if there are any updates
    if data['result']:
        # Get the chat ID of the latest message
        chat_id = data['result'][-1]['message']['chat']['id']
        print("Chat ID:", chat_id)
    else:
        print("No recent updates.")
else:
    print("Failed to fetch updates.")
