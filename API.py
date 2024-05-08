import random
import requests
import time

# Define your bot token and chat ID
bot_token = '7060819285:AAFmCNBitSGjkUFCaznIZU1MGJcx4hog1fc'
group_chat_id = '-4192540905'


def send_telegram_message(bot_token, chat_id, message):
    url = f'https://api.telegram.org/bot{bot_token}/sendMessage'
    data = {
        'chat_id': chat_id,
        'text': message
    }
    response = requests.post(url, data=data)
    return response


def get_access_token():
    # First API details
    first_api_url = "https://api-external.mynagad.com/ibas/api/token"
    username = "ibasExtUserV1"
    password = "ibasExtUserV1"
    payload = f'username={username}&password={password}&grant_type=password'
    headers = {
        'Content-Type': 'application/x-www-form-urlencoded',
        'Authorization': 'Basic TmFnYWRfaUJBUysrX1NlcnZpY2U6aUJBUysrX1NlcnZpY2U=',
        'Cookie': 'BIGipServerPOOL_NGD_EXT_API_DMZ=!NEQvV7JN8F'
                  '+hEmYU88hpJBfzTxR0Hbr4rq9p1C82nf7o2KMUWRIp4nimtkPr9AjxBzxcEb35v1cspso=; '
                  'TS01cc51d9'
                  '=013ee6d89b5bd90e8d3e846cf9a00c9c2f2d575034d075d6934001abd38cac93dd32c4eea7f0a28b5adea21b020318b4a8051f307ad1f6b5a516814330d3a2d7418e4ad56d'
    }

    retries = 3  # Maximum number of retries
    for attempt in range(retries):
        try:
            # Make request to first API
            response = requests.post(first_api_url, data=payload, headers=headers)
            response.raise_for_status()  # Raise an error for non-2xx responses
            access_token = response.json().get('access_token')
            return access_token
        except requests.exceptions.RequestException as e:
            # Handle request exceptions
            print(f"Attempt {attempt + 1} failed. Retrying...")
            time.sleep(2)  # Wait for 2 seconds before retrying
    else:
        # If all retries fail, send an error message
        error_message = "Access token API cannot be reached after multiple attempts."
        print(error_message)
        send_telegram_message(bot_token, group_chat_id, error_message)
        return None


def call_second_api(access_token):
    # # Generate a random last digit for the mobile number
    # random_digit = random.randint(0, 9)
    # mobile_number = "01301352705"[:-1] + str(random_digit)

    # Second API details with updated mobile number
    second_api_url = f"https://api-external.mynagad.com/ibas/api/ACSAPIService/GetMFSClientInfo?mobile=01301352705"
    headers = {
        'Authorization': f'Bearer {access_token}'
    }

    retries = 3  # Maximum number of retries
    for attempt in range(retries):
        try:
            # Make request to second API and measure response time
            start_time = time.time()
            response = requests.get(second_api_url, headers=headers)
            response.raise_for_status()  # Raise an error for non-2xx responses
            end_time = time.time()

            # Process response data
            print("Second API Response:", response.json())
            # Calculate response time
            response_time = round(end_time - start_time, 2)
            return "success", response_time
        except requests.exceptions.RequestException as e:
            # Handle request exceptions
            print(f"Attempt {attempt + 1} failed. Retrying...")
            time.sleep(2)  # Wait for 2 seconds before retrying
    else:
        # If all retries fail, send an error message
        error_message = "Get client API cannot be reached after multiple attempts."
        print(error_message)
        send_telegram_message(bot_token, group_chat_id, error_message)
        return "failure", None


def main():
    # Get access token
    access_token = get_access_token()

    if access_token:
        # Call second API with access token
        second_api_response, response_time = call_second_api(access_token)
        if second_api_response == "success":
            # Send message to Telegram indicating that the portal is up and running
            message = f"Tokenization API and get client API services status: Up and running! Response time :{response_time} seconds."
            response = send_telegram_message(bot_token, group_chat_id, message)
            if response.status_code == 200:
                print("Message sent successfully to Telegram group!")
            else:
                print("Failed to send message to Telegram group.")
                print(response.text)
        else:
            print("Failed to get second API response.")
    else:
        print("Failed to obtain access token.")


if __name__ == "__main__":
    main()