import requests
import cfg as c
import socket
import time
import win32com.client as win32
from tqdm import tqdm
import subprocess
import logging
import threading

logging.basicConfig(filename='eReebot.log', level=logging.ERROR)


def check_internet_connection():
    try:
        socket.create_connection(("www.google.com", 80))
        return True
    except OSError:
        pass
    return False


def check_credentials():
    try:
        bot_token = c.TOKEN
        chat_id = c.chat_id
        if not bot_token or not chat_id:
            raise ValueError("Invalid bot token or chat ID")
        return True
    except Exception as e:
        logging.error("Error: ", e)
        return False


def save_files():
    try:
        # Save open Excel files
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        for wb in excel.Workbooks:
            wb.Save()
            print(f"File {wb.Name} saved.")
        excel.Quit()

        # save open Word files
        word = win32.gencache.EnsureDispatch('Word.Application')
        for doc in word.Documents:
            doc.Save()
            print(f"File {doc.Name} saved.")
        word.Quit()
    except Exception as e:
        print(e)
    return True


def counter_reboot():
    print("Rebooting...")
    thread = threading.Thread(target=countdown)
    thread.start()


def countdown():
    with tqdm(total=100) as pbar:
        for i in range(100):
            pbar.update(1)
            time.sleep(0.1)
    # Send a message to the Telegram bot
    if not message_bot():
        return
    subprocess.call("shutdown /f /r /t 0", shell=True)
    print("bb...")


def message_bot():
    # Check internet connection
    if not check_internet_connection():
        return

    # Get the hostname of the PC
    hostname = socket.gethostname()

    # Define the bot's token and the chat ID of the recipient (in this case, yourself)
    bot_token = c.TOKEN
    chat_id = c.chat_id

    # Define the message text
    message = f"☝⚠️ {hostname} - Экстренно перезагружен!!!"

    try:
        response = requests.post(f"https://api.telegram.org/bot{bot_token}/sendMessage",
                                 params={"chat_id": chat_id},
                                 json={"text": message})
    except Exception as e:
        logging.error("Error: ", e)
    return True


def main():
    # Check internet connection
    if not check_internet_connection():
        return

    # Check credentials
    if not check_credentials():
        return

    # Save open Excel files
    if not save_files():
        return

    # Reboot the PC
    counter_reboot()

    # Send a message to the Telegram bot
    # if not message_bot():
    #     return


if __name__ == '__main__':
    main()

# Change the code, make the counter function in a separate thread, and that at 100% it calls the restart function of the PC.
