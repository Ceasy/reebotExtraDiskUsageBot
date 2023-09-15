import subprocess

import pythoncom
import requests
import cfg as c
import socket
import win32com.client as win32
import logging
import concurrent.futures
import ctypes

logging.basicConfig(filename='eReebot.log', filemode='w', format='%(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger()
pythoncom.CoInitialize()


def check_internet_connection():
    try:
        socket.create_connection(("www.google.com", 80))
        return True
    except OSError as e:
        logging.error(f"Error checking internet connection: {e}")
        return False


def check_credentials():
    try:
        bot_token = getattr(c, 'TOKEN', None)
        chat_id = getattr(c, 'CHAT_ID', None)
        if not bot_token or not chat_id:
            raise ValueError("Invalid bot token or chat ID")
        return True
    except ValueError as e:
        logging.error(f"Error checking credentials: {e}")
        return False


def save_files():
    try:
        pythoncom.CoInitialize()
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
        return True
    except win32.pywintypes.com_error as e:
        logging.error(f"Error saving files: {e}")
        return False
    finally:
        pythoncom.CoUninitialize()


def clear_recycle():
    try:
        SHERB_NOCONFIRMATION = 0x00000001
        SHERB_NOPROGRESSUI = 0x00000002
        SHERB_NOSOUND = 0x00000004
        SHERB_REASON = SHERB_NOCONFIRMATION | SHERB_NOPROGRESSUI | SHERB_NOSOUND

        result = ctypes.windll.shell32.SHEmptyRecycleBinA(None, None, SHERB_REASON)

        if result == 0:
            logger.info("Recycle bin cleared successfully!")
        else:
            logger.error(f"Failed to empty Recycle Bin. Error code: {result}")

    except Exception as e:
        logger.error(f"Error clearing recycle bin: {e}")


def message_bot():
    # Check internet connection
    if not check_internet_connection():
        return

    # Get the hostname of the PC
    hostname = socket.gethostname()

    # Define the bot's token and the chat ID of the recipient (in this case, yourself)
    bot_token = c.TOKEN
    chat_id = c.CHAT_ID

    # Define the message text
    message = f"☝⚠️ {hostname} - Экстренно перезагружен!!!"

    try:
        response = requests.post(f"https://api.telegram.org/bot{bot_token}/sendMessage",
                                 params={"chat_id": chat_id},
                                 json={"text": message})
        response.raise_for_status()
        return True
    except requests.exceptions.RequestException as e:
        logging.error(f"Error messaging bot: {e}")
        return False


def main():
    # press Win+L to lock the PC
    subprocess.Popen("rundll32.exe user32.dll,LockWorkStation", shell=True)
    with concurrent.futures.ThreadPoolExecutor() as executor:
        # Save open Excel files
        executor.submit(save_files)
        # Clear recycle bin
        executor.submit(clear_recycle)
        # Check internet connection
        internet_connection = check_internet_connection()
        if internet_connection:
            # Check credentials
            if check_credentials():
                # Send a message to the Telegram bot
                executor.submit(message_bot)
        else:
            print("No internet connection...")
    # Reboot the PC
    subprocess.Popen("shutdown /f /r /t 0", shell=True)


if __name__ == '__main__':
    main()
