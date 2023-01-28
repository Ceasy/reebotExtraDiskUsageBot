import requests
import cfg as c
import socket
import win32com.client as win32
import subprocess
import logging
import winshell

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
        return True
    except Exception as e:
        # print(f"error'{e}'")
        if e == "(-2147221005, 'Недопустимая строка с указанием класса', None, None)":
            print("Exel is not installed on this PC..")
        return False


def clear_recycle():
    try:
        winshell.recycle_bin().empty(confirm=False, show_progress=False, sound=False)
        print("Recycle bin cleared.")
        return True
    except Exception as e:
        # print(f"{e}")
        if e == "(-2147418113, 'Разрушительный сбой', None, None)":
            print("Recycle bin is empty.")
        return False


def reboot():
    print("Rebooting...")
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
        return True
    except Exception as e:
        logging.error("Error: ", e)
        return False


def main():
    # Save open Excel files
    if not save_files():
        print("Open files are not saved...")
    # Clear recycle bin
    if not clear_recycle():
        print("Nothing to clean in the recycle can...")
    # Check internet connection
    if check_internet_connection():
        # Check credentials
        if check_credentials():
            # Send a message to the Telegram bot
            if not message_bot():
                print("Log is not sent...")
    else:
        print("No internet connection...")
    # Reboot the PC
    reboot()


if __name__ == '__main__':
    # press Win+L to lock the PC
    subprocess.call("rundll32.exe user32.dll,LockWorkStation", shell=True)
    main()
