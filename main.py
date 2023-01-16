import psutil
import matplotlib.pyplot as plt
import requests
import cfg as c
import os
import socket
import time
import win32com.client as win32
from tqdm import tqdm
import subprocess
import logging

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


# def close_programs():
#     try:
#         # Close all open programs
#         os.system("taskkill /f /im *")
#         return True
#     except Exception as e:
#         logging.error("Error: ", e)
#         return False


def save_file():
    try:
        # Save open Excel files
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        for wb in excel.Workbooks:
            wb.Save()
            print("File ", wb.Name, " saved.")
        excel.Quit()

        # save open Word files
        word = win32.gencache.EnsureDispatch('Word.Application')
        for doc in word.Documents:
            doc.Save()
            print("File ", doc.Name, " saved.")
        word.Quit()

        # return True
    except Exception as e:
        logging.error("Error: ", e)
        # return False

    return True


def counter_reboot():
    print("Rebooting...")
    # Create a progress bar
    with tqdm(total=100) as pbar:
        for i in range(100):
            pbar.update(1)
            time.sleep(0.1)
    return True


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

    # Define a list of desired disk letters
    desired_disks = ["Z", "X", "Y"]

    # Check for all available drives on the system except the C drive
    for disk in psutil.disk_partitions():
        if os.path.exists(disk.device) and disk.device != "C:\\":
            # Check if the current disk is in the desired disks list
            if disk.device[0] in desired_disks:
                # Get the disk usage information
                disk_info = psutil.disk_usage(disk.device)
                disk_name = disk.device + " disk"
                disk_total = disk_info.total / (1024.0 ** 3)
                disk_used = disk_info.used / (1024.0 ** 3)
                disk_free = disk_info.free / (1024.0 ** 3)
                # Plot the disk usage information
                labels = ['Used', 'Free']
                sizes = [disk_used, disk_free]
                plt.pie(sizes, labels=labels, autopct='%1.1f%%')
                plt.title(disk_name + ' Usage')
                # Add text to the image showing the total, used, and free disk space in GB
                plt.text(0, -1, f'Total: {disk_total:.2f} GB\nUsed: {disk_used:.2f} GB\nFree: {disk_free:.2f} GB',
                         fontsize=10)
                # Save the plot to an image file
                plt.savefig(disk_name + '_usage.png')

                # Check if the image file was created successfully before attempting to send it via Telegram API
                if os.path.exists(disk_name + '_usage.png'):
                    # Send the message with the disk usage graph attached using the Telegram API
                    try:
                        response = requests.post(f"https://api.telegram.org/bot{bot_token}/sendPhoto",
                                                 params={"chat_id": chat_id},
                                                 files={"photo": open(disk_name + '_usage.png', "rb")},
                                                 data={"caption": message})
                    except Exception as e:
                        logging.error("Error: ", e)

                    # Delete the disk usage image file
                    os.remove(disk_name + '_usage.png')
    return True


def main():
    # Check internet connection
    if not check_internet_connection():
        return

    # Check credentials
    if not check_credentials():
        return

    # Close all open programs
    # if not close_programs():
    #     return

    # Save open Excel files
    if not save_file():
        return

    # Reboot the PC
    if not counter_reboot():
        return

    # Send a message to the Telegram bot
    if not message_bot():
        return

    # Reboot the PC
    # subprocess.call("shutdown /r /t 1")
    subprocess.call("shutdown /f /r /t 0", shell=True)
    print("bb...")


if __name__ == '__main__':
    main()
