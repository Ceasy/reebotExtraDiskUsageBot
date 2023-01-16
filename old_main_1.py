import psutil
import matplotlib.pyplot as plt
import requests
import cfg as c
import os
import socket
import time
import win32com.client as win32
from tqdm import tqdm


def save_file():
    # Save open Excel files
    try:
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        for wb in excel.Workbooks:
            wb.Save()
            print("File ", wb.Name, " saved.")
        excel.Quit()
    except Exception as e:
        print("Error: ", e)

    # save open Word files
    try:
        word = win32.gencache.EnsureDispatch('Word.Application')
        for doc in word.Documents:
            doc.Save()
            print("File ", doc.Name, " saved.")
        word.Quit()
    except Exception as e:
        print("Error: ", e)

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
    # Get the hostname of the PC
    hostname = socket.gethostname()

    # Define the bot's token and the chat ID of the recipient (in this case, yourself)
    bot_token = c.TOKEN
    chat_id = c.chat_id

    # Define the message text
    message = f"☝⚠️ {hostname} - Экстренно перезагружен!!!"

    # Check for disks Z, X, Y
    disks = ['Z', 'X', 'Y']
    for disk in disks:
        if os.path.exists(disk + ':\\'):
            # Get the disk usage information
            disk_info = psutil.disk_usage(disk + ':\\')
            disk_name = disk + " disk"
            disk_total = disk_info.total / (1024.0 ** 3)
            disk_used = disk_info.used / (1024.0 ** 3)
            disk_free = disk_info.free / (1024.0 ** 3)
            # Plot the disk usage information
            labels = ['Used', 'Free']
            sizes = [disk_used, disk_free]
            plt.pie(sizes, labels=labels, autopct='%1.1f%%')
            plt.title(disk_name + ' Usage')

            # Add text to the image showing the total, used, and free disk space in GB
            plt.text(0, -1, f'Total: {disk_total:.2f} GB\nUsed: {disk_used:.2f} GB\nFree: {disk_free:.2f} GB', fontsize=10)

            # Save the plot to an image file
            plt.savefig(disk_name + '_usage.png')

            # Send the message with the disk usage graph attached using the Telegram API
            response = requests.post(f"https://api.telegram.org/bot{bot_token}/sendPhoto",
                                     params={"chat_id": chat_id},
                                     files={"photo": open(disk_name + '_usage.png', "rb")},
                                     data={"caption": message})

            # Delete the disk usage image file
            os.remove(disk_name + '_usage.png')

    return True


def reboot():
    # reboot the computer
    os.system("shutdown /f /r /t 0")
    print("bb...")


if __name__ == '__main__':
    if counter_reboot() and message_bot() and save_file():
        reboot()
