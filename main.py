import shutil
import subprocess
import time
import pythoncom
import requests
import cfg as c
import socket
import logging
import concurrent.futures
import ctypes
import os
import glob


class InternetConnectionError(Exception):
    """Raised when there is an error with the internet connection."""
    pass


class CredentialsError(Exception):
    """Raised when there is an error with the bot credentials."""
    pass


class FileSaveError(Exception):
    """Raised when there is an error saving files."""
    pass


class RecycleBinError(Exception):
    """Raised when there is an error clearing the recycle bin."""
    pass


class OfficeFolderClearError(Exception):
    """Raised when there is an error clearing office folders."""
    pass


class TempFolderClearError(Exception):
    """Raised when there is an error clearing the temp folder."""
    pass


class OneCCacheClearError(Exception):
    """Raised when there is an error clearing the 1C cache."""
    pass


log_directory_path = os.path.join(os.environ['LOCALAPPDATA'], 'eReboot')
os.makedirs(log_directory_path, exist_ok=True)

log_file_path = os.path.join(log_directory_path, 'eReebot.log')

logging.basicConfig(filename=log_file_path, level=logging.INFO,
                    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')

logger = logging.getLogger()
pythoncom.CoInitialize()


def check_internet_connection():
    connection = None
    try:
        connection = socket.create_connection(("www.google.com", 80))
        return True
    except OSError as e:
        logging.exception("Error checking internet connection")
        raise InternetConnectionError("Unable to connect to the internet.") from e
    finally:
        if connection:
            connection.close()


def check_credentials():
    try:
        bot_token = getattr(c, 'TOKEN', None)
        chat_id = getattr(c, 'CHAT_ID', None)
        if not bot_token or not chat_id:
            raise ValueError("Invalid bot token or chat ID")
        return True
    except ValueError as e:
        logging.exception("Error checking credentials")
        raise CredentialsError("Invalid bot credentials.") from e


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
        logging.exception("Error clearing recycle bin")
        raise RecycleBinError("Error clearing the recycle bin.") from e


def clear_office_folders():
    folders_to_clear = [
        os.path.join(os.getenv('APPDATA'), 'Microsoft', 'Word'),
        os.path.join(os.getenv('APPDATA'), 'Microsoft', 'Excel'),
        os.path.join(os.getenv('APPDATA'), 'Microsoft', 'PowerPoint'),
        os.path.join(os.getenv('APPDATA'), 'Microsoft', 'Outlook'),
    ]

    for folder in folders_to_clear:
        if os.path.exists(folder):
            files = glob.glob(os.path.join(folder, '*'))
            for f in files:
                try:
                    if os.path.isfile(f):
                        os.remove(f)
                    elif os.path.isdir(f):
                        shutil.rmtree(f)
                except Exception as e:
                    logging.exception(f"Error clearing folder {folder}")
                    raise OfficeFolderClearError(f"Error clearing folder {folder}.") from e
        else:
            logger.info(f"Folder {folder} does not exist, skipping...")


def clear_temp_folder(folder_path=None):
    if not folder_path:
        folder_path = os.getenv('TEMP')

    if folder_path:
        logger.info(f"Clearing the folder: {folder_path}")
        files = glob.glob(os.path.join(folder_path, '*'))
        for f in files:
            try:
                if os.path.isfile(f):
                    os.remove(f)
                elif os.path.isdir(f):
                    shutil.rmtree(f)
            except PermissionError:
                logging.warning(f"Permission error while trying to delete {f}. Skipping...")
            except Exception as e:
                logging.exception("Error clearing folder")
                raise TempFolderClearError("Error clearing the folder.") from e


def clear_1c_cache():
    try:
        user_profile_path = os.environ['USERPROFILE']
        roaming_path = os.path.join(user_profile_path, 'AppData', 'Roaming', '1C', '1cv8')
        local_path = os.path.join(user_profile_path, 'AppData', 'Local', '1C', '1cv8')

        paths_to_clear = [roaming_path, local_path]

        for path in paths_to_clear:
            if os.path.exists(path):
                shutil.rmtree(path)
                logger.info(f"Cleared 1C cache at {path}")
            else:
                logger.warning(f"1C cache folder at {path} does not exist")
    except Exception as e:
        logging.exception("Error clearing 1C cache")
        raise OneCCacheClearError("Error clearing the 1C cache.") from e


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
    message = f"☝⚠️ {hostname} - Экстренно перезагружается..."

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
        # Check internet connection
        internet_connection = check_internet_connection()
        if internet_connection:
            # Check credentials
            if check_credentials():
                # Send a message to the Telegram bot
                executor.submit(message_bot)
        else:
            print("No internet connection...")
        time.sleep(2)
        # Clear recycle bin
        executor.submit(clear_recycle)
        # Clear Office folders
        executor.submit(clear_office_folders)
        # Clear 1C cache folders
        executor.submit(clear_1c_cache)
        # Clear TEMP folder
        executor.submit(clear_temp_folder)
    # Reboot the PC
    subprocess.Popen("shutdown /f /r /t 0", shell=True)


if __name__ == '__main__':
    main()
