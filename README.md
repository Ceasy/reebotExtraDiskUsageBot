# Description.

eReebot is a Python script that allows you to reboot your PC and send a message to a Telegram bot, notifying that the PC has been rebooted.

The script performs the following actions:

Saves open Excel and Word files.
Clears the recycle bin.
Checks for internet connection and Telegram bot credentials.
Sends a message to the Telegram bot with the hostname of the PC and a notification that the PC has been rebooted.
Reboots the PC.
This script can be useful for those who need to reboot their PC regularly and want to be notified when the reboot is complete. The script can also be useful for those who want to keep their PC clean and clear the recycle bin regularly.

To use the script, you need to have Python installed on your PC and configure the Telegram bot token and chat ID in the cfg.py file.

The script can be further optimized by running functions in parallel and using non-blocking calls.

Feel free to use and modify the script to suit your needs.

#Installation
Download or clone the repository
Create a bot in Telegram and get the token
Replace bot_token and chat_id in the cfg.py file with your own values
Run the script
#Usage
Run the script
Wait for the script to complete
You will receive a message in Telegram that the PC has been rebooted
Contributing
Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

main.py - actual main program
old_main_1.py - old main program (version 1). This version worked stable, without additional functions, but with sending screenshots.
old_main_2.py - dublicat of main program. This version works now, it has implemented additional functions such as connection check, API check .
setup.py - To convert from py to cython format.
test.py - To test the program.
test2.py - To test the program.

#License
MIT

#Additional
This script is not for commercial use and is provided as is. Use at your own risk. The author will not be held liable for any damage caused by this script.
