import unittest
import main


class TestYourScript(unittest.TestCase):
    def test_check_internet_connection(self):
        result = main.check_internet_connection()
        self.assertTrue(result)

    def test_check_credentials(self):
        result = main.check_credentials()
        self.assertTrue(result)

    def test_close_programs(self):
        result = main.close_programs()
        self.assertTrue(result)

    def test_save_file(self):
        result = main.save_file()
        self.assertTrue(result)

    def test_counter_reboot(self):
        result = main.counter_reboot()
        self.assertTrue(result)

    def test_message_bot(self):
        result = main.message_bot()
        self.assertTrue(result)
