import unittest
import requests_mock
import cfg
import two
import mock
from unittest.mock import MagicMock
import win32com.client as win32


class TestEReebot(unittest.TestCase):
    @mock.patch('socket.create_connection')
    def test_check_internet_connection(self, mock_create_connection):
        # Test case with internet connection
        mock_create_connection.return_value = True
        self.assertTrue(two.check_internet_connection())

        # Test case without internet connection
        mock_create_connection.side_effect = OSError
        self.assertFalse(two.check_internet_connection())

    @mock.patch('cfg.TOKEN')
    @mock.patch('cfg.chat_id')
    def test_check_credentials(self, mock_token, mock_chat_id):
        # Test case with valid credentials
        mock_token.return_value = '1234567890'
        mock_chat_id.return_value = '1234567890'
        self.assertTrue(two.check_credentials())

        # Test case with invalid credentials
        mock_token.return_value = None
        mock_chat_id.return_value = None
        self.assertFalse(two.check_credentials())

    @mock.patch('win32com.client.gencache.EnsureDispatch')
    def test_save_files(self, mock_EnsureDispatch):
        # Test case with Excel installed
        mock_EnsureDispatch.return_value = True
        self.assertTrue(two.save_files())

        # Test case with Excel not installed
        mock_EnsureDispatch.side_effect = Exception
        self.assertFalse(two.save_files())

    @mock.patch('winshell.recycle_bin')
    def test_clear_recycle(self, mock_recycle_bin):
        # Test case with recycle bin cleared
        mock_recycle_bin.empty.return_value = True
        self.assertTrue(two.clear_recycle())

        # Test case with recycle bin not cleared
        mock_recycle_bin.empty.return_value = False
        self.assertFalse(two.clear_recycle())



    # class TestEReebot(unittest.TestCase):
    #     def test_save_files(self):
    #         mock_excel = MagicMock()
    #         mock_word = MagicMock()
    #         win32.gencache.EnsureDispatch = MagicMock(side_effect=[mock_excel, mock_word])
    #         mock_excel.Workbooks = [MagicMock(), MagicMock()]
    #         mock_word.Documents = [MagicMock(), MagicMock()]
    #         self.assertTrue(two.save_files())


if __name__ == '__main__':
    unittest.main()
