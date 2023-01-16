import unittest
import mock
from main import message_bot


class TestMessageBot(unittest.TestCase):

    @mock.patch('main.check_internet_connection')
    @mock.patch('main.requests.post')
    def test_message_bot_success(self, mock_post, mock_internet):
        # Arrange
        mock_internet.return_value = True
        mock_post.return_value.status_code = 200

        # Act
        result = message_bot()

        # Assert
        self.assertEqual(result, True)
        mock_post.assert_called_once()

    @mock.patch('main.check_internet_connection')
    def test_message_bot_no_internet(self, mock_internet):
        # Arrange
        mock_internet.return_value = False

        # Act
        result = message_bot()

        # Assert
        self.assertEqual(result, None)

    @mock.patch('main.check_internet_connection')
    @mock.patch('main.requests.post')
    def test_message_bot_api_error(self, mock_post, mock_internet):
        # Arrange
        mock_internet.return_value = True
        mock_post.return_value.status_code = 400

        # Act
        result = message_bot()

        # Assert
        self.assertEqual(result, None)


if __name__ == '__main__':
    unittest.main()
