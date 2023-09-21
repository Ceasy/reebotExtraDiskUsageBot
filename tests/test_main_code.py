import os
import unittest
from main import check_internet_connection,clear_temp_folder, TempFolderClearError
import tempfile


class TestMainCode(unittest.TestCase):

    def test_check_internet_connection(self):
        # Тест 1: Проверка, что функция возвращает True при успешном подключении к интернету
        self.assertTrue(check_internet_connection())

    # Тест 2: Проверка, что функция генерирует исключение InternetConnectionError при неудачной попытке подключения

    def test_check_credentials(self):
        from main_code_modified import check_credentials, CredentialsError

        # Тест 1: Проверка корректной работы функции при наличии действительных учетных данных
        self.assertTrue(check_credentials())


    def test_clear_temp_folder(self):
        # Тест 1: Проверка успешной очистки временной папки
        # Создаем уникальную тестовую временную папку и добавляем в нее несколько тестовых файлов
        test_temp_folder_path = tempfile.mkdtemp()
        test_file_paths = [os.path.join(test_temp_folder_path, 'test_file_' + str(i) + '.txt') for i in range(3)]
        for path in test_file_paths:
            with open(path, 'w') as f:
                f.write("test content")

        # Вызываем функцию clear_temp_folder с путем к тестовой папке
        clear_temp_folder()

        # Проверяем, что все файлы в тестовой папке были удалены
        test_files_deleted = all(not os.path.exists(path) for path in test_file_paths)
        self.assertTrue(test_files_deleted)

        # Тест 2: Проверка обработки ошибок функцией
        # Попробуем вызвать функцию clear_temp_folder с путем к несуществующей папке
        non_existent_folder_path = "/path/to/nonexistent/folder"
        with self.assertRaises(TempFolderClearError):
            clear_temp_folder(folder_path=non_existent_folder_path)


if __name__ == '__main__':
    unittest.main()