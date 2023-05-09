class FilePathException(Exception):
    def __init__(self, message="Проверьте указанный путь к файлу!"):
        self.message = message

    def __str__(self):
        return self.message
    