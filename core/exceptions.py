class FilePathException(Exception):
    def __init__(self, message="Проверьте указанный путь к файлу!"):
        self.message = message

    def __str__(self):
        return self.message


class CheckCalculationData(Exception):
    def __init__(self, message="Проверь данные для расчета."):
        super().__init__(message)
        # self.message=message

    # def __str__(self) -> str:
    #     return self.message


class AddPlsPolePathException(Exception):
    def __init__(self, message="Добавь ссылки к отчетам POLE"):
        self.message = message

    def __str__(self) -> str:
        return self.message
    

class FindInitialDataException(Exception):
    def __init__(self, message: str) -> None:
        super().__init__(message)
