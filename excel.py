import pandas as pd


class Excel_read:
    """Класс работы с эксель файлами."""

    def __init__(self, file_name: str) -> None:
        """Конструктор класс Exel_read.

        Args:
             file_name: Название файла.
        """

        self.__file_name = f'{file_name}.xlsx'
        self.__file_work = pd.read_excel(self.__file_name, sheet_name= 'Sheet1')
        self.__file_work.drop_duplicates(inplace= True)

        self.__now_date = pd.Timestamp.now()

    def warranty_sort(self) -> pd.DataFrame:
        """Метод сортировки данных по гарантии.

        Returns:
              new_data: Данные по гарантии.
        """

        current_warranty = []
        current_warranty_id = []
        clinic_id =[]

        try:
            self.__file_work['warranty_until'] = pd.to_datetime(self.__file_work['warranty_until'], format= 'mixed', errors= 'coerce')

            for index, row in self.__file_work.iterrows():
                if row['warranty_until'] >= self.__now_date:
                    clinic_id.append(row['clinic_id'])
                    current_warranty_id.append(row['device_id'])
                    current_warranty.append(row['warranty_until'])

            new_data = pd.DataFrame({
                'clinic_id': clinic_id,
                'device_id': current_warranty_id,
                'warranty_until': current_warranty
            })

            with pd.ExcelWriter(self.__file_name, mode= 'a', engine= 'openpyxl') as writer:
                new_data.to_excel(writer, header= True, sheet_name= 'Warranty')

        except Exception:
            print('Произошла ошибка при сортировке данных по гарантии')

        else:
            return new_data

    def top_issues(self) -> pd.DataFrame:
        """Метод поиска топ 100 клиник с наибольшим количеством проблем.

        Returns:
              top_clinic: топ 100 клиник по количеству проблем.
        """

        new_top_clinic = []
        try:
            sort_file = self.__file_work.sort_values('issues_reported_12mo', ascending= False)

            for index, row in sort_file.head(100).iterrows():
                new_top_clinic.append([row['clinic_id'], row['clinic_name'], row['issues_reported_12mo']])

            top_clinic = pd.DataFrame(new_top_clinic, columns= ['clinic_id', 'clinic_name', 'issues_reported_12mo'])

            with pd.ExcelWriter(self.__file_name, mode= 'a', engine= 'openpyxl') as writer:
                top_clinic.to_excel(writer, header= True, sheet_name= 'Top_100_clinic')

        except Exception:
            print('Произошла ошибка при поиске клиник с проблемами')

        else:
            return top_clinic

    def calibration_report(self) -> str:
        """Метод создания отчёта о калибровке.

        Returns:
              отчёт по датам калибровки.
        """

        self.__file_work['last_calibration_date'] = pd.to_datetime(self.__file_work['last_calibration_date'], format='mixed', errors='coerce')
        self.__file_work['install_date'] = pd.to_datetime(self.__file_work['install_date'], format='mixed', errors='coerce')

        count_normal = 0
        count_error = 0
        count_none = 0

        for index, row in self.__file_work.iterrows():
            if self.__now_date >= row['last_calibration_date'] >= row['install_date']:
                count_normal += 1
            elif row['last_calibration_date'] < row['install_date']:
                count_error += 1
            else:
                count_none += 1

        return (
            f'Количество удовлетворяющих условия дат калибровки: {count_normal}\n'
            f'Количество ошибочных дат: {count_error}\n'
            f'Количество отсутствующих данных: {count_none}\n'
            )

    def new_master_table(self) -> None:
        """Метод создания сводной таблицы."""

        new_table = pd.merge(self.top_issues(), self.warranty_sort(), on= 'clinic_id', how= 'outer')

        try:
            with pd.ExcelWriter(self.__file_name, mode= 'a', engine= 'openpyxl') as writer:
                new_table.to_excel(writer, header= True, sheet_name='master_table')

        except Exception:
            print('\nПроизошла ошибка!')