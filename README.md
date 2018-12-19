# Описание

Програма для создания XLS файла из CSV - плагировщик расписания для ISU ITMO.

## Установка

Используйте пакетный менеджер [pip3](https://pip.pypa.io/en/stable/) для установки всех необходимых пакетов.
```bash
git clone https://github.com/TexnoMann/csvToXlsTable.git
```

```bash
pip3 install -r requirements.txt
```

## Запуск
Для запуска на Ubuntu:
```bash
./run.sh [Путь до csv файла] [Путь сохранения xls файла][Список групп при желании]
```
## Примечание
Открывать можно файлы называя как абсолютный путь, так и относительный. Расширение файлу давать не обязательно, в случае отсутствия его в имени файла программа самостоятельно даст расширение файлу:

Пример использования абсолютного пути:
```bash
./run.sh /home/user/csvToXls/program/text.csv /home/user/csvToXls/program/text.xls P3138 P3134
```

## License
[MIT](https://choosealicense.com/licenses/mit/)
