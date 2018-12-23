# Описание

Програма для создания XLS файла из CSV - планировщик расписания для ISU ITMO.

## Установка

Используйте пакетный менеджер [pip3](https://pip.pypa.io/en/stable/) для установки всех необходимых пакетов.
```bash
git clone https://github.com/TexnoMann/csvToXlsTable.git
```

```bash
pip3 install -r requirements.txt
```

## Запуск
Для запуска программы на Ubuntu:
```bash
./run.sh [Путь до csv файла] [Путь сохранения xls файла] [-g|--groups Список групп]
```
## Примечание
Открыть файл можно использовав, как абсолютный, так и относительный путь. Расширение файлу давать не обязательно, в случае его отсутствия в имени файла программа самостоятельно даст расширение файлу. Для автоматического добавления в план списка групп необходимо написать:
```bash
-g [Список групп]
```

Пример с использованием абсолютного пути:
```bash
./run.sh /home/user/csvToXls/program/text.csv /home/user/csvToXls/program/text.xls -g P3138 P3134
```

## License
[MIT](https://choosealicense.com/licenses/mit/)
