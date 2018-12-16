# Описание

Програма для создания XLS файла из CSV.

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
./run.sh [Путь до csv файла] [Путь сохранения xls файла]
```
## Примечание
На данный момент реализовано открытие файлов только с использованием абсолютного пути:  
Пример:
```bash
./run.sh /home/user/program/text.csv /home/user/program/text.xls
```

## License
[MIT](https://choosealicense.com/licenses/mit/)
