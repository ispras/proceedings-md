# Proceedings

Этот репозиторий содержит скрипт для автоматической конвертации `.md` в `.docx`,
соответствующий требованиям форматирования Трудов Института системного
программирования РАН.

## How to use:

```bash
git clone https://github.com/ispras/proceedings-md
cd proceedings-md
npm install
sudo apt-get install pandoc
```

Файл `sample/sample.md` содержит стандартный шаблон статьи для Трудов ИСП РАН,
представленный в `.md`-формате. Скрипт `src/main.js` выполняет конвертацию.

```
node src/main.js sample/sample.md sample/sample.docx
````

Если вы хотите использовать конвертер в контейнере, необходимо выполнить
следующие команды:
```
docker build -t proceedings-md .
./run-container.sh proceedings-md sample/sample.md sample/sample.docx
```

## Notes

Скрипт несколько сырой. Ошибки могут быть нечитаемыми. Некоторые версии Microsoft Word
ругаются на то, что документ повреждён, но все равно открывают его. Перед отправкой
документа рекомендуется открыть документ в Word, перепроверить форматирование, и сохранить
заново.

Happy researching!
