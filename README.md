# Proceedings Markdown

**Осторожно, вайбкод**

Экспериментальная переработка `proceedings-md`, обкатанная на статье SeqSan.
Конвертирует Markdown в DOCX для «Трудов ИСП РАН», при необходимости экспортирует PDF.
Перед отправкой статьи проверьте готовый документ в Word.

Пример: `sample/sample.md`. Правила разметки: [FORMAT.md](FORMAT.md).

## Сборка в контейнере

Нужен Docker. Node.js, TeX и LibreOffice на компьютере устанавливать не нужно.
Из корня репозитория:

```sh
docker build -t proceedings-md-build .
docker run --rm --user "$(id -u):$(id -g)" \
  -v "$PWD:/work" proceedings-md-build \
  /work/sample/sample.md /work/build/sample.docx --figures --pdf
```

Результат: `build/sample.docx` и `build/sample.pdf`. Без `--pdf` собирается только
DOCX. Без `--figures` используются готовые PNG/SVG из `figures/`.
PDF экспортирует LibreOffice: это удобное превью, его разбивка на страницы
может отличаться от Microsoft Word.

Образ содержит сборщик и шаблон. После изменения `tools/`, `resources/` или
зависимостей образ нужно пересобрать. Правки статьи и рисунков не требуют
пересборки образа. Базовый образ закреплён по digest, пакеты Debian
берутся из снимка на 1 октября 2026 года, зависимости npm — из lockfile.

## Локальная сборка

Нужны Node.js 22 или новее, XeLaTeX и `pdftocairo` из Poppler.
Для PDF также нужен `soffice` из LibreOffice.

Debian/Ubuntu:

```sh
sudo apt-get update
sudo apt-get install texlive-xetex texlive-latex-extra \
  texlive-fonts-recommended poppler-utils libreoffice-writer fonts-liberation
npm ci --ignore-scripts
npm run build
npm run build -- --figures --pdf
```

На macOS подойдут MacTeX и Poppler. Для одинакового окружения на разных
компьютерах используйте контейнер.

```sh
npm run check                       # метаданные, ссылки, наличие рисунков
npm run typecheck                   # типы сборщика
npm test                            # регрессионные тесты
npm run format:check                # оформление кода и инструкции
npm run build -- input.md output.docx
```

`make`, `make check`, `make test` и `make pdf` вызывают соответствующие команды.
Пути рисунков считаются относительно входного Markdown, а не рабочего каталога.
Шаблон и шрифты находятся относительно кода сборщика.

## Что делает сборщик

1. Разбирает Markdown и YAML, проверяет метаданные и изображения.
2. Нумерует разделы, подписи и теоремы по порядку определений. Проверяет ссылки
   и библиографию, включая ссылки вперёд.
3. При `--figures` пересобирает рисунки из TikZ в PNG/SVG через общий PDF.
4. Рендерит уникальные формулы одним документом XeLaTeX. В DOCX формулы
   вставляются как PNG с разрешением 600 dpi и выравниванием по базовой линии.
5. Заполняет шаблон статьи и проверяет XML, связи файлов и идентификаторы рисунков.
6. При `--pdf` экспортирует DOCX через отдельный профиль LibreOffice.

Формулы не являются редактируемыми уравнениями Word. Рисунки вставляются
с PNG для совместимости и SVG для Word с поддержкой векторных изображений.
Отсутствующие ссылки и изображения останавливают сборку.

Кэш формул хранится в `.cache/formulas/`, рисунков — в `.cache/figures/`. Изменение формулы, исходника рисунка, преамбулы или версии
XeLaTeX меняет ключ кэша. `make clean` удаляет кэш и результаты сборки.
Исходники не дополняются промежуточными JSON или документами.
При неизменном окружении повторная сборка DOCX даёт одинаковые байты.
PDF содержит метку времени экспорта.

## Проверки

`npm run typecheck`, `npm test` и `npm run format:check` проверяют сборщик.
`npm run check -- sample/sample.md` проверяет пример.
Полная проверка включает `npm run build -- --figures --pdf`.

## Рисунки и шаблон

`figures/source/` рядом со входным Markdown содержит исходники TikZ. Готовые SVG и PNG находятся в
`figures/` и позволяют собирать статью без пересборки графики.
При `--figures` пересобираются рисунки, используемые статьёй. Результат
хранится в кэше, исходные файлы и готовые копии в Git не изменяются.
После редактирования TikZ используйте `--figures`. Команда `npm run figures -- sample/sample.md`
обновляет готовые PNG/SVG в `figures/` для просмотра Markdown и быстрой сборки.

Шаблон в `resources/isp-reference.docx` задаёт оформление сборника.
Реквизиты образца и жёлтые выделения сохранены для редакционной подготовки.
Реквизиты публикации можно добавить в поля цитирования во frontmatter.
ORCID печатается как указан в метаданных, включая нулевое значение шаблона.

Происхождение кода, шаблона и шрифтов описано в [resources/PROVENANCE.md](resources/PROVENANCE.md).
