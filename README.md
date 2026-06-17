# Conference Editor

Conference Editor помогает быстро готовить бейджи и сертификаты для мероприятий по Excel-таблицам и PNG/JPG-шаблонам.

В проекте есть визуальный редактор разметки и пакетный генератор. Можно один раз настроить, где на макете должны стоять имя, город и название доклада, а потом автоматически собрать все картинки из таблицы.

## Что умеет

- Генерирует двухсторонние бейджи: русская сторона и английская сторона рядом в одном файле.
- Генерирует сертификаты по списку участников и названий докладов.
- Собирает сертификаты в A4 PDF: один сертификат на страницу.
- Собирает бейджи в A4 PDF: несколько бейджей на одном листе с автоматической раскладкой.
- Читает `.xlsx` файлы.
- Работает с PNG/JPG шаблонами.
- Позволяет двигать текстовые блоки мышкой прямо поверх шаблона.
- Синхронизирует настройки RU/EN сторон бейджа: координаты, размер рамки, шрифт, цвет и жирность.
- Умеет автоматически центрировать блоки по X, Y или сразу по двум осям.
- Поддерживает системные шрифты, включая ввод названия вроде `Times New Roman`.
- Позволяет менять размер шрифта, цвет, жирность и размер бокса текста.
- Запускает генерацию прямо из окна редактора.
- Собирается в приложения для Windows, macOS Intel и macOS Apple Silicon через GitHub Actions.

## Формат Excel

Для бейджей нужны колонки:

```text
ФИО на русском
ФИО на англе
Город на русском
город на англе
```

Для сертификатов нужны колонки:

```text
ФИО
название доклада
```

В папке `examples/data` лежат тестовые таблицы, а в `examples/templates` лежат примерные шаблоны.

## Локальный запуск

Установи зависимости:

```bash
python3 -m pip install -r requirements.txt
```

Открой визуальный редактор:

```bash
python3 layout_editor.py
```

В редакторе можно:

- открыть шаблон;
- выбрать тип документа: `badges` или `certificates`;
- настроить текстовые блоки;
- выбрать Excel;
- выбрать папку вывода;
- выбрать формат: `pdf`, `png`, `jpg` или `jpeg`;
- нажать `Запустить генерацию`.

По умолчанию редактор предлагает `pdf`, потому что это самый удобный режим для печати.

## Запуск генератора через терминал

Бейджи:

```bash
python3 sertificat.py badges \
  --excel examples/data/badges.xlsx \
  --template examples/templates/badge_template.png \
  --output examples/output/badges \
  --config layout.example.json \
  --format pdf
```

PDF будет сохранен как `badges_a4.pdf`.

Сертификаты:

```bash
python3 sertificat.py certificates \
  --excel examples/data/certificates.xlsx \
  --template examples/templates/certificate_template.png \
  --output examples/output/certificates \
  --config layout.example.json \
  --format pdf
```

PDF будет сохранен как `certificates_a4.pdf`.

Если вместо PDF нужны отдельные картинки, укажи `--format png`, `--format jpg` или `--format jpeg`.

## Сборка приложений через GitHub Actions

Workflow лежит в `.github/workflows/build-desktop.yml`.

Он собирает:

- `ConferenceEditor-Windows-x64`
- `ConferenceEditor-macOS-Intel`
- `ConferenceEditor-macOS-Apple-Silicon`

Сборка запускается вручную из вкладки **Actions** или автоматически при пуше git-тега вида `v0.1.0`.

macOS-приложения подписываются ad-hoc подписью, но не notarized. На macOS при первом запуске может понадобиться открыть приложение через правый клик → **Open**.

## Команды для публикации в пустой репозиторий

Ты просил не пушить автоматически, поэтому вот команды, которые можно выполнить самому:

```bash
cd /Users/macbookpro/Documents/fastapi-tutorial/smth
git init
git add .
git commit -m "Initial Conference Editor app"
git branch -M main
git remote add origin https://github.com/KeyGoddle/Conference_editor.git
git push -u origin main
```

Чтобы запустить сборку приложений через тег:

```bash
git tag v0.1.0
git push origin v0.1.0
```

После этого артефакты появятся во вкладке **Actions** внутри завершенного workflow run.
