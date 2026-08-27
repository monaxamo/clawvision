# ClawVision 1.0: визуальные саммари для OpenClaw без облаков

После релиза Codex `$visualize` я задумался: а почему бы такой же локальный skill не сделать для OpenClaw? Так появился **ClawVision** — инструмент, который превращает историю сессии OpenClaw в компактную HTML-карточку с экспортом в Markdown и PowerPoint.

## Что умеет

- **4 таба:** главный вывод, формат диалога, что построено, следующие шаги.
- **Схема потока** и **метрические карточки**.
- **Переключатель языков EN/RU/ZH**.
- **Светлая/тёмная тема**.
- **Кнопки экспорта** Markdown и PowerPoint прямо в карточке.
- **PNG-скриншоты**: общий и по одному на каждую вкладку.

## Зачем

Иногда за часовой диалог с ассистентом теряется общая картина. ClawVision даёт одностраничный артефакт, который можно:
- сохранить в daily notes,
- отправить в Telegram,
- вставить в презентацию,
- опубликовать на clawhub.

## Как работает

Skill использует `sessions_history` и локальную модель через `node_inference`, чтобы извлечь структуру разговора. Затем локальный Python-скрипт на Playwright и `python-pptx` рендерит HTML, PNG, Markdown и PPTX.

```bash
python scripts/generate_visual.py \
  --summary summary.json \
  --output ./out \
  --png --md --pptx \
  --lang ru
```

## Пример

Вот как выглядит карточка для реальной сессии по запуску vibe.egetech.ru:

![Коллаж вкладок](session_vibe_setup_ru_2026-08-27_collage.png)

## Технический стек

- Python 3.10+
- Playwright + Chromium
- python-pptx
- Self-contained HTML: inline CSS и JS, нет внешних зависимостей.

## Будущее

**ClawVision 1.0** уже работает. **ClawVision 2.0** добавит аналитику: статистику сообщений, использование инструментов, извлечение тем и сущностей, CSS-only графики и автоматические инсайты.

## Ссылки

- GitHub: https://github.com/monaxamo/clawvision
- clawhub: https://clawhub.ai/skills/clawvision
- Релиз 1.0.0: https://github.com/monaxamo/clawvision/releases/tag/v1.0.0

---

*Автор: Maximius. Лицензия MIT.*
