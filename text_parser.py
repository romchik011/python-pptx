import re
from dataclasses import dataclass, field
from typing import List, Optional

BULLET_RE = re.compile(r"^\s*(?:[-*•–—]|\d+[.)])\s+(.*)$")


@dataclass
class SlideContent:
    title: str
    bullets: List[str] = field(default_factory=list)
    paragraphs: List[str] = field(default_factory=list)


@dataclass
class Deck:
    title: str
    subtitle: Optional[str] = None
    slides: List[SlideContent] = field(default_factory=list)


def _looks_like_title(line: str) -> bool:
    line = line.strip()
    return 0 < len(line) <= 80 and not line.endswith((".", "!", "?", ";", ","))


def _fill_slide(slide: SlideContent, lines: List[str]) -> None:
    """Раскладывает строки по пунктам списка и абзацам."""
    paragraph: List[str] = []

    def flush():
        if paragraph:
            slide.paragraphs.append(" ".join(paragraph).strip())
            paragraph.clear()

    for raw in lines:
        if not raw.strip():
            flush()
            continue
        m = BULLET_RE.match(raw)
        if m:
            flush()
            slide.bullets.append(m.group(1).strip())
        else:
            paragraph.append(raw.strip())
    flush()


def parse_text(text: str, default_title: str = "Презентация") -> Deck:
    """
    Превращает обычный текст в структуру презентации.

    Формат (всё необязательно):
        # Название презентации
        Подзаголовок (строка сразу после названия)

        ## Заголовок слайда
        - пункт списка
        - ещё пункт
        Обычный абзац текста.

    Если заголовков "##" нет, каждый блок текста, отделённый пустой строкой,
    становится отдельным слайдом (короткая первая строка — заголовок).
    """
    text = text.replace("\r\n", "\n").strip()
    if not text:
        raise ValueError("Текст пустой — введите тему или план презентации.")

    lines = text.split("\n")
    deck = Deck(title=default_title)

    # Название презентации: "# ..."
    if lines and lines[0].lstrip().startswith("# ") and not lines[0].lstrip().startswith("## "):
        deck.title = lines[0].lstrip()[2:].strip() or default_title
        lines = lines[1:]
        # подзаголовок — непустая строка сразу после названия, если это не заголовок/список
        if lines and lines[0].strip() and not lines[0].lstrip().startswith("#") and not BULLET_RE.match(lines[0]):
            deck.subtitle = lines[0].strip()
            lines = lines[1:]

    has_headings = any(l.lstrip().startswith("## ") for l in lines)

    if has_headings:
        current: Optional[SlideContent] = None
        buffer: List[str] = []
        for line in lines:
            if line.lstrip().startswith("## "):
                if current is not None:
                    _fill_slide(current, buffer)
                    deck.slides.append(current)
                current = SlideContent(title=line.lstrip()[3:].strip())
                buffer = []
            elif current is not None:
                buffer.append(line)
            # текст до первого "##" без заголовка используем как подзаголовок
            elif line.strip() and deck.subtitle is None:
                deck.subtitle = line.strip()
        if current is not None:
            _fill_slide(current, buffer)
            deck.slides.append(current)
    else:
        blocks = [b for b in re.split(r"\n\s*\n", "\n".join(lines)) if b.strip()]
        for i, block in enumerate(blocks, start=1):
            block_lines = block.split("\n")
            first = block_lines[0]
            if len(block_lines) > 1 and _looks_like_title(first) and not BULLET_RE.match(first):
                slide = SlideContent(title=first.strip())
                _fill_slide(slide, block_lines[1:])
            else:
                slide = SlideContent(title=f"Слайд {i}")
                _fill_slide(slide, block_lines)
            deck.slides.append(slide)

    if not deck.slides:
        raise ValueError("Не найдено ни одного слайда. Добавьте заголовки '## Название слайда'.")
    return deck
