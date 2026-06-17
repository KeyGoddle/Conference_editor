from __future__ import annotations

import copy
import json
import sys
import tkinter as tk
from pathlib import Path
from tkinter import colorchooser, filedialog, messagebox, ttk
from typing import Any

from PIL import Image, ImageDraw, ImageTk

from sertificat import (
    DEFAULT_CONFIG,
    TextBlock,
    draw_centered_text,
    render_documents,
    system_font_catalog,
)


FIELD_LABELS = {
    "full_name_ru": "ФИО RU",
    "full_name_en": "ФИО EN",
    "city_ru": "Город RU",
    "city_en": "Город EN",
    "talk_title": "Название доклада",
}

SAMPLE_TEXT = {
    "full_name_ru": "Иванов Иван",
    "full_name_en": "Ivan Ivanov",
    "city_ru": "Москва",
    "city_en": "Moscow",
    "talk_title": "Как автоматизировать сертификаты",
}

FIELD_OPTIONS = {
    "badges_ru": ["full_name_ru", "city_ru"],
    "badges_en": ["full_name_en", "city_en"],
    "certificates": ["full_name_ru", "talk_title"],
}

BADGE_FIELD_PAIRS = {
    "full_name_ru": "full_name_en",
    "full_name_en": "full_name_ru",
    "city_ru": "city_en",
    "city_en": "city_ru",
}

MIRRORED_BLOCK_KEYS = [
    "x",
    "y",
    "max_width",
    "max_height",
    "font_size",
    "fill",
    "font_path",
    "bold",
    "stroke_width",
    "stroke_fill",
]

def resource_path(relative_path: str) -> Path:
    base_path = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base_path / relative_path


DEFAULT_TEMPLATES = {
    "badges": resource_path("examples/templates/badge_template.png"),
    "certificates": resource_path("examples/templates/certificate_template.png"),
}

DEFAULT_EXCEL_FILES = {
    "badges": resource_path("examples/data/badges.xlsx"),
    "certificates": resource_path("examples/data/certificates.xlsx"),
}

DEFAULT_OUTPUT_DIRS = {
    "badges": Path.cwd() / "generated_badges",
    "certificates": Path.cwd() / "generated_certificates",
}

DEFAULT_LAYOUT_FILE = resource_path("layout.example.json")


class LayoutEditor:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Редактор разметки бейджей и сертификатов")
        self.root.minsize(1120, 760)

        self.config_path = Path.cwd() / "layout.example.json"
        self.config = self.load_initial_config()
        self.template_path = DEFAULT_TEMPLATES["badges"]
        self.image: Image.Image | None = None
        self.preview_photo: ImageTk.PhotoImage | None = None
        self.scale = 1.0
        self.selected_index = 0
        self.dragging = False
        self.drag_offset_x = 0
        self.drag_offset_y = 0

        self.mode_var = tk.StringVar(value="badges")
        self.side_var = tk.StringVar(value="ru")
        self.block_var = tk.StringVar()
        self.field_var = tk.StringVar(value="full_name_ru")
        self.x_var = tk.IntVar(value=500)
        self.y_var = tk.IntVar(value=450)
        self.max_width_var = tk.IntVar(value=760)
        self.max_height_var = tk.IntVar(value=130)
        self.font_size_var = tk.IntVar(value=50)
        self.fill_var = tk.StringVar(value="#111111")
        self.font_path_var = tk.StringVar(value="")
        self.bold_var = tk.BooleanVar(value=False)
        self.side_gap_var = tk.IntVar(value=int(self.config["badges"].get("side_gap", 24)))
        self.font_options = self.load_font_options()
        self.excel_path_var = tk.StringVar(value=str(DEFAULT_EXCEL_FILES["badges"]))
        self.output_dir_var = tk.StringVar(value=str(DEFAULT_OUTPUT_DIRS["badges"]))
        self.sheet_var = tk.StringVar(value="")
        self.output_format_var = tk.StringVar(value="png")
        self.generation_status_var = tk.StringVar(value="")

        self.build_ui()
        self.refresh_side_options()
        self.open_template_path(self.template_path)
        self.select_first_block()
        self.refresh_all()

    def load_initial_config(self) -> dict[str, Any]:
        if self.config_path.exists():
            with self.config_path.open("r", encoding="utf-8") as file:
                return json.load(file)
        if DEFAULT_LAYOUT_FILE.exists():
            with DEFAULT_LAYOUT_FILE.open("r", encoding="utf-8") as file:
                return json.load(file)
        return copy.deepcopy(DEFAULT_CONFIG)

    def load_font_options(self) -> list[str]:
        names = {Path(path).stem for path in system_font_catalog().values()}
        return [""] + sorted(names, key=str.casefold)

    def build_ui(self) -> None:
        self.root.columnconfigure(0, weight=1)
        self.root.columnconfigure(1, weight=0)
        self.root.rowconfigure(0, weight=1)

        canvas_frame = ttk.Frame(self.root, padding=10)
        canvas_frame.grid(row=0, column=0, sticky="nsew")
        canvas_frame.rowconfigure(0, weight=1)
        canvas_frame.columnconfigure(0, weight=1)

        self.canvas = tk.Canvas(canvas_frame, bg="#e5e7eb", highlightthickness=0)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.canvas.bind("<Button-1>", self.on_canvas_down)
        self.canvas.bind("<B1-Motion>", self.on_canvas_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_canvas_up)
        self.canvas.bind("<Configure>", lambda _event: self.refresh_preview())

        panel = self.create_scrollable_panel()

        self.add_file_controls(panel)
        self.add_mode_controls(panel)
        self.add_block_controls(panel)
        self.add_text_controls(panel)
        self.add_generation_controls(panel)
        self.add_action_controls(panel)

    def create_scrollable_panel(self) -> ttk.Frame:
        outer = ttk.Frame(self.root, padding=(0, 10, 10, 10))
        outer.grid(row=0, column=1, sticky="ns")
        outer.rowconfigure(0, weight=1)
        outer.columnconfigure(0, weight=1)

        self.settings_canvas = tk.Canvas(outer, width=340, highlightthickness=0)
        scrollbar = ttk.Scrollbar(outer, orient="vertical", command=self.settings_canvas.yview)
        panel = ttk.Frame(self.settings_canvas)

        panel_window = self.settings_canvas.create_window((0, 0), window=panel, anchor="nw")
        self.settings_canvas.configure(yscrollcommand=scrollbar.set)
        self.settings_canvas.grid(row=0, column=0, sticky="ns")
        scrollbar.grid(row=0, column=1, sticky="ns")

        panel.bind(
            "<Configure>",
            lambda _event: self.settings_canvas.configure(scrollregion=self.settings_canvas.bbox("all")),
        )
        self.settings_canvas.bind(
            "<Configure>",
            lambda event: self.settings_canvas.itemconfigure(panel_window, width=event.width),
        )
        self.settings_canvas.bind("<Enter>", self.bind_settings_scroll)
        self.settings_canvas.bind("<Leave>", self.unbind_settings_scroll)
        panel.bind("<Enter>", self.bind_settings_scroll)
        panel.bind("<Leave>", self.unbind_settings_scroll)

        return panel

    def bind_settings_scroll(self, _event: tk.Event) -> None:
        self.root.bind_all("<MouseWheel>", self.on_settings_mousewheel)
        self.root.bind_all("<Button-4>", self.on_settings_mousewheel)
        self.root.bind_all("<Button-5>", self.on_settings_mousewheel)

    def unbind_settings_scroll(self, _event: tk.Event) -> None:
        self.root.unbind_all("<MouseWheel>")
        self.root.unbind_all("<Button-4>")
        self.root.unbind_all("<Button-5>")

    def on_settings_mousewheel(self, event: tk.Event) -> None:
        if getattr(event, "num", None) == 4:
            units = -3
        elif getattr(event, "num", None) == 5:
            units = 3
        else:
            delta = getattr(event, "delta", 0)
            units = -3 if delta > 0 else 3
        self.settings_canvas.yview_scroll(units, "units")

    def add_file_controls(self, panel: ttk.Frame) -> None:
        box = ttk.LabelFrame(panel, text="Файлы", padding=10)
        box.pack(fill="x", pady=(0, 8))

        ttk.Button(box, text="Открыть шаблон", command=self.choose_template).pack(fill="x")
        ttk.Button(box, text="Открыть layout JSON", command=self.choose_layout).pack(fill="x", pady=4)
        ttk.Button(box, text="Сохранить layout JSON", command=self.save_layout).pack(fill="x")

    def add_mode_controls(self, panel: ttk.Frame) -> None:
        box = ttk.LabelFrame(panel, text="Документ", padding=10)
        box.pack(fill="x", pady=(0, 8))

        ttk.Label(box, text="Тип").grid(row=0, column=0, sticky="w")
        mode = ttk.Combobox(box, textvariable=self.mode_var, values=["badges", "certificates"], state="readonly", width=22)
        mode.grid(row=0, column=1, sticky="ew", padx=(8, 0))
        mode.bind("<<ComboboxSelected>>", self.on_mode_changed)

        ttk.Label(box, text="Сторона").grid(row=1, column=0, sticky="w", pady=(8, 0))
        self.side_combo = ttk.Combobox(box, textvariable=self.side_var, state="readonly", width=22)
        self.side_combo.grid(row=1, column=1, sticky="ew", padx=(8, 0), pady=(8, 0))
        self.side_combo.bind("<<ComboboxSelected>>", self.on_side_changed)

        ttk.Label(box, text="Зазор сторон").grid(row=2, column=0, sticky="w", pady=(8, 0))
        gap = ttk.Spinbox(box, from_=0, to=300, textvariable=self.side_gap_var, width=8, command=self.apply_controls)
        gap.grid(row=2, column=1, sticky="w", padx=(8, 0), pady=(8, 0))
        gap.bind("<KeyRelease>", lambda _event: self.apply_controls())

    def add_block_controls(self, panel: ttk.Frame) -> None:
        box = ttk.LabelFrame(panel, text="Поле", padding=10)
        box.pack(fill="x", pady=(0, 8))

        ttk.Label(box, text="Выбранное").grid(row=0, column=0, sticky="w")
        self.block_combo = ttk.Combobox(box, textvariable=self.block_var, state="readonly", width=22)
        self.block_combo.grid(row=0, column=1, sticky="ew", padx=(8, 0))
        self.block_combo.bind("<<ComboboxSelected>>", self.on_block_changed)

        ttk.Label(box, text="Данные").grid(row=1, column=0, sticky="w", pady=(8, 0))
        self.field_combo = ttk.Combobox(box, textvariable=self.field_var, state="readonly", width=22)
        self.field_combo.grid(row=1, column=1, sticky="ew", padx=(8, 0), pady=(8, 0))
        self.field_combo.bind("<<ComboboxSelected>>", lambda _event: self.apply_controls())

        buttons = ttk.Frame(box)
        buttons.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        ttk.Button(buttons, text="Добавить", command=self.add_block).pack(side="left", expand=True, fill="x")
        ttk.Button(buttons, text="Удалить", command=self.delete_block).pack(side="left", expand=True, fill="x", padx=(8, 0))

    def add_text_controls(self, panel: ttk.Frame) -> None:
        box = ttk.LabelFrame(panel, text="Бокс / рамка текста", padding=10)
        box.pack(fill="x", pady=(0, 8))

        rows = [
            ("X", self.x_var, 0, 5000),
            ("Y", self.y_var, 0, 5000),
            ("Ширина", self.max_width_var, 10, 5000),
            ("Высота", self.max_height_var, 10, 5000),
            ("Размер", self.font_size_var, 8, 300),
        ]
        for row, (label, variable, start, end) in enumerate(rows):
            ttk.Label(box, text=label).grid(row=row, column=0, sticky="w", pady=(0, 6))
            spin = ttk.Spinbox(box, from_=start, to=end, textvariable=variable, width=10, command=self.apply_controls)
            spin.grid(row=row, column=1, sticky="ew", padx=(8, 0), pady=(0, 6))
            spin.bind("<KeyRelease>", lambda _event: self.apply_controls())

        center_row = ttk.Frame(box)
        center_row.grid(row=5, column=0, columnspan=2, sticky="ew", pady=(2, 8))
        ttk.Button(center_row, text="Центр X", command=lambda: self.center_selected("x")).pack(
            side="left", expand=True, fill="x"
        )
        ttk.Button(center_row, text="Центр Y", command=lambda: self.center_selected("y")).pack(
            side="left", expand=True, fill="x", padx=(6, 0)
        )
        ttk.Button(center_row, text="Центр XY", command=lambda: self.center_selected("xy")).pack(
            side="left", expand=True, fill="x", padx=(6, 0)
        )

        ttk.Button(box, text="Все блоки по центру X", command=self.center_all_blocks_x).grid(
            row=6, column=0, columnspan=2, sticky="ew", pady=(0, 10)
        )

        ttk.Label(box, text="Цвет").grid(row=7, column=0, sticky="w", pady=(0, 6))
        color_row = ttk.Frame(box)
        color_row.grid(row=7, column=1, sticky="ew", padx=(8, 0), pady=(0, 6))
        ttk.Entry(color_row, textvariable=self.fill_var, width=10).pack(side="left", fill="x", expand=True)
        ttk.Button(color_row, text="...", width=3, command=self.choose_color).pack(side="left", padx=(4, 0))

        ttk.Label(box, text="Шрифт").grid(row=8, column=0, sticky="w", pady=(0, 6))
        font_row = ttk.Frame(box)
        font_row.grid(row=8, column=1, sticky="ew", padx=(8, 0), pady=(0, 6))
        font_combo = ttk.Combobox(font_row, textvariable=self.font_path_var, values=self.font_options, width=18)
        font_combo.pack(side="left", fill="x", expand=True)
        font_combo.bind("<<ComboboxSelected>>", lambda _event: self.apply_controls())
        font_combo.bind("<KeyRelease>", lambda _event: self.apply_controls())
        ttk.Button(font_row, text="...", width=3, command=self.choose_font).pack(side="left", padx=(4, 0))

        ttk.Checkbutton(box, text="Жирный", variable=self.bold_var, command=self.apply_controls).grid(
            row=9, column=0, columnspan=2, sticky="w", pady=(2, 6)
        )

        ttk.Button(box, text="Применить", command=self.apply_controls).grid(
            row=10, column=0, columnspan=2, sticky="ew", pady=(10, 0)
        )

    def add_generation_controls(self, panel: ttk.Frame) -> None:
        box = ttk.LabelFrame(panel, text="Генерация", padding=10)
        box.pack(fill="x", pady=(0, 8))

        ttk.Label(box, text="Excel").grid(row=0, column=0, sticky="w", pady=(0, 6))
        excel_row = ttk.Frame(box)
        excel_row.grid(row=0, column=1, sticky="ew", padx=(8, 0), pady=(0, 6))
        ttk.Entry(excel_row, textvariable=self.excel_path_var, width=18).pack(side="left", fill="x", expand=True)
        ttk.Button(excel_row, text="...", width=3, command=self.choose_excel).pack(side="left", padx=(4, 0))

        ttk.Label(box, text="Папка").grid(row=1, column=0, sticky="w", pady=(0, 6))
        output_row = ttk.Frame(box)
        output_row.grid(row=1, column=1, sticky="ew", padx=(8, 0), pady=(0, 6))
        ttk.Entry(output_row, textvariable=self.output_dir_var, width=18).pack(side="left", fill="x", expand=True)
        ttk.Button(output_row, text="...", width=3, command=self.choose_output_dir).pack(side="left", padx=(4, 0))

        ttk.Label(box, text="Лист").grid(row=2, column=0, sticky="w", pady=(0, 6))
        ttk.Entry(box, textvariable=self.sheet_var, width=22).grid(row=2, column=1, sticky="ew", padx=(8, 0), pady=(0, 6))

        ttk.Label(box, text="Формат").grid(row=3, column=0, sticky="w", pady=(0, 6))
        fmt = ttk.Combobox(
            box,
            textvariable=self.output_format_var,
            values=["png", "jpg", "jpeg"],
            state="readonly",
            width=10,
        )
        fmt.grid(row=3, column=1, sticky="w", padx=(8, 0), pady=(0, 6))

        ttk.Button(box, text="Запустить генерацию", command=self.run_generation).grid(
            row=4, column=0, columnspan=2, sticky="ew", pady=(8, 0)
        )
        ttk.Label(box, textvariable=self.generation_status_var, wraplength=300).grid(
            row=5, column=0, columnspan=2, sticky="w", pady=(8, 0)
        )

    def add_action_controls(self, panel: ttk.Frame) -> None:
        box = ttk.LabelFrame(panel, text="Подсказка", padding=10)
        box.pack(fill="x")
        text = (
            "Клик по рамке выбирает поле.\n"
            "Перетаскивание двигает поле.\n"
            "Ширина и высота задают область,\n"
            "куда текст должен помещаться."
        )
        ttk.Label(box, text=text, justify="left").pack(anchor="w")

    def current_mode(self) -> str:
        return self.mode_var.get()

    def badge_sides(self) -> list[dict[str, Any]]:
        badges = self.config.setdefault("badges", {})
        sides = badges.get("sides")
        if isinstance(sides, list) and sides:
            return sides
        badges["sides"] = copy.deepcopy(DEFAULT_CONFIG["badges"]["sides"])
        return badges["sides"]

    def current_blocks(self) -> list[dict[str, Any]]:
        mode = self.current_mode()
        if mode == "badges":
            side_name = self.side_var.get()
            sides = self.badge_sides()
            for side in sides:
                if side.get("name") == side_name:
                    return side.setdefault("blocks", [])
            return sides[0].setdefault("blocks", [])

        certs = self.config.setdefault("certificates", {})
        return certs.setdefault("blocks", copy.deepcopy(DEFAULT_CONFIG["certificates"]["blocks"]))

    def available_fields(self) -> list[str]:
        if self.current_mode() == "certificates":
            return FIELD_OPTIONS["certificates"]
        if self.side_var.get() == "en":
            return FIELD_OPTIONS["badges_en"]
        return FIELD_OPTIONS["badges_ru"]

    def current_block(self) -> dict[str, Any] | None:
        blocks = self.current_blocks()
        if not blocks:
            return None
        self.selected_index = max(0, min(self.selected_index, len(blocks) - 1))
        return blocks[self.selected_index]

    def refresh_side_options(self) -> None:
        if self.current_mode() != "badges":
            self.side_combo.configure(values=[], state="disabled")
            return
        names = [str(side.get("name", f"side_{index}")) for index, side in enumerate(self.badge_sides(), start=1)]
        self.side_combo.configure(values=names, state="readonly")
        if self.side_var.get() not in names:
            self.side_var.set(names[0])

    def refresh_block_options(self) -> None:
        blocks = self.current_blocks()
        values = [
            f"{index + 1}. {FIELD_LABELS.get(str(block.get('field')), str(block.get('field')))}"
            for index, block in enumerate(blocks)
        ]
        self.block_combo.configure(values=values)
        if values:
            self.selected_index = max(0, min(self.selected_index, len(values) - 1))
            self.block_var.set(values[self.selected_index])
        else:
            self.block_var.set("")

        fields = self.available_fields()
        self.field_combo.configure(values=fields)

    def select_first_block(self) -> None:
        self.selected_index = 0
        self.refresh_block_options()
        self.load_controls_from_block()

    def refresh_all(self) -> None:
        self.refresh_side_options()
        self.refresh_block_options()
        self.load_controls_from_block()
        self.refresh_preview()

    def load_controls_from_block(self) -> None:
        block = self.current_block()
        if not block:
            return
        self.field_var.set(str(block.get("field", self.available_fields()[0])))
        self.x_var.set(int(block.get("x", 500)))
        self.y_var.set(int(block.get("y", 450)))
        self.max_width_var.set(int(block.get("max_width", 700)))
        self.max_height_var.set(int(block.get("max_height", 120)))
        self.font_size_var.set(int(block.get("font_size", 40)))
        self.fill_var.set(str(block.get("fill", "#111111")))
        self.font_path_var.set(str(block.get("font_path") or ""))
        self.bold_var.set(bool(block.get("bold", False)))
        self.side_gap_var.set(int(self.config.get("badges", {}).get("side_gap", 24)))

    def apply_controls(self) -> None:
        block = self.current_block()
        if not block:
            return

        try:
            x = int(self.x_var.get())
            y = int(self.y_var.get())
            max_width = int(self.max_width_var.get())
            max_height = int(self.max_height_var.get())
            font_size = int(self.font_size_var.get())
            side_gap = int(self.side_gap_var.get())
        except (tk.TclError, ValueError):
            return

        block["field"] = self.field_var.get()
        block["x"] = x
        block["y"] = y
        block["max_width"] = max_width
        block["max_height"] = max_height
        block["font_size"] = font_size
        block["fill"] = self.fill_var.get().strip() or "#111111"
        block["bold"] = bool(self.bold_var.get())
        block.pop("align", None)

        font_path = self.font_path_var.get().strip()
        if font_path:
            block["font_path"] = font_path
        else:
            block.pop("font_path", None)

        self.config.setdefault("badges", {})["side_gap"] = side_gap
        self.mirror_badge_block(block)
        self.refresh_block_options()
        self.refresh_preview()

    def mirror_badge_block(self, block: dict[str, Any]) -> None:
        if self.current_mode() != "badges":
            return

        source_field = str(block.get("field", ""))
        target_field = BADGE_FIELD_PAIRS.get(source_field)
        if not target_field:
            return

        current_side = self.side_var.get()
        target_block: dict[str, Any] | None = None
        target_side: dict[str, Any] | None = None
        for side in self.badge_sides():
            if side.get("name") == current_side:
                continue
            for candidate in side.setdefault("blocks", []):
                if candidate.get("field") == target_field:
                    target_block = candidate
                    target_side = side
                    break
            if target_block is not None:
                break

        if target_block is None:
            for side in self.badge_sides():
                if side.get("name") != current_side:
                    target_side = side
                    break
            if target_side is None:
                return
            target_block = {"field": target_field}
            target_side.setdefault("blocks", []).append(target_block)

        target_block["field"] = target_field
        for key in MIRRORED_BLOCK_KEYS:
            if key in block:
                target_block[key] = copy.deepcopy(block[key])
            else:
                target_block.pop(key, None)

    def center_selected(self, axis: str) -> None:
        if not self.image:
            return
        if "x" in axis:
            self.x_var.set(self.image.width // 2)
        if "y" in axis:
            self.y_var.set(self.image.height // 2)
        self.apply_controls()

    def center_all_blocks_x(self) -> None:
        if not self.image:
            return
        for block in self.current_blocks():
            block["x"] = self.image.width // 2
            self.mirror_badge_block(block)
        self.load_controls_from_block()
        self.refresh_preview()

    def choose_template(self) -> None:
        path = filedialog.askopenfilename(
            title="Выбери PNG/JPG шаблон",
            filetypes=[("Images", "*.png *.jpg *.jpeg"), ("All files", "*.*")],
        )
        if path:
            self.open_template_path(Path(path))

    def open_template_path(self, path: Path) -> None:
        if not path.exists():
            return
        self.template_path = path
        self.image = Image.open(path).convert("RGBA")
        self.refresh_preview()

    def choose_layout(self) -> None:
        path = filedialog.askopenfilename(
            title="Открыть layout JSON",
            filetypes=[("JSON", "*.json"), ("All files", "*.*")],
        )
        if not path:
            return
        self.config_path = Path(path)
        with self.config_path.open("r", encoding="utf-8") as file:
            self.config = json.load(file)
        self.selected_index = 0
        self.refresh_all()

    def save_layout(self) -> None:
        self.apply_controls()
        path = filedialog.asksaveasfilename(
            title="Сохранить layout JSON",
            initialfile=self.config_path.name,
            defaultextension=".json",
            filetypes=[("JSON", "*.json"), ("All files", "*.*")],
        )
        if not path:
            return
        self.config_path = Path(path)
        with self.config_path.open("w", encoding="utf-8") as file:
            json.dump(self.config, file, ensure_ascii=False, indent=2)
            file.write("\n")
        messagebox.showinfo("Готово", f"Настройки сохранены:\n{self.config_path}")

    def save_layout_to_current_path(self) -> None:
        self.apply_controls()
        with self.config_path.open("w", encoding="utf-8") as file:
            json.dump(self.config, file, ensure_ascii=False, indent=2)
            file.write("\n")

    def choose_excel(self) -> None:
        path = filedialog.askopenfilename(
            title="Выбери Excel файл",
            filetypes=[("Excel", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            self.excel_path_var.set(path)

    def choose_output_dir(self) -> None:
        path = filedialog.askdirectory(title="Выбери папку для готовых файлов")
        if path:
            self.output_dir_var.set(path)

    def run_generation(self) -> None:
        mode = self.current_mode()
        excel_path = Path(self.excel_path_var.get()).expanduser()
        output_dir = Path(self.output_dir_var.get()).expanduser()
        sheet_name = self.sheet_var.get().strip() or None
        forced_format = self.output_format_var.get().strip() or None

        if not excel_path.exists():
            messagebox.showerror("Ошибка", f"Excel файл не найден:\n{excel_path}")
            return
        if not self.template_path.exists():
            messagebox.showerror("Ошибка", f"Шаблон не найден:\n{self.template_path}")
            return

        try:
            self.generation_status_var.set("Генерация...")
            self.root.update_idletasks()
            self.save_layout_to_current_path()
            count = render_documents(
                mode=mode,
                excel_path=excel_path,
                template_path=self.template_path,
                output_dir=output_dir,
                config_path=self.config_path,
                sheet_name=sheet_name,
                forced_format=forced_format,
            )
        except Exception as exc:
            self.generation_status_var.set("Ошибка генерации")
            messagebox.showerror("Ошибка генерации", str(exc))
            return

        message = f"Готово: создано файлов — {count}. Папка: {output_dir}"
        self.generation_status_var.set(message)
        messagebox.showinfo("Готово", message)

    def choose_color(self) -> None:
        color = colorchooser.askcolor(color=self.fill_var.get())[1]
        if color:
            self.fill_var.set(color)
            self.apply_controls()

    def choose_font(self) -> None:
        path = filedialog.askopenfilename(
            title="Выбери файл шрифта",
            filetypes=[("Fonts", "*.ttf *.otf *.ttc"), ("All files", "*.*")],
        )
        if path:
            self.font_path_var.set(path)
            self.apply_controls()

    def on_mode_changed(self, _event: tk.Event) -> None:
        mode = self.current_mode()
        default_template = DEFAULT_TEMPLATES[mode]
        if default_template.exists():
            self.open_template_path(default_template)
        self.excel_path_var.set(str(DEFAULT_EXCEL_FILES[mode]))
        self.output_dir_var.set(str(DEFAULT_OUTPUT_DIRS[mode]))
        self.selected_index = 0
        self.refresh_all()

    def on_side_changed(self, _event: tk.Event) -> None:
        self.selected_index = 0
        self.refresh_all()

    def on_block_changed(self, _event: tk.Event) -> None:
        value = self.block_var.get()
        if value:
            self.selected_index = max(0, int(value.split(".", 1)[0]) - 1)
        self.load_controls_from_block()
        self.refresh_preview()

    def add_block(self) -> None:
        fields = self.available_fields()
        blocks = self.current_blocks()
        field = fields[min(len(blocks), len(fields) - 1)]
        block = {
            "field": field,
            "x": self.image.width // 2 if self.image else 500,
            "y": self.image.height // 2 if self.image else 450,
            "max_width": 700,
            "max_height": 120,
            "font_size": 40,
            "fill": "#111111",
        }
        blocks.append(block)
        self.mirror_badge_block(block)
        self.selected_index = len(blocks) - 1
        self.refresh_all()

    def delete_block(self) -> None:
        blocks = self.current_blocks()
        if not blocks:
            return
        del blocks[self.selected_index]
        self.selected_index = max(0, self.selected_index - 1)
        self.refresh_all()

    def canvas_to_image(self, event: tk.Event) -> tuple[int, int]:
        canvas_width = max(1, self.canvas.winfo_width())
        canvas_height = max(1, self.canvas.winfo_height())
        if not self.image:
            return 0, 0

        display_width = int(self.image.width * self.scale)
        display_height = int(self.image.height * self.scale)
        offset_x = (canvas_width - display_width) // 2
        offset_y = (canvas_height - display_height) // 2
        x = int((event.x - offset_x) / self.scale)
        y = int((event.y - offset_y) / self.scale)
        return x, y

    def block_hit_index(self, x: int, y: int) -> int | None:
        for index, block in enumerate(self.current_blocks()):
            left = int(block.get("x", 0)) - int(block.get("max_width", 0)) // 2
            right = int(block.get("x", 0)) + int(block.get("max_width", 0)) // 2
            top = int(block.get("y", 0)) - int(block.get("max_height", 0)) // 2
            bottom = int(block.get("y", 0)) + int(block.get("max_height", 0)) // 2
            if left <= x <= right and top <= y <= bottom:
                return index
        return None

    def on_canvas_down(self, event: tk.Event) -> None:
        x, y = self.canvas_to_image(event)
        hit = self.block_hit_index(x, y)
        if hit is not None:
            self.selected_index = hit
            self.refresh_block_options()
            self.load_controls_from_block()
            block = self.current_block()
            self.drag_offset_x = x - int(block.get("x", 0)) if block else 0
            self.drag_offset_y = y - int(block.get("y", 0)) if block else 0
        else:
            self.drag_offset_x = 0
            self.drag_offset_y = 0
            self.x_var.set(x)
            self.y_var.set(y)
            self.apply_controls()

        self.dragging = True

    def on_canvas_drag(self, event: tk.Event) -> None:
        if not self.dragging:
            return
        x, y = self.canvas_to_image(event)
        self.x_var.set(x - self.drag_offset_x)
        self.y_var.set(y - self.drag_offset_y)
        self.apply_controls()

    def on_canvas_up(self, _event: tk.Event) -> None:
        self.dragging = False

    def make_preview_image(self) -> Image.Image | None:
        if not self.image:
            return None

        preview = self.image.copy()
        draw = ImageDraw.Draw(preview)
        blocks = self.current_blocks()

        for index, block in enumerate(blocks):
            left = int(block.get("x", 0)) - int(block.get("max_width", 0)) // 2
            right = int(block.get("x", 0)) + int(block.get("max_width", 0)) // 2
            top = int(block.get("y", 0)) - int(block.get("max_height", 0)) // 2
            bottom = int(block.get("y", 0)) + int(block.get("max_height", 0)) // 2
            color = "#2563eb" if index == self.selected_index else "#94a3b8"
            draw.rounded_rectangle((left, top, right, bottom), radius=10, outline=color, width=4)

            try:
                text_block = TextBlock(
                    field=str(block.get("field", "")),
                    x=int(block.get("x", 0)),
                    y=int(block.get("y", 0)),
                    max_width=int(block.get("max_width", 100)),
                    max_height=int(block.get("max_height", 100)),
                    font_size=int(block.get("font_size", 40)),
                    fill=str(block.get("fill", "#111111")),
                    font_path=block.get("font_path"),
                    bold=bool(block.get("bold", False)),
                    stroke_width=int(block.get("stroke_width", 0)),
                    stroke_fill=block.get("stroke_fill"),
                )
                draw_centered_text(preview, text_block, SAMPLE_TEXT.get(text_block.field, text_block.field))
            except Exception:
                draw.text((left + 10, top + 10), str(block.get("field", "")), fill="#ef4444")

        return preview

    def refresh_preview(self) -> None:
        preview = self.make_preview_image()
        if preview is None:
            return

        canvas_width = max(1, self.canvas.winfo_width() - 20)
        canvas_height = max(1, self.canvas.winfo_height() - 20)
        self.scale = min(canvas_width / preview.width, canvas_height / preview.height, 1.0)
        display_size = (max(1, int(preview.width * self.scale)), max(1, int(preview.height * self.scale)))
        resized = preview.resize(display_size, Image.Resampling.LANCZOS)
        self.preview_photo = ImageTk.PhotoImage(resized)

        self.canvas.delete("all")
        x = (self.canvas.winfo_width() - display_size[0]) // 2
        y = (self.canvas.winfo_height() - display_size[1]) // 2
        self.canvas.create_image(x, y, anchor="nw", image=self.preview_photo)


def main() -> None:
    root = tk.Tk()
    LayoutEditor(root)
    root.mainloop()


if __name__ == "__main__":
    main()
