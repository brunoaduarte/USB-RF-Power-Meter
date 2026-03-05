from __future__ import annotations

import platform
import queue
import subprocess
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from dataclasses import dataclass
from math import ceil
from typing import Callable

from PIL import Image, ImageDraw, ImageFont
from serial.tools import list_ports

from usb_rf_power_meter.protocol import Measurement, SyncEntry
from usb_rf_power_meter.serial_worker import SerialWorker

Y_MIN_DBM = -70.0
Y_MAX_DBM = 40.0
POINT_SPACING = 24
CHART_HEIGHT = 520
CHART_WIDTH = 900
POLL_INTERVAL_MS = 50
MAX_DRAWN_SAMPLES_PER_BATCH = 200
CONNECT_SYNC_DELAY_MS = 1000
APPEARANCE_POLL_INTERVAL_MS = 1500
RATE_OPTIONS = {
    "S0 - 1s (Slow)": "S0",
    "S1 - 200 ms (Fast)": "S1",
    "S2 - 500 ns (Very fast)": "S2",
}
SYNC_COMMAND = "Read"
SYNC_PROFILE_COMMAND_PREFIXES = "ABCDEFGHI"
IGNORED_PORTS = {
    "/dev/cu.Bluetooth-Incoming-Port",
    "/dev/cu.debug-console",
    "/dev/cu.wlan-debug",
}
CHART_FILE_EXTENSION = ".csv"
EXPORT_FILE_EXTENSION = ".jpg"


@dataclass(frozen=True)
class AppPalette:
    window_bg: str
    surface_bg: str
    card_bg: str
    chart_bg: str
    log_bg: str
    text: str
    muted_text: str
    danger: str
    grid_major: str
    grid_minor: str
    chart_border: str
    entry_bg: str
    entry_fg: str
    selection_bg: str
    selection_fg: str


def is_dark_mode() -> bool:
    if platform.system() != "Darwin":
        return False

    result = subprocess.run(
        ["defaults", "read", "-g", "AppleInterfaceStyle"],
        capture_output=True,
        text=True,
        check=False,
    )
    return result.returncode == 0 and result.stdout.strip().lower() == "dark"


def build_palette(prefers_dark: bool) -> AppPalette:
    if prefers_dark:
        return AppPalette(
            window_bg="#2f2f31",
            surface_bg="#343436",
            card_bg="#3b3b3e",
            chart_bg="#202124",
            log_bg="#242528",
            text="#f5f5f7",
            muted_text="#d1d1d6",
            danger="#ff3b30",
            grid_major="#6a6a6e",
            grid_minor="#45464a",
            chart_border="#8a8a8f",
            entry_bg="#4a4a4d",
            entry_fg="#f5f5f7",
            selection_bg="#0a84ff",
            selection_fg="#ffffff",
        )

    return AppPalette(
        window_bg="#ececec",
        surface_bg="#ececec",
        card_bg="#f4f4f4",
        chart_bg="#ffffff",
        log_bg="#ffffff",
        text="#111111",
        muted_text="#4f4f4f",
        danger="#ff3b30",
        grid_major="#808080",
        grid_minor="#c2c2c2",
        chart_border="#606060",
        entry_bg="#ffffff",
        entry_fg="#111111",
        selection_bg="#0a84ff",
        selection_fg="#ffffff",
    )


def format_dbm(value: float) -> str:
    return f"{value:.1f}dBm"


def format_microwatts(value: float) -> str:
    return f"{value:.2f}".replace(".", ",") + "uW"


def dbm_to_microwatts(dbm_value: float) -> float:
    return 10 ** (dbm_value / 10.0) * 1000.0


class SignalChart(ttk.Frame):
    def __init__(self, master: tk.Misc, palette: AppPalette) -> None:
        super().__init__(master)
        self._palette = palette
        self._samples: list[float] = []
        self._hover_y: float | None = None
        self._hover_x: float | None = None
        self._axis_width = 54
        self._top_padding = 18
        self._bottom_padding = 30
        self._left_padding = 12
        self._base_point_spacing = float(POINT_SPACING)
        self._point_spacing = float(POINT_SPACING)
        self._zoom_mode = "base"

        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)

        self._axis_canvas = tk.Canvas(
            self,
            width=self._axis_width,
            height=CHART_HEIGHT,
            bg=self._palette.chart_bg,
            highlightthickness=0,
        )
        self._axis_canvas.grid(row=0, column=0, sticky="ns")

        self._plot_canvas = tk.Canvas(
            self,
            width=CHART_WIDTH,
            height=CHART_HEIGHT,
            bg=self._palette.chart_bg,
            highlightthickness=0,
            xscrollincrement=POINT_SPACING,
        )
        self._plot_canvas.grid(row=0, column=1, sticky="nsew")
        self._plot_canvas.bind("<Configure>", self._on_plot_resize)
        self._plot_canvas.bind("<Motion>", self._on_plot_hover)
        self._plot_canvas.bind("<Leave>", self._hide_hover_value)

        self._scrollbar = ttk.Scrollbar(self, orient="horizontal", command=self._plot_canvas.xview)
        self._scrollbar.grid(row=1, column=1, sticky="ew")
        self._plot_canvas.configure(xscrollcommand=self._scrollbar.set)

        self._draw_axis()
        self._redraw()

    def apply_palette(self, palette: AppPalette) -> None:
        self._palette = palette
        self._axis_canvas.configure(bg=self._palette.chart_bg)
        self._plot_canvas.configure(bg=self._palette.chart_bg)
        self._draw_axis()
        self._redraw()

    def append(self, dbm_value: float) -> None:
        self._samples.append(dbm_value)
        self._redraw()

    def append_many(self, dbm_values: list[float]) -> None:
        if not dbm_values:
            return
        self._samples.extend(dbm_values)
        if self._zoom_mode == "fit":
            self._point_spacing = self._fit_point_spacing()
        self._redraw()

    def clear(self) -> None:
        self._samples.clear()
        self._point_spacing = self._base_point_spacing
        self._zoom_mode = "base"
        self._redraw()

    def set_samples(self, dbm_values: list[float]) -> None:
        self._samples = list(dbm_values)
        self._point_spacing = self._base_point_spacing
        self._zoom_mode = "base"
        self._redraw()

    def samples(self) -> list[float]:
        return list(self._samples)

    def export_to_jpeg(self, destination: Path) -> None:
        export_width = CHART_WIDTH
        export_spacing = self._export_fit_spacing(export_width)
        self._render_chart_image(export_width, export_spacing).save(destination, format="JPEG", quality=95)

    def has_samples(self) -> bool:
        return bool(self._samples)

    def zoom_out(self) -> None:
        if not self._samples:
            return

        fit_spacing = self._fit_point_spacing()
        if self._point_spacing <= fit_spacing:
            self._point_spacing = fit_spacing
            self._zoom_mode = "fit"
            self._redraw()
            return

        self._point_spacing = max(fit_spacing, self._point_spacing / 2.0)
        self._zoom_mode = "fit" if self._point_spacing <= fit_spacing else "custom"
        self._redraw()

    def zoom_to_fit(self) -> None:
        if not self._samples:
            return
        self._point_spacing = self._fit_point_spacing()
        self._zoom_mode = "fit"
        self._redraw()

    def reset_zoom(self) -> None:
        self._point_spacing = self._base_point_spacing
        self._zoom_mode = "base"
        self._redraw()

    def can_zoom_out(self) -> bool:
        return self.has_samples() and self._point_spacing > self._fit_point_spacing()

    def can_zoom_to_fit(self) -> bool:
        return self.has_samples() and self._zoom_mode != "fit"

    def can_reset_zoom(self) -> bool:
        return self.has_samples() and self._zoom_mode != "base"

    def _render_chart_image(self, plot_width: int, point_spacing: float) -> Image.Image:
        total_width = self._axis_width + plot_width
        image = Image.new("RGB", (total_width, CHART_HEIGHT), self._palette.chart_bg)
        draw = ImageDraw.Draw(image)
        label_font = self._load_font(14, bold=True)
        tick_font = self._load_font(12)
        index_font = self._load_font(10)

        draw.rectangle((0, 0, total_width, CHART_HEIGHT), fill=self._palette.chart_bg)
        zero_y = self._map_y(0.0)
        draw.text((4, zero_y - 7), "dBm", fill=self._palette.text, font=label_font)
        for tick in range(int(Y_MAX_DBM), int(Y_MIN_DBM) - 1, -5):
            y = self._map_y(float(tick))
            tick_text = str(tick)
            text_bbox = draw.textbbox((0, 0), tick_text, font=tick_font)
            text_width = text_bbox[2] - text_bbox[0]
            text_height = text_bbox[3] - text_bbox[1]
            draw.text((self._axis_width - 6 - text_width, y - text_height / 2), tick_text, fill=self._palette.text, font=tick_font)
            color = self._palette.grid_major if tick % 10 == 0 else self._palette.grid_minor
            draw.line((self._axis_width, y, self._axis_width + plot_width, y), fill=color, width=1)

        plot_bottom = CHART_HEIGHT - self._bottom_padding
        total_columns = len(self._samples)
        visible_columns = max(1, int((plot_width - self._left_padding * 2) / max(1.0, point_spacing)))
        label_step = max(1, int(ceil(28.0 / max(1.0, point_spacing))))
        for index in range(visible_columns + 1):
            x = self._axis_width + int(round(self._left_padding + index * point_spacing))
            color = self._palette.grid_major if index % 5 == 0 else self._palette.grid_minor
            draw.line((x, self._top_padding, x, plot_bottom), fill=color, width=1)
            if total_columns > 0 and index < total_columns and index % label_step == 0:
                index_text = str(index)
                text_bbox = draw.textbbox((0, 0), index_text, font=index_font)
                text_width = text_bbox[2] - text_bbox[0]
                draw.text((x - text_width / 2, plot_bottom + 6), index_text, fill=self._palette.text, font=index_font)

        draw.rectangle(
            (
                self._axis_width + self._left_padding,
                self._top_padding,
                self._axis_width + plot_width - self._left_padding,
                plot_bottom,
            ),
            outline=self._palette.chart_border,
            width=1,
        )

        if len(self._samples) > 1:
            points: list[float] = []
            for index, dbm_value in enumerate(self._samples):
                x = self._axis_width + self._left_padding + index * point_spacing
                points.extend((x, self._map_y(dbm_value)))
            draw.line(points, fill=self._palette.danger, width=2)
        elif len(self._samples) == 1:
            x = self._axis_width + self._left_padding
            y = self._map_y(self._samples[0])
            draw.ellipse((x - 2, y - 2, x + 2, y + 2), fill=self._palette.danger, outline=self._palette.danger)

        return image

    def _export_fit_spacing(self, width: int) -> float:
        if len(self._samples) <= 1:
            return self._base_point_spacing

        usable_width = max(1.0, float(width - self._left_padding * 2))
        spacing = usable_width / float(len(self._samples) - 1)
        return min(self._base_point_spacing, spacing)

    def _load_font(self, size: int, *, bold: bool = False) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
        candidates = (
            [
                "/System/Library/Fonts/Supplemental/Arial Bold.ttf",
                "/System/Library/Fonts/Supplemental/Arial.ttf",
                "/System/Library/Fonts/Supplemental/Helvetica.ttc",
            ]
            if bold
            else [
                "/System/Library/Fonts/Supplemental/Arial.ttf",
                "/System/Library/Fonts/Supplemental/Helvetica.ttc",
            ]
        )
        for candidate in candidates:
            try:
                return ImageFont.truetype(candidate, size)
            except OSError:
                continue
        return ImageFont.load_default()

    def _draw_axis(self) -> None:
        self._axis_canvas.delete("all")
        zero_y = self._map_y(0.0)
        self._axis_canvas.create_text(
            4,
            zero_y,
            text="dBm",
            anchor="w",
            font=("Segoe UI", 10, "bold"),
            fill=self._palette.text,
        )

        for tick in range(int(Y_MAX_DBM), int(Y_MIN_DBM) - 1, -5):
            y = self._map_y(float(tick))
            self._axis_canvas.create_text(
                self._axis_width - 6,
                y,
                text=str(tick),
                anchor="e",
                font=("Segoe UI", 9),
                fill=self._palette.text,
            )

    def _redraw(self) -> None:
        self._plot_canvas.delete("all")
        self._plot_canvas.configure(xscrollincrement=max(1, int(round(self._point_spacing))))
        viewport_width = self._viewport_width()
        width = max(viewport_width, self._total_plot_width())
        self._plot_canvas.configure(scrollregion=(0, 0, width, CHART_HEIGHT))
        self._draw_grid(width)

        if len(self._samples) > 1:
            points: list[float] = []
            for index, dbm_value in enumerate(self._samples):
                points.extend((self._x_for_index(index), self._map_y(dbm_value)))
            self._plot_canvas.create_line(*points, fill=self._palette.danger, width=2, smooth=False)

        elif len(self._samples) == 1:
            x = self._x_for_index(0)
            y = self._map_y(self._samples[0])
            self._plot_canvas.create_oval(
                x - 2,
                y - 2,
                x + 2,
                y + 2,
                fill=self._palette.danger,
                outline=self._palette.danger,
            )

        if width > viewport_width:
            self._plot_canvas.xview_moveto(1.0)
        else:
            self._plot_canvas.xview_moveto(0.0)

        self._redraw_hover_value()

    def _draw_grid(self, width: int) -> None:
        plot_bottom = CHART_HEIGHT - self._bottom_padding

        for tick in range(int(Y_MAX_DBM), int(Y_MIN_DBM) - 1, -5):
            y = self._map_y(float(tick))
            color = self._palette.grid_major if tick % 10 == 0 else self._palette.grid_minor
            self._plot_canvas.create_line(0, y, width, y, fill=color)

        total_columns = len(self._samples)
        spacing = max(1.0, self._point_spacing)
        visible_columns = max(1, int((width - self._left_padding * 2) / spacing))
        label_step = max(1, int(ceil(28.0 / spacing)))
        for index in range(visible_columns + 1):
            x = self._x_for_index(index)
            color = self._palette.grid_major if index % 5 == 0 else self._palette.grid_minor
            self._plot_canvas.create_line(x, self._top_padding, x, plot_bottom, fill=color)
            if total_columns > 0 and index < total_columns and index % label_step == 0:
                self._plot_canvas.create_text(
                    x,
                    plot_bottom + 14,
                    text=str(index),
                    anchor="n",
                    font=("Segoe UI", 8),
                    fill=self._palette.text,
                )

        self._plot_canvas.create_rectangle(
            self._left_padding,
            self._top_padding,
            width - self._left_padding,
            plot_bottom,
            outline=self._palette.chart_border,
        )

    def _map_y(self, value: float) -> float:
        bounded = min(max(value, Y_MIN_DBM), Y_MAX_DBM)
        ratio = (Y_MAX_DBM - bounded) / (Y_MAX_DBM - Y_MIN_DBM)
        usable_height = CHART_HEIGHT - self._top_padding - self._bottom_padding
        return self._top_padding + ratio * usable_height

    def _map_value_from_y(self, y: float) -> float:
        usable_height = CHART_HEIGHT - self._top_padding - self._bottom_padding
        bounded_y = min(max(y, self._top_padding), CHART_HEIGHT - self._bottom_padding)
        ratio = (bounded_y - self._top_padding) / usable_height
        value = Y_MAX_DBM - ratio * (Y_MAX_DBM - Y_MIN_DBM)
        return min(max(value, Y_MIN_DBM), Y_MAX_DBM)

    def _x_for_index(self, index: int) -> int:
        return int(round(self._left_padding + index * self._point_spacing))

    def _total_plot_width(self) -> int:
        if not self._samples:
            return 0
        return int(round(self._left_padding * 2 + max(1, len(self._samples) - 1) * self._point_spacing))

    def _fit_point_spacing(self) -> float:
        if len(self._samples) <= 1:
            return self._base_point_spacing

        usable_width = max(1.0, float(self._viewport_width() - self._left_padding * 2))
        spacing = usable_width / float(len(self._samples) - 1)
        return min(self._base_point_spacing, spacing)

    def _viewport_width(self) -> int:
        current_width = self._plot_canvas.winfo_width()
        return current_width if current_width > 1 else CHART_WIDTH

    def _on_plot_resize(self, _event: tk.Event[tk.Misc]) -> None:
        if self._zoom_mode == "fit":
            self._point_spacing = self._fit_point_spacing()
        self._redraw()

    def _on_plot_hover(self, event: tk.Event[tk.Misc]) -> None:
        y = float(event.y)
        if y < self._top_padding or y > CHART_HEIGHT - self._bottom_padding:
            self._hide_hover_value()
            return

        self._hover_x = float(event.x)
        self._hover_y = y
        self._redraw_hover_value()

    def _redraw_hover_value(self) -> None:
        if self._hover_y is None or self._hover_x is None:
            self._plot_canvas.delete("hover_value")
            return

        y = self._hover_y
        value = self._map_value_from_y(y)
        y_text = f"{value:.2f} dBm"
        text_x = min(self._hover_x + 12, self._viewport_width() - 8)
        text_y = max(self._top_padding + 8, min(y - 12, CHART_HEIGHT - self._bottom_padding - 8))
        self._plot_canvas.delete("hover_value")
        self._plot_canvas.create_line(
            0,
            y,
            self._viewport_width(),
            y,
            fill=self._palette.chart_border,
            width=1,
            dash=(2, 4),
            tags="hover_value",
        )

        text_id = self._plot_canvas.create_text(
            text_x,
            text_y,
            text=y_text,
            anchor="sw",
            font=("Segoe UI", 9, "bold"),
            fill=self._palette.text,
            tags="hover_value",
        )
        x1, y1, x2, y2 = self._plot_canvas.bbox(text_id)
        self._plot_canvas.create_rectangle(
            x1 - 6,
            y1 - 4,
            x2 + 6,
            y2 + 4,
            fill=self._palette.surface_bg,
            outline=self._palette.chart_border,
            tags="hover_value",
        )
        self._plot_canvas.tag_raise(text_id)

        sample_index = self._hover_sample_index()
        if sample_index is None:
            return

        sample_x = self._x_for_index(sample_index)
        sample_value = self._samples[sample_index]
        sample_y = self._map_y(sample_value)
        self._plot_canvas.create_line(
            sample_x,
            self._top_padding,
            sample_x,
            CHART_HEIGHT - self._bottom_padding,
            fill=self._palette.grid_major,
            width=1,
            dash=(2, 4),
            tags="hover_value",
        )
        self._plot_canvas.create_oval(
            sample_x - 4,
            sample_y - 4,
            sample_x + 4,
            sample_y + 4,
            fill=self._palette.danger,
            outline=self._palette.selection_fg,
            width=1,
            tags="hover_value",
        )
        sample_text = f"{sample_value:.2f} dBm"
        sample_text_x = min(sample_x + 10, self._viewport_width() - 8)
        sample_text_y = max(self._top_padding + 18, sample_y - 14)
        sample_text_id = self._plot_canvas.create_text(
            sample_text_x,
            sample_text_y,
            text=sample_text,
            anchor="sw",
            font=("Segoe UI", 9, "bold"),
            fill=self._palette.text,
            tags="hover_value",
        )
        sx1, sy1, sx2, sy2 = self._plot_canvas.bbox(sample_text_id)
        self._plot_canvas.create_rectangle(
            sx1 - 6,
            sy1 - 4,
            sx2 + 6,
            sy2 + 4,
            fill=self._palette.surface_bg,
            outline=self._palette.chart_border,
            tags="hover_value",
        )
        self._plot_canvas.tag_raise(sample_text_id)

    def _hover_sample_index(self) -> int | None:
        if self._hover_x is None or not self._samples:
            return None

        relative_x = self._hover_x - self._left_padding
        estimated = int(round(relative_x / max(1.0, self._point_spacing)))
        return min(max(estimated, 0), len(self._samples) - 1)

    def _hide_hover_value(self, _event: tk.Event[tk.Misc] | None = None) -> None:
        self._hover_x = None
        self._hover_y = None
        self._plot_canvas.delete("hover_value")


class HoverTooltip:
    def __init__(self, widget: tk.Widget, text_source: callable[[], str], *, delay_ms: int = 250) -> None:
        self._widget = widget
        self._text_source = text_source
        self._delay_ms = delay_ms
        self._after_id: str | None = None
        self._tooltip: tk.Toplevel | None = None

        self._widget.bind("<Enter>", self._schedule, add="+")
        self._widget.bind("<Leave>", self._hide, add="+")
        self._widget.bind("<ButtonPress>", self._hide, add="+")

    def _schedule(self, _event: tk.Event[tk.Misc]) -> None:
        self._cancel_pending()
        self._after_id = self._widget.after(self._delay_ms, self._show)

    def _show(self) -> None:
        self._after_id = None
        text = self._text_source().strip()
        if not text or not self._widget.winfo_viewable():
            return

        self._tooltip = tk.Toplevel(self._widget)
        self._tooltip.wm_overrideredirect(True)
        self._tooltip.attributes("-topmost", True)
        label = tk.Label(
            self._tooltip,
            text=text,
            bg="#101113",
            fg="#f5f5f7",
            padx=8,
            pady=4,
            relief="solid",
            borderwidth=1,
            font=("Segoe UI", 10),
        )
        label.pack()
        x = self._widget.winfo_rootx() + (self._widget.winfo_width() // 2)
        y = self._widget.winfo_rooty() + self._widget.winfo_height() + 8
        self._tooltip.wm_geometry(f"+{x}+{y}")

    def _hide(self, _event: tk.Event[tk.Misc] | None = None) -> None:
        self._cancel_pending()
        if self._tooltip is not None:
            self._tooltip.destroy()
            self._tooltip = None

    def _cancel_pending(self) -> None:
        if self._after_id is not None:
            self._widget.after_cancel(self._after_id)
            self._after_id = None


class RFPowerMeterApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("USB RF Power Meter")
        self.root.geometry("1500x760")
        self.root.minsize(1280, 700)
        self._dark_mode = is_dark_mode()
        self._palette = build_palette(self._dark_mode)
        self.root.configure(bg=self._palette.window_bg)

        self._event_queue: "queue.Queue[tuple[str, object]]" = queue.Queue()
        self._worker: SerialWorker | None = None
        self._connected_port: str | None = None
        self._max_measurement: Measurement | None = None
        self._port_display_to_device: dict[str, str] = {}

        self.port_var = tk.StringVar()
        self.connect_button_var = tk.StringVar(value="Connect")
        self.connect_icon_var = tk.StringVar(value="🔌")
        self.settings_button_var = tk.StringVar(value="Settings")
        self.rate_var = tk.StringVar(value=next(iter(RATE_OPTIONS)))
        self.connection_var = tk.StringVar(value="Disconnected")
        self.max_var = tk.StringVar(value="--.-dBm")
        self.max_uw_var = tk.StringVar(value="--,--uW")
        self.dbm_var = tk.StringVar(value="--.-dBm")
        self.uw_var = tk.StringVar(value="--,--uW")
        self.waveform_var = tk.StringVar(value="Waiting for data")
        self._settings_visible = False
        self._metric_labels: list[tuple[tk.Label, str]] = []
        self._toolbar_tooltips: list[HoverTooltip] = []
        self._sync_entries: list[SyncEntry] = []
        self._sync_editor: tk.Entry | None = None
        self._sync_editor_item_id: str | None = None
        self._sync_editor_column: str | None = None

        self._build_ui()
        self._apply_theme()
        self.refresh_ports()
        self.root.after(POLL_INTERVAL_MS, self._poll_events)
        self.root.after(APPEARANCE_POLL_INTERVAL_MS, self._poll_appearance)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _build_ui(self) -> None:
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(1, weight=1)

        toolbar = ttk.Frame(self.root, padding=(12, 12, 12, 0), style="App.TFrame")
        toolbar.grid(row=0, column=0, sticky="ew")
        toolbar.columnconfigure(4, weight=1)

        ttk.Label(toolbar, text="Port").grid(row=0, column=0, sticky="w")
        self.port_combo = ttk.Combobox(toolbar, textvariable=self.port_var, state="readonly", width=14)
        self.port_combo.grid(row=0, column=1, sticky="w", padx=(8, 8))

        self.refresh_ports_button = self._create_toolbar_button(
            toolbar,
            icon="↻",
            tooltip="Refresh Ports",
            command=self.refresh_ports,
        )
        self.refresh_ports_button.grid(row=0, column=2, sticky="w")

        self.connect_button = self._create_toolbar_button(
            toolbar,
            textvariable=self.connect_icon_var,
            tooltip=self.connect_button_var,
            command=self.toggle_connection,
        )
        self.connect_button.grid(row=0, column=3, sticky="w", padx=(8, 8))

        self.settings_button = self._create_toolbar_button(
            toolbar,
            icon="⚙",
            tooltip=self.settings_button_var,
            command=self.toggle_settings_panel,
        )
        self.settings_button.grid(
            row=0,
            column=4,
            sticky="e",
        )

        self.content = ttk.Frame(self.root, padding=12, style="App.TFrame")
        self.content.grid(row=1, column=0, sticky="nsew")
        self.content.columnconfigure(0, weight=1)
        self.content.rowconfigure(0, weight=1)

        self.left_panel = ttk.Frame(self.content, style="App.TFrame")
        self.left_panel.grid(row=0, column=0, sticky="nsew")
        self.left_panel.columnconfigure(0, weight=1)
        self.left_panel.rowconfigure(1, weight=1)

        self.right_panel = ttk.Frame(self.content, padding=(12, 0, 0, 0), style="App.TFrame")
        self.right_panel.columnconfigure(0, weight=1)
        self.right_panel.rowconfigure(3, weight=1)

        self._build_metrics(self.left_panel)

        chart_frame = ttk.Frame(self.left_panel, style="App.TFrame")
        chart_frame.grid(row=1, column=0, sticky="nsew", pady=(12, 0))
        chart_frame.columnconfigure(0, weight=1)
        chart_frame.rowconfigure(0, weight=1)

        self.chart = SignalChart(chart_frame, self._palette)
        self.chart.grid(row=0, column=0, sticky="nsew")

        self.chart_toolbar_overlay = ttk.Frame(chart_frame, style="ChartOverlay.TFrame", padding=4)
        self.chart_toolbar_overlay.place(relx=1.0, x=-48, y=10, anchor="ne")
        self.clear_chart_button = self._create_toolbar_button(
            self.chart_toolbar_overlay,
            icon="🗑",
            tooltip="Clear Chart",
            command=self.clear_chart,
            width=2,
            style="ChartToolbar.TButton",
        )
        self.clear_chart_button.grid(row=0, column=0, sticky="w")
        self.load_chart_button = self._create_toolbar_button(
            self.chart_toolbar_overlay,
            icon="📂",
            tooltip="Load Chart",
            command=self.load_chart,
            width=2,
            style="ChartToolbar.TButton",
        )
        self.load_chart_button.grid(row=0, column=1, sticky="w", padx=(4, 0))
        self.save_chart_button = self._create_toolbar_button(
            self.chart_toolbar_overlay,
            icon="🗃️",
            tooltip="Save Chart",
            command=self.save_chart,
            width=2,
            style="ChartToolbar.TButton",
        )
        self.save_chart_button.grid(row=0, column=2, sticky="w", padx=(4, 0))
        self.export_chart_button = self._create_toolbar_button(
            self.chart_toolbar_overlay,
            icon="🖼️",
            tooltip="Export Chart",
            command=self.export_chart,
            width=2,
            style="ChartToolbar.TButton",
        )
        self.export_chart_button.grid(row=0, column=3, sticky="w", padx=(4, 0))
        self.zoom_out_button = self._create_toolbar_button(
            self.chart_toolbar_overlay,
            icon="－",
            tooltip="Zoom Out",
            command=self.zoom_out_chart,
            width=2,
            style="ChartToolbar.TButton",
        )
        self.zoom_out_button.grid(row=0, column=4, sticky="w", padx=(4, 0))
        self.zoom_fit_button = self._create_toolbar_button(
            self.chart_toolbar_overlay,
            icon="🔎",
            tooltip="Zoom to Fit",
            command=self.zoom_to_fit_chart,
            width=2,
            style="ChartToolbar.TButton",
        )
        self.zoom_fit_button.grid(row=0, column=5, sticky="w", padx=(4, 0))
        self.reset_zoom_button = self._create_toolbar_button(
            self.chart_toolbar_overlay,
            icon="◎",
            tooltip="Reset Zoom",
            command=self.reset_chart_zoom,
            width=2,
            style="ChartToolbar.TButton",
        )
        self.reset_zoom_button.grid(row=0, column=6, sticky="w", padx=(4, 0))

        self._build_controls(self.right_panel)
        self._update_chart_controls_state()
        self._set_settings_panel_visible(False)

    def _build_metrics(self, parent: ttk.Frame) -> None:
        metrics = ttk.Frame(parent, style="App.TFrame")
        metrics.grid(row=0, column=0, sticky="ew")
        metrics.columnconfigure(0, weight=1)
        metrics.columnconfigure(1, weight=1)

        dbm_frame = ttk.Frame(metrics, padding=8, style="Card.TFrame")
        dbm_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        dbm_frame.columnconfigure(0, weight=1)
        dbm_frame.columnconfigure(1, weight=0)
        self._make_metric_label(dbm_frame, text="Power(dBm):", font=("Segoe UI", 15, "bold")).grid(
            row=0, column=0, sticky="w"
        )
        self._make_metric_label(dbm_frame, text="MAX", font=("Segoe UI", 13), fg_role="danger").grid(
            row=0, column=1, sticky="e", padx=(16, 0)
        )
        self._make_metric_label(dbm_frame, textvariable=self.dbm_var, font=("Segoe UI", 28, "bold")).grid(
            row=1, column=0, sticky="w"
        )
        self._make_metric_label(dbm_frame, textvariable=self.max_var, font=("Segoe UI", 16, "bold"), fg_role="danger").grid(
            row=1, column=1, sticky="e", padx=(16, 0)
        )

        uw_frame = ttk.Frame(metrics, padding=8, style="Card.TFrame")
        uw_frame.grid(row=0, column=1, sticky="nsew")
        uw_frame.columnconfigure(0, weight=1)
        uw_frame.columnconfigure(1, weight=0)
        self._make_metric_label(uw_frame, text="Power(W):", font=("Segoe UI", 15, "bold")).grid(
            row=0, column=0, sticky="w"
        )
        self._make_metric_label(uw_frame, text="MAX", font=("Segoe UI", 13), fg_role="danger").grid(
            row=0, column=1, sticky="e", padx=(16, 0)
        )
        self._make_metric_label(uw_frame, textvariable=self.uw_var, font=("Segoe UI", 28, "bold")).grid(
            row=1, column=0, sticky="w"
        )
        self._make_metric_label(uw_frame, textvariable=self.max_uw_var, font=("Segoe UI", 16, "bold"), fg_role="danger").grid(
            row=1, column=1, sticky="e", padx=(16, 0)
        )

    def _build_controls(self, parent: ttk.Frame) -> None:
        command_frame = ttk.LabelFrame(parent, text="Device Commands", padding=12, style="App.TLabelframe")
        command_frame.grid(row=0, column=0, sticky="ew")
        command_frame.columnconfigure(1, weight=1)

        ttk.Label(command_frame, text="Refresh rate").grid(row=0, column=0, sticky="w")
        self.refresh_rate_combo = ttk.Combobox(
            command_frame,
            textvariable=self.rate_var,
            state="readonly",
            values=list(RATE_OPTIONS.keys()),
        )
        self.refresh_rate_combo.grid(row=0, column=1, sticky="ew", padx=(8, 0))
        self.refresh_rate_combo.bind("<<ComboboxSelected>>", self._on_rate_changed)

        self.synchronize_button = ttk.Button(command_frame, text="Synchronize", command=self.synchronize)
        self.synchronize_button.grid(
            row=1,
            column=0,
            sticky="ew",
            pady=(10, 0),
        )

        sync_frame = ttk.LabelFrame(parent, text="Synchronized Profiles", padding=12, style="App.TLabelframe")
        sync_frame.grid(row=1, column=0, sticky="nsew", pady=(12, 0))
        sync_frame.columnconfigure(0, weight=1)
        sync_frame.rowconfigure(0, weight=1)

        self.sync_table = ttk.Treeview(
            sync_frame,
            columns=("index", "frequency", "offset"),
            show="headings",
            height=10,
        )
        self.sync_table.heading("index", text="#")
        self.sync_table.heading("frequency", text="Frequency (MHz)")
        self.sync_table.heading("offset", text="Offset (dBm)")
        self.sync_table.column("index", anchor="center", width=48, stretch=False)
        self.sync_table.column("frequency", anchor="center", width=130)
        self.sync_table.column("offset", anchor="center", width=110)
        self.sync_table.grid(row=0, column=0, sticky="nsew")
        self.sync_table.bind("<Button-1>", self._on_sync_profile_single_click, add="+")
        self.sync_table.bind("<Double-1>", self._on_sync_profile_double_click)

        sync_scrollbar = ttk.Scrollbar(sync_frame, orient="vertical", command=self.sync_table.yview)
        sync_scrollbar.grid(row=0, column=1, sticky="ns")
        self.sync_table.configure(yscrollcommand=sync_scrollbar.set)

        waveform_frame = ttk.LabelFrame(parent, text="Latest Waveform Packet", padding=12, style="App.TLabelframe")
        waveform_frame.grid(row=2, column=0, sticky="ew", pady=(12, 0))
        ttk.Label(waveform_frame, textvariable=self.waveform_var, font=("Consolas", 12)).pack(anchor="w")

        log_frame = ttk.LabelFrame(parent, text="Message Log", padding=12, style="App.TLabelframe")
        log_frame.grid(row=3, column=0, sticky="nsew", pady=(12, 0))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        self.log_text = tk.Text(log_frame, height=10, wrap="word", state="disabled")
        self.log_text.grid(row=0, column=0, sticky="nsew")
        log_scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        log_scrollbar.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=log_scrollbar.set)

        self._set_device_controls_connected(False)

    def _make_metric_label(
        self,
        parent: tk.Misc,
        *,
        font: tuple[str, int] | tuple[str, int, str],
        fg_role: str = "text",
        text: str | None = None,
        textvariable: tk.StringVar | None = None,
    ) -> tk.Label:
        foreground = self._palette.danger if fg_role == "danger" else self._palette.text
        label = tk.Label(
            parent,
            text=text,
            textvariable=textvariable,
            font=font,
            bg=self._palette.card_bg,
            fg=foreground,
        )
        self._metric_labels.append((label, fg_role))
        return label

    def _create_toolbar_button(
        self,
        parent: ttk.Frame,
        *,
        command: object,
        tooltip: str | tk.StringVar,
        icon: str | None = None,
        textvariable: tk.StringVar | None = None,
        width: int = 3,
        style: str = "Toolbar.TButton",
    ) -> ttk.Button:
        button = ttk.Button(
            parent,
            text=icon,
            textvariable=textvariable,
            command=command,
            width=width,
            style=style,
        )
        tooltip_source = tooltip.get if isinstance(tooltip, tk.StringVar) else (lambda: tooltip)
        self._toolbar_tooltips.append(HoverTooltip(button, tooltip_source))
        return button

    def _apply_theme(self) -> None:
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        self.root.configure(bg=self._palette.window_bg)
        style.configure(".", background=self._palette.window_bg, foreground=self._palette.text)
        style.configure("App.TFrame", background=self._palette.window_bg)
        style.configure("Card.TFrame", background=self._palette.card_bg)
        style.configure("App.TLabelframe", background=self._palette.surface_bg, foreground=self._palette.text)
        style.configure("App.TLabelframe.Label", background=self._palette.surface_bg, foreground=self._palette.text)
        style.configure("TLabel", background=self._palette.window_bg, foreground=self._palette.text)
        style.configure("TButton", background=self._palette.surface_bg, foreground=self._palette.text)
        style.map("TButton", background=[("active", self._palette.card_bg), ("disabled", self._palette.surface_bg)])
        style.configure(
            "Toolbar.TButton",
            background=self._palette.surface_bg,
            foreground=self._palette.text,
            font=("Segoe UI Symbol", 18, "bold"),
            padding=(2, 2),
        )
        style.configure(
            "ChartToolbar.TButton",
            background=self._palette.surface_bg,
            foreground=self._palette.text,
            font=("Segoe UI Symbol", 12, "bold"),
            padding=(1, 1),
            borderwidth=0,
            relief="flat",
        )
        style.configure("ChartOverlay.TFrame", background=self._palette.surface_bg)
        style.map(
            "Toolbar.TButton",
            background=[("active", self._palette.card_bg), ("disabled", self._palette.surface_bg)],
            foreground=[("disabled", self._palette.muted_text)],
        )
        style.map(
            "ChartToolbar.TButton",
            background=[("active", self._palette.card_bg), ("disabled", self._palette.surface_bg)],
            foreground=[("disabled", self._palette.muted_text)],
        )
        style.configure(
            "TCombobox",
            fieldbackground=self._palette.entry_bg,
            background=self._palette.surface_bg,
            foreground=self._palette.entry_fg,
            arrowcolor=self._palette.text,
        )
        style.map(
            "TCombobox",
            fieldbackground=[("readonly", self._palette.entry_bg)],
            selectbackground=[("readonly", self._palette.selection_bg)],
            selectforeground=[("readonly", self._palette.selection_fg)],
        )
        style.configure(
            "Treeview",
            background=self._palette.log_bg,
            fieldbackground=self._palette.log_bg,
            foreground=self._palette.text,
        )
        style.configure("Treeview.Heading", background=self._palette.surface_bg, foreground=self._palette.text)

        for label, fg_role in self._metric_labels:
            label.configure(
                bg=self._palette.card_bg,
                fg=self._palette.danger if fg_role == "danger" else self._palette.text,
            )

        self.log_text.configure(
            bg=self._palette.log_bg,
            fg=self._palette.text,
            insertbackground=self._palette.text,
            selectbackground=self._palette.selection_bg,
            selectforeground=self._palette.selection_fg,
        )
        self.chart.apply_palette(self._palette)

    def _poll_appearance(self) -> None:
        dark_mode = is_dark_mode()
        if dark_mode != self._dark_mode:
            self._dark_mode = dark_mode
            self._palette = build_palette(self._dark_mode)
            self._apply_theme()
        self.root.after(APPEARANCE_POLL_INTERVAL_MS, self._poll_appearance)

    def refresh_ports(self) -> None:
        port_infos = sorted(
            (port for port in list_ports.comports() if port.device not in IGNORED_PORTS),
            key=lambda port: port.device,
        )
        display_ports = [self._format_port_label(port) for port in port_infos]
        self._port_display_to_device = {
            display_label: port.device for display_label, port in zip(display_ports, port_infos, strict=False)
        }
        self.port_combo.configure(values=display_ports)
        combo_width = max((len(port) for port in display_ports), default=14) + 2
        self.port_combo.configure(width=combo_width)
        if display_ports:
            if self.port_var.get() not in display_ports:
                self.port_var.set(display_ports[0])
        else:
            self.port_var.set("")
        detected_ports = ", ".join(display_ports) if display_ports else "none"
        self._append_log(f"Detected ports: {detected_ports}")

    def connect(self) -> None:
        if self._worker is not None:
            self._append_log("A serial connection is already active.")
            return

        selected_label = self.port_var.get().strip()
        if not selected_label:
            messagebox.showerror("USB RF Power Meter", "Select a COM port before connecting.")
            return

        selected_port = self._port_display_to_device.get(selected_label, selected_label)

        self._append_log(f"Opening {selected_port} at 9600 baud.")
        self.connection_var.set(f"Connecting to {selected_port}...")
        self._set_device_controls_connected(False)
        self.refresh_ports_button.configure(state="disabled")
        self._worker = SerialWorker(selected_port, self._event_queue)
        self._worker.start()

    def _format_port_label(self, port: object) -> str:
        device = getattr(port, "device", "")
        description = getattr(port, "description", "")
        if platform.system() == "Windows" and description and description != device and description.lower() != "n/a":
            return f"{device} - {description}"
        return device

    def disconnect(self) -> None:
        if self._worker is None:
            return
        self._append_log("Disconnect requested.")
        self._worker.stop()

    def toggle_connection(self) -> None:
        if self._worker is None:
            self.connect()
            return
        self.disconnect()

    def synchronize(self) -> None:
        self._send_command(SYNC_COMMAND)

    def clear_chart(self) -> None:
        self.chart.clear()
        self._reset_chart_metrics()
        self._update_chart_controls_state()
        self._append_log("Chart cleared.")

    def save_chart(self) -> None:
        samples = self.chart.samples()
        if not samples:
            messagebox.showinfo("USB RF Power Meter", "There is no chart data to save yet.")
            return

        file_path = filedialog.asksaveasfilename(
            parent=self.root,
            title="Save Chart Data",
            defaultextension=CHART_FILE_EXTENSION,
            filetypes=[("Chart CSV", "*.csv"), ("All files", "*.*")],
            initialfile="rf-power-chart.csv",
        )
        if not file_path:
            return

        output = Path(file_path)
        try:
            output.write_text(
                "index,dbm\n" + "\n".join(f"{index},{value:.6f}" for index, value in enumerate(samples)) + "\n",
                encoding="utf-8",
            )
        except OSError as exc:
            messagebox.showerror("USB RF Power Meter", f"Could not save chart data:\n{exc}")
            return
        self._append_log(f"Saved {len(samples)} chart samples to {output}.")

    def load_chart(self) -> None:
        file_path = filedialog.askopenfilename(
            parent=self.root,
            title="Load Chart Data",
            filetypes=[("Chart CSV", "*.csv"), ("All files", "*.*")],
        )
        if not file_path:
            return

        try:
            samples = self._read_chart_samples(Path(file_path))
        except ValueError as exc:
            messagebox.showerror("USB RF Power Meter", str(exc))
            return
        except OSError as exc:
            messagebox.showerror("USB RF Power Meter", f"Could not read chart data:\n{exc}")
            return

        self.chart.set_samples(samples)
        self._restore_loaded_chart_state(samples)
        self._update_chart_controls_state()
        self._append_log(f"Loaded {len(samples)} chart samples from {file_path}.")

    def export_chart(self) -> None:
        if not self.chart.has_samples():
            messagebox.showinfo("USB RF Power Meter", "There is no chart data to export yet.")
            return

        file_path = filedialog.asksaveasfilename(
            parent=self.root,
            title="Export Chart",
            defaultextension=EXPORT_FILE_EXTENSION,
            filetypes=[("JPEG image", "*.jpg"), ("All files", "*.*")],
            initialfile="rf-power-chart.jpg",
        )
        if not file_path:
            return

        try:
            self.chart.export_to_jpeg(Path(file_path))
        except subprocess.CalledProcessError as exc:
            messagebox.showerror("USB RF Power Meter", f"Could not export chart image:\n{exc}")
            return
        except OSError as exc:
            messagebox.showerror("USB RF Power Meter", f"Could not export chart image:\n{exc}")
            return

        self._append_log(f"Exported chart image to {file_path}.")

    def zoom_out_chart(self) -> None:
        self.chart.zoom_out()
        self._update_chart_controls_state()

    def zoom_to_fit_chart(self) -> None:
        self.chart.zoom_to_fit()
        self._update_chart_controls_state()

    def reset_chart_zoom(self) -> None:
        self.chart.reset_zoom()
        self._update_chart_controls_state()

    def toggle_settings_panel(self) -> None:
        self._set_settings_panel_visible(not self._settings_visible)

    def _on_rate_changed(self, _event: object) -> None:
        command = RATE_OPTIONS[self.rate_var.get()]
        if self._worker is not None:
            self._send_command(command)
        else:
            self._append_log(f"Refresh rate selected: {command} (will be sent after connect)")

    def _send_command(self, command: str) -> None:
        if self._worker is None:
            self._append_log(f"Cannot send {command}: device is disconnected.")
            return
        self._worker.send_command(command)

    def _read_chart_samples(self, file_path: Path) -> list[float]:
        lines = [line.strip() for line in file_path.read_text(encoding="utf-8").splitlines() if line.strip()]
        if not lines:
            raise ValueError("The selected chart file is empty.")

        samples: list[float] = []
        for line in lines:
            if line.lower() == "index,dbm":
                continue

            if "," in line:
                _, dbm_text = line.split(",", 1)
            else:
                dbm_text = line

            try:
                samples.append(float(dbm_text))
            except ValueError as exc:
                raise ValueError(f"Invalid chart sample in {file_path.name}: {line}") from exc

        if not samples:
            raise ValueError("The selected chart file does not contain any samples.")
        return samples

    def _restore_loaded_chart_state(self, samples: list[float]) -> None:
        latest_dbm = samples[-1]
        max_dbm = max(samples)
        self.dbm_var.set(format_dbm(latest_dbm))
        self.uw_var.set(format_microwatts(dbm_to_microwatts(latest_dbm)))
        self.max_var.set(format_dbm(max_dbm))
        self.max_uw_var.set(format_microwatts(dbm_to_microwatts(max_dbm)))
        self.waveform_var.set(f"Loaded file with {len(samples)} samples")
        self._max_measurement = Measurement(
            raw="loaded-from-file",
            dbm=max_dbm,
            microwatts=dbm_to_microwatts(max_dbm),
        )

    def _reset_chart_metrics(self) -> None:
        self._max_measurement = None
        self.max_var.set("--.-dBm")
        self.max_uw_var.set("--,--uW")
        self.dbm_var.set("--.-dBm")
        self.uw_var.set("--,--uW")
        self.waveform_var.set("Waiting for data")

    def _poll_events(self) -> None:
        measurement_batch: list[Measurement] = []
        while True:
            try:
                event_type, payload = self._event_queue.get_nowait()
            except queue.Empty:
                break
            if event_type == "measurements":
                measurement_batch.extend(payload)  # type: ignore[arg-type]
                continue
            self._handle_event(event_type, payload)

        if measurement_batch:
            self._update_measurements(measurement_batch)

        self.root.after(POLL_INTERVAL_MS, self._poll_events)

    def _handle_event(self, event_type: str, payload: object) -> None:
        if event_type == "connected":
            port = str(payload)
            self._connected_port = port
            self.connect_button_var.set("Disconnect")
            self.connect_icon_var.set("❌")
            self.connection_var.set(f"Connected: {port}")
            self._set_device_controls_connected(True)
            self._append_log(f"Connected to {port}.")
            self._send_command(RATE_OPTIONS[self.rate_var.get()])
            self.root.after(CONNECT_SYNC_DELAY_MS, self._synchronize_if_connected)
            return

        if event_type == "disconnected":
            port = str(payload)
            if self._connected_port == port:
                self._connected_port = None
            self._worker = None
            self.connect_button_var.set("Connect")
            self.connect_icon_var.set("🔌")
            self.connection_var.set("Disconnected")
            self._set_device_controls_connected(False)
            self._append_log(f"Disconnected from {port}.")
            return

        if event_type == "error":
            self._append_log(str(payload))
            self.connect_button_var.set("Connect")
            self.connect_icon_var.set("🔌")
            self.connection_var.set("Error")
            self._set_device_controls_connected(False)
            return

        if event_type == "command_sent":
            self._append_log(f"> {payload}")
            return

        if event_type == "sync":
            self._update_sync_entries(payload)  # type: ignore[arg-type]
            return

        if event_type == "log":
            self._append_log(str(payload))

    def _update_measurement(self, measurement: Measurement) -> None:
        self.dbm_var.set(format_dbm(measurement.dbm))
        self.uw_var.set(format_microwatts(measurement.microwatts))
        self.waveform_var.set(measurement.raw)

        if self._max_measurement is None or measurement.dbm > self._max_measurement.dbm:
            self._max_measurement = measurement
            self.max_var.set(format_dbm(measurement.dbm))
            self.max_uw_var.set(format_microwatts(measurement.microwatts))

    def _update_measurements(self, measurements: list[Measurement]) -> None:
        latest = measurements[-1]
        self._update_measurement(latest)
        self.chart.append_many(self._downsample_measurements(measurements))
        self._update_chart_controls_state()

    def _downsample_measurements(self, measurements: list[Measurement]) -> list[float]:
        if len(measurements) <= MAX_DRAWN_SAMPLES_PER_BATCH:
            return [measurement.dbm for measurement in measurements]

        step = ceil(len(measurements) / MAX_DRAWN_SAMPLES_PER_BATCH)
        sampled = [measurements[index].dbm for index in range(0, len(measurements), step)]
        if sampled[-1] != measurements[-1].dbm:
            sampled.append(measurements[-1].dbm)
        return sampled

    def _update_sync_entries(self, entries: list[SyncEntry]) -> None:
        self._close_sync_editor()
        self._sync_entries = list(entries)
        for item_id in self.sync_table.get_children():
            self.sync_table.delete(item_id)

        for index, entry in enumerate(entries):
            self.sync_table.insert(
                "",
                "end",
                values=(str(index), str(entry.frequency_mhz), f"{entry.offset_dbm:+.1f}"),
            )

        formatted = ", ".join(f"{entry.frequency_mhz:04d}MHz {entry.offset_dbm:+.1f}dBm" for entry in entries)
        self._append_log(f"< Sync profiles: {formatted}")

    def _on_sync_profile_double_click(self, event: tk.Event[tk.Misc]) -> None:
        item_id = self.sync_table.identify_row(event.y)
        if not item_id:
            return
        if self.sync_table.identify_column(event.x) in {"#2", "#3"}:
            return

        try:
            row_index = self.sync_table.index(item_id)
        except (IndexError, tk.TclError):
            return

        self._send_sync_profile_command(row_index)

    def _on_sync_profile_single_click(self, event: tk.Event[tk.Misc]) -> None:
        item_id = self.sync_table.identify_row(event.y)
        column = self.sync_table.identify_column(event.x)
        if not item_id or column not in {"#2", "#3"}:
            self._close_sync_editor()
            return

        self.root.after_idle(lambda: self._open_sync_editor(item_id, column))

    def _open_sync_editor(self, item_id: str, column: str) -> None:
        self._close_sync_editor()
        bbox = self.sync_table.bbox(item_id, column)
        if not bbox:
            return

        x, y, width, height = bbox
        values = self.sync_table.item(item_id, "values")
        value_index = int(column[1:]) - 1
        if value_index >= len(values):
            return

        editor = tk.Entry(
            self.sync_table,
            justify="center",
            relief="solid",
            borderwidth=1,
        )
        editor.insert(0, str(values[value_index]))
        editor.select_range(0, "end")
        editor.place(x=x, y=y, width=width, height=height)
        editor.focus_set()
        editor.bind("<Return>", lambda _event: self._commit_sync_editor())
        editor.bind("<Escape>", lambda _event: self._close_sync_editor())
        editor.bind("<FocusOut>", lambda _event: self._close_sync_editor())

        self._sync_editor = editor
        self._sync_editor_item_id = item_id
        self._sync_editor_column = column

    def _commit_sync_editor(self) -> None:
        if self._sync_editor is None or self._sync_editor_item_id is None or self._sync_editor_column is None:
            return

        try:
            row_index = self.sync_table.index(self._sync_editor_item_id)
            entry = self._sync_entries[row_index]
        except (IndexError, tk.TclError):
            self._close_sync_editor()
            return

        raw_value = self._sync_editor.get().strip()
        try:
            if self._sync_editor_column == "#2":
                entry.frequency_mhz = int(raw_value)
            elif self._sync_editor_column == "#3":
                entry.offset_dbm = round(float(raw_value), 1)
        except ValueError:
            self._append_log(f"Invalid sync profile value: {raw_value}")
            self._close_sync_editor()
            return

        self._update_sync_table_row(row_index)
        self._send_sync_profile_command(row_index)
        self._close_sync_editor()

    def _close_sync_editor(self) -> None:
        if self._sync_editor is not None:
            self._sync_editor.destroy()
        self._sync_editor = None
        self._sync_editor_item_id = None
        self._sync_editor_column = None

    def _update_sync_table_row(self, row_index: int) -> None:
        try:
            item_id = self.sync_table.get_children()[row_index]
            entry = self._sync_entries[row_index]
        except IndexError:
            return

        self.sync_table.item(
            item_id,
            values=(str(row_index), str(entry.frequency_mhz), f"{entry.offset_dbm:+.1f}"),
        )

    def _send_sync_profile_command(self, row_index: int) -> None:
        try:
            entry = self._sync_entries[row_index]
            prefix = SYNC_PROFILE_COMMAND_PREFIXES[row_index]
        except IndexError:
            return

        command = f"{prefix}{entry.frequency_mhz:04d}{entry.offset_dbm:+05.1f}"
        self._send_command(command)

    def _append_log(self, message: str) -> None:
        self.log_text.configure(state="normal")
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")
        if int(float(self.log_text.index("end-1c").split(".")[0])) > 200:
            self.log_text.delete("1.0", "30.0")
        self.log_text.configure(state="disabled")

    def _set_device_controls_connected(self, connected: bool) -> None:
        self.refresh_ports_button.configure(state="disabled" if connected else "normal")
        self.refresh_rate_combo.configure(state="readonly" if connected else "disabled")
        self.synchronize_button.configure(state="normal" if connected else "disabled")

    def _update_chart_controls_state(self) -> None:
        self.clear_chart_button.configure(state="normal" if self.chart.has_samples() else "disabled")
        self.save_chart_button.configure(state="normal" if self.chart.has_samples() else "disabled")
        self.export_chart_button.configure(state="normal" if self.chart.has_samples() else "disabled")
        self.zoom_out_button.configure(state="normal" if self.chart.can_zoom_out() else "disabled")
        self.zoom_fit_button.configure(state="normal" if self.chart.can_zoom_to_fit() else "disabled")
        self.reset_zoom_button.configure(state="normal" if self.chart.can_reset_zoom() else "disabled")

    def _set_settings_panel_visible(self, visible: bool) -> None:
        self._settings_visible = visible
        self.settings_button_var.set("Hide Settings" if visible else "Settings")
        if visible:
            self.right_panel.grid(row=0, column=1, sticky="nsew")
            self.content.columnconfigure(1, weight=0)
        else:
            self.right_panel.grid_remove()
            self.content.columnconfigure(1, weight=0)

    def _synchronize_if_connected(self) -> None:
        if self._worker is not None:
            self.synchronize()

    def _on_close(self) -> None:
        if self._worker is not None:
            self._worker.stop()
        self.root.destroy()


def main() -> int:
    root = tk.Tk()
    app = RFPowerMeterApp(root)
    root.mainloop()
    return 0
