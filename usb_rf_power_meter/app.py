from __future__ import annotations

import platform
import queue
import subprocess
import tkinter as tk
from tkinter import messagebox, ttk
from dataclasses import dataclass
from math import ceil

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
CONNECT_SYNC_DELAY_MS = 500
APPEARANCE_POLL_INTERVAL_MS = 1500
RATE_OPTIONS = {
    "S0 - 1s (Slow)": "S0",
    "S1 - 200 ms (Fast)": "S1",
    "S2 - 500 ns (Very fast)": "S2",
}
SYNC_COMMAND = "Read"
IGNORED_PORTS = {
    "/dev/cu.Bluetooth-Incoming-Port",
    "/dev/cu.debug-console",
    "/dev/cu.wlan-debug",
}


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


class SignalChart(ttk.Frame):
    def __init__(self, master: tk.Misc, palette: AppPalette) -> None:
        super().__init__(master)
        self._palette = palette
        self._samples: list[float] = []
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

    def _draw_axis(self) -> None:
        self._axis_canvas.delete("all")
        zero_y = self._map_y(0.0)
        self._axis_canvas.create_text(
            4,
            zero_y,
            text="dBm",
            anchor="w",
            font=("Segoe UI", 7, "bold"),
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

        self.port_var = tk.StringVar()
        self.connect_button_var = tk.StringVar(value="Connect")
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
        toolbar.columnconfigure(8, weight=1)

        ttk.Label(toolbar, text="Port").grid(row=0, column=0, sticky="w")
        self.port_combo = ttk.Combobox(toolbar, textvariable=self.port_var, state="readonly", width=14)
        self.port_combo.grid(row=0, column=1, sticky="w", padx=(8, 8))

        self.refresh_ports_button = ttk.Button(
            toolbar,
            text="↻",
            width=3,
            style="Refresh.TButton",
            command=self.refresh_ports,
        )
        self.refresh_ports_button.grid(row=0, column=2, sticky="w")

        self.connect_button = ttk.Button(toolbar, textvariable=self.connect_button_var, command=self.toggle_connection)
        self.connect_button.grid(row=0, column=3, sticky="w", padx=(8, 8))

        self.clear_chart_button = ttk.Button(toolbar, text="Clear Chart", command=self.clear_chart)
        self.clear_chart_button.grid(row=0, column=4, sticky="w")
        self.zoom_out_button = ttk.Button(toolbar, text="Zoom Out", command=self.zoom_out_chart)
        self.zoom_out_button.grid(row=0, column=5, sticky="w", padx=(8, 0))
        self.zoom_fit_button = ttk.Button(toolbar, text="Zoom to Fit", command=self.zoom_to_fit_chart)
        self.zoom_fit_button.grid(row=0, column=6, sticky="w", padx=(8, 0))
        self.reset_zoom_button = ttk.Button(toolbar, text="Reset Zoom", command=self.reset_chart_zoom)
        self.reset_zoom_button.grid(row=0, column=7, sticky="w", padx=(8, 0))
        ttk.Button(toolbar, textvariable=self.settings_button_var, command=self.toggle_settings_panel).grid(
            row=0,
            column=9,
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

        chart_frame = ttk.LabelFrame(self.left_panel, text="Waveform", style="App.TLabelframe")
        chart_frame.grid(row=1, column=0, sticky="nsew", pady=(12, 0))
        chart_frame.columnconfigure(0, weight=1)
        chart_frame.rowconfigure(0, weight=1)

        self.chart = SignalChart(chart_frame, self._palette)
        self.chart.grid(row=0, column=0, sticky="nsew")

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
            columns=("frequency", "offset"),
            show="headings",
            height=10,
        )
        self.sync_table.heading("frequency", text="Frequency (MHz)")
        self.sync_table.heading("offset", text="Offset (dBm)")
        self.sync_table.column("frequency", anchor="center", width=130)
        self.sync_table.column("offset", anchor="center", width=110)
        self.sync_table.grid(row=0, column=0, sticky="nsew")

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
            "Refresh.TButton",
            background=self._palette.surface_bg,
            foreground=self._palette.text,
            font=("Segoe UI Symbol", 20, "bold"),
            padding=(2, 0),
        )
        style.map(
            "Refresh.TButton",
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
        ports = sorted(port.device for port in list_ports.comports() if port.device not in IGNORED_PORTS)
        self.port_combo.configure(values=ports)
        combo_width = max((len(port) for port in ports), default=14)
        self.port_combo.configure(width=combo_width)
        if ports:
            if self.port_var.get() not in ports:
                self.port_var.set(ports[0])
        else:
            self.port_var.set("")
        self._append_log(f"Detected ports: {', '.join(ports) if ports else 'none'}")

    def connect(self) -> None:
        if self._worker is not None:
            self._append_log("A serial connection is already active.")
            return

        selected_port = self.port_var.get().strip()
        if not selected_port:
            messagebox.showerror("USB RF Power Meter", "Select a COM port before connecting.")
            return

        self._append_log(f"Opening {selected_port} at 9600 baud.")
        self.connection_var.set(f"Connecting to {selected_port}...")
        self._set_device_controls_connected(False)
        self.refresh_ports_button.configure(state="disabled")
        self._worker = SerialWorker(selected_port, self._event_queue)
        self._worker.start()

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
        self._max_measurement = None
        self.max_var.set("--.-dBm")
        self.max_uw_var.set("--,--uW")
        self._update_chart_controls_state()
        self._append_log("Chart cleared.")

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
            self.connection_var.set("Disconnected")
            self._set_device_controls_connected(False)
            self._append_log(f"Disconnected from {port}.")
            return

        if event_type == "error":
            self._append_log(str(payload))
            self.connect_button_var.set("Connect")
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
        for item_id in self.sync_table.get_children():
            self.sync_table.delete(item_id)

        for entry in entries:
            self.sync_table.insert(
                "",
                "end",
                values=(f"{entry.frequency_mhz:04d}", f"{entry.offset_dbm:+.1f}"),
            )

        formatted = ", ".join(f"{entry.frequency_mhz:04d}MHz {entry.offset_dbm:+.1f}dBm" for entry in entries)
        self._append_log(f"< Sync profiles: {formatted}")

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
