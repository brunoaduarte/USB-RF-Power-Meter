from __future__ import annotations

import queue
import threading
from typing import Any

import serial

from usb_rf_power_meter.protocol import SerialProtocolParser


class SerialWorker(threading.Thread):
    def __init__(self, port: str, event_queue: "queue.Queue[tuple[str, Any]]") -> None:
        super().__init__(daemon=True)
        self._port = port
        self._event_queue = event_queue
        self._stop_event = threading.Event()
        self._command_queue: "queue.Queue[str]" = queue.Queue()

    def run(self) -> None:
        try:
            connection = serial.Serial(
                port=self._port,
                baudrate=9600,
                timeout=0.2,
                write_timeout=1,
            )
        except serial.SerialException as exc:
            self._event_queue.put(("error", f"Unable to open {self._port}: {exc}"))
            self._event_queue.put(("disconnected", self._port))
            return

        parser = SerialProtocolParser()
        self._event_queue.put(("connected", self._port))

        try:
            while not self._stop_event.is_set():
                self._flush_pending_commands(connection)

                chunk = connection.read(connection.in_waiting or 1)
                if not chunk:
                    continue

                decoded = chunk.decode("ascii", errors="ignore")
                events = parser.feed(decoded)
                measurements = [payload for event_type, payload in events if event_type == "measurement"]
                for event_type, payload in events:
                    if event_type == "measurement":
                        continue
                    self._event_queue.put((event_type, payload))
                if measurements:
                    self._event_queue.put(("measurements", measurements))
        except serial.SerialException as exc:
            self._event_queue.put(("error", f"Serial error on {self._port}: {exc}"))
        finally:
            if connection.is_open:
                connection.close()
            self._event_queue.put(("disconnected", self._port))

    def send_command(self, command: str) -> None:
        cleaned = command.strip()
        if cleaned:
            self._command_queue.put(cleaned)

    def stop(self) -> None:
        self._stop_event.set()

    def _flush_pending_commands(self, connection: serial.Serial) -> None:
        while True:
            try:
                command = self._command_queue.get_nowait()
            except queue.Empty:
                return

            connection.write(f"{command}\r\n".encode("ascii"))
            connection.flush()
            self._event_queue.put(("command_sent", command))
