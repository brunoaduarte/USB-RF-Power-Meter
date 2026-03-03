from __future__ import annotations

from dataclasses import dataclass
import re

STREAM_PACKET_RE = re.compile(r"a[+-]?\d{3}00\d+uA")
SYNC_RESPONSE_RE = re.compile(r"R(?:\d{4}[+-]\d{2}\.\d)+(?:[A-Z])?")
SYNC_FIELD_RE = re.compile(r"(\d{4})([+-]\d{2}\.\d)")


@dataclass(slots=True)
class Measurement:
    raw: str
    dbm: float
    microwatts: float


@dataclass(slots=True)
class SyncEntry:
    frequency_mhz: int
    offset_dbm: float


def parse_stream_packet(raw: str) -> Measurement:
    text = raw.strip()
    match = re.fullmatch(r"a(?P<dbm>[+-]?\d{3})00(?P<uw>\d+)uA", text)
    if not match:
        raise ValueError(f"Unsupported stream packet: {raw!r}")

    dbm = int(match.group("dbm")) / 10.0
    microwatts = int(match.group("uw")) / 100.0
    return Measurement(raw=text, dbm=dbm, microwatts=microwatts)


def parse_sync_response(raw: str) -> list[SyncEntry]:
    text = raw.strip()
    if not SYNC_RESPONSE_RE.fullmatch(text):
        raise ValueError(f"Unsupported sync response: {raw!r}")

    if text[-1].isalpha():
        text = text[:-1]

    entries = [
        SyncEntry(frequency_mhz=int(freq), offset_dbm=float(offset))
        for freq, offset in SYNC_FIELD_RE.findall(text[1:])
    ]
    if not entries:
        raise ValueError(f"Sync response did not contain entries: {raw!r}")
    return entries


class SerialProtocolParser:
    def __init__(self) -> None:
        self._buffer = ""

    def feed(self, text: str) -> list[tuple[str, object]]:
        self._buffer += text.replace("\x00", "")
        events: list[tuple[str, object]] = []

        self._consume_line_messages(events)

        while True:
            packet_match = STREAM_PACKET_RE.search(self._buffer)
            if not packet_match:
                break

            prefix = self._buffer[: packet_match.start()].strip()
            if prefix:
                events.extend(self._parse_message(prefix))

            packet_text = packet_match.group(0)
            events.append(("measurement", parse_stream_packet(packet_text)))
            self._buffer = self._buffer[packet_match.end() :]
            self._consume_line_messages(events)

        self._trim_buffer()
        return events

    def _consume_line_messages(self, events: list[tuple[str, object]]) -> None:
        while True:
            newline_match = re.search(r"\r\n|\r|\n", self._buffer)
            if not newline_match:
                return

            line = self._buffer[: newline_match.start()].strip()
            self._buffer = self._buffer[newline_match.end() :]
            if line:
                events.extend(self._parse_message(line))

    def _parse_message(self, text: str) -> list[tuple[str, object]]:
        if not text:
            return []

        if SYNC_RESPONSE_RE.fullmatch(text):
            try:
                return [("sync", parse_sync_response(text))]
            except ValueError:
                return [("log", f"Unparsed sync payload: {text}")]

        if text.startswith("a"):
            try:
                return [("measurement", parse_stream_packet(text))]
            except ValueError:
                return [("log", f"Unparsed waveform payload: {text}")]

        return [("log", f"Serial text: {text}")]

    def _trim_buffer(self) -> None:
        if len(self._buffer) <= 2048:
            return

        last_packet = max(self._buffer.rfind("a"), self._buffer.rfind("R"))
        self._buffer = self._buffer[last_packet:] if last_packet >= 0 else ""
