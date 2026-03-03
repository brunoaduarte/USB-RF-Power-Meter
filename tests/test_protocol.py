import unittest

from usb_rf_power_meter.protocol import (
    SerialProtocolParser,
    parse_stream_packet,
    parse_sync_response,
)


class ProtocolTests(unittest.TestCase):
    def test_parse_stream_packet(self) -> None:
        measurement = parse_stream_packet("a-43300004uA")
        self.assertEqual(measurement.raw, "a-43300004uA")
        self.assertEqual(measurement.dbm, -43.3)
        self.assertEqual(measurement.microwatts, 0.04)

    def test_parse_stream_packet_without_fixed_separator(self) -> None:
        measurement = parse_stream_packet("a-11407113uA")
        self.assertEqual(measurement.raw, "a-11407113uA")
        self.assertEqual(measurement.dbm, -11.4)
        self.assertEqual(measurement.microwatts, 71.13)

    def test_parse_sync_response(self) -> None:
        response = "R0006+20.00013+20.00027+20.00040+20.00433+20.00915+20.02450+20.05800+00.05800+00.0"
        entries = parse_sync_response(response)
        self.assertEqual(len(entries), 9)
        self.assertEqual(entries[0].frequency_mhz, 6)
        self.assertEqual(entries[0].offset_dbm, 20.0)
        self.assertEqual(entries[-1].frequency_mhz, 5800)
        self.assertEqual(entries[-1].offset_dbm, 0.0)

    def test_parse_sync_response_with_trailing_ack(self) -> None:
        response = "R0006+20.00013+20.00027+20.00040+20.00433+20.00915+20.02450+20.05800+00.05800+00.0A"
        entries = parse_sync_response(response)
        self.assertEqual(len(entries), 9)
        self.assertEqual(entries[0].frequency_mhz, 6)
        self.assertEqual(entries[-1].frequency_mhz, 5800)

    def test_parser_handles_mixed_stream_and_sync_payloads(self) -> None:
        parser = SerialProtocolParser()
        events = parser.feed("Read\r\nR0006+20.00013+20.0\r\na-43300004uA")

        self.assertEqual(events[0], ("log", "Serial text: Read"))
        self.assertEqual(events[1][0], "sync")
        self.assertEqual(events[2][0], "measurement")

    def test_parser_handles_sync_payload_with_trailing_ack(self) -> None:
        parser = SerialProtocolParser()
        events = parser.feed("R0006+20.00013+20.00027+20.00040+20.00433+20.00915+20.02450+20.05800+00.05800+00.0A\r\n")

        self.assertEqual(len(events), 1)
        self.assertEqual(events[0][0], "sync")

    def test_parser_handles_back_to_back_stream_packets(self) -> None:
        parser = SerialProtocolParser()
        payload = "a-11407113uAa-12505560uAa-14903219uAa-08613789uAa-11007812uAa-11506933uAa-09012427uAa-09211798uAa-19601078uA"
        events = parser.feed(payload)

        self.assertEqual(len(events), 9)
        self.assertTrue(all(event[0] == "measurement" for event in events))
        self.assertEqual(events[0][1].dbm, -11.4)
        self.assertEqual(events[0][1].microwatts, 71.13)
        self.assertEqual(events[-1][1].dbm, -19.6)
        self.assertEqual(events[-1][1].microwatts, 10.78)


if __name__ == "__main__":
    unittest.main()
