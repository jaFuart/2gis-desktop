import contextlib
import io
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
from xml.etree import ElementTree
from zipfile import ZipFile

import parse_2gis
from app import city_code, file_name
from cities import CITIES
from desktop_backend import History, publish_result, save_checkpoint


class ExportTests(unittest.TestCase):
    def test_xlsx_splits_at_the_row_limit_and_repeats_headers(self):
        columns = ("Телефон", "Рейтинг")
        cards = [(f"+7000000000{index}", f"4.{index}") for index in range(1, 6)]
        ns = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
        office = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        package = "http://schemas.openxmlformats.org/package/2006/relationships"
        content = "http://schemas.openxmlformats.org/package/2006/content-types"

        with tempfile.TemporaryDirectory() as folder, patch.object(parse_2gis, "EXCEL_ROWS_PER_SHEET", 3):
            for count, expected_lengths in ((0, [1]), (2, [3]), (4, [3, 3]), (5, [3, 3, 2])):
                with self.subTest(count=count):
                    output = Path(folder) / f"{count}.xlsx"
                    parse_2gis.write_xlsx(output, columns, cards[:count])
                    with ZipFile(output) as archive:
                        workbook = ElementTree.fromstring(archive.read("xl/workbook.xml"))
                        relationships = ElementTree.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
                        types = ElementTree.fromstring(archive.read("[Content_Types].xml"))
                        targets = {item.get("Id"): item.get("Target")
                                   for item in relationships.findall(f"{{{package}}}Relationship")}
                        overrides = {item.get("PartName") for item in types.findall(f"{{{content}}}Override")}
                        sheets = workbook.findall("x:sheets/x:sheet", ns)
                        self.assertEqual(len(sheets), len(expected_lengths))
                        self.assertIn("styles.xml", targets.values())
                        exported = []
                        for number, (entry, length) in enumerate(zip(sheets, expected_lengths), 1):
                            self.assertEqual(entry.get("name"), "Карточки" if number == 1 else f"Карточки {number}")
                            path = f"xl/{targets[entry.get(f'{{{office}}}id')]}"
                            self.assertIn(f"/{path}", overrides)
                            sheet = ElementTree.fromstring(archive.read(path))
                            rows = sheet.findall("x:sheetData/x:row", ns)
                            self.assertEqual([row.get("r") for row in rows], [str(i) for i in range(1, length + 1)])
                            strings = []
                            for row in rows:
                                cells = row.findall("x:c", ns)
                                self.assertTrue(all(cell.get("t") == "inlineStr" and cell.get("s") == "1"
                                                    for cell in cells))
                                strings.append(tuple(cell.findtext("x:is/x:t", namespaces=ns) for cell in cells))
                            self.assertEqual(strings[0], columns)
                            exported.extend(strings[1:])
                        self.assertEqual(exported, cards[:count])

    def test_cancel_after_first_card_keeps_text_export(self):
        progress = []
        row = {"Название места": "Свои", "Рейтинг": "4.8", "Номера телефонов": "+79069392192"}

        def browse(args, _):
            yield row
            parse_2gis.cancel_checkpoint(args)
            self.fail("Expected the next card to be cancelled")

        with tempfile.TemporaryDirectory() as folder:
            output = Path(folder) / "partial.xlsx"
            with patch.object(parse_2gis, "collect_browser", browse), contextlib.redirect_stdout(io.StringIO()):
                rc = parse_2gis.main(["--output", str(output)],
                                    on_progress=lambda event, data: progress.append((event, data)),
                                    is_cancelled=lambda: bool(progress))
            self.assertEqual(rc, 130)
            self.assertEqual([event for event, _ in progress], ["row", "saving"])
            with ZipFile(output) as archive:
                xml = ElementTree.fromstring(archive.read("xl/worksheets/sheet1.xml"))
            ns = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
            cells = {cell.get("r"): cell for cell in xml.findall(".//x:c", ns)}
            self.assertEqual(cells["K2"].get("t"), "inlineStr")
            self.assertEqual(cells["V2"].findtext("x:is/x:t", namespaces=ns), "+79069392192")

    def test_cancel_before_results_creates_no_file(self):
        def browse(args, _):
            parse_2gis.cancel_checkpoint(args)
            yield {}

        with tempfile.TemporaryDirectory() as folder, patch.object(parse_2gis, "collect_browser", browse):
            output = Path(folder) / "cancelled.xlsx"
            self.assertEqual(parse_2gis.main(["--output", str(output)], is_cancelled=lambda: True), 130)
            self.assertFalse(output.exists())

    def test_the_contact_group_keeps_only_number_labels(self):
        links = [("tel:+79069392192", "+7‒906‒939..."), ("tel:+79030744333", "Позвонить")]
        self.assertEqual(parse_2gis.contact_fields(links)["Номера телефонов"], "+79069392192")

    def test_checkpoint_publishes_without_overwriting(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder)
            staged, target = path / "stage.csv", path / "result.csv"
            save_checkpoint(staged, [{"Название места": "Кафе"}])
            publish_result(staged, target)
            with self.assertRaises(FileExistsError):
                publish_result(staged, target)
            self.assertEqual(staged.read_bytes(), target.read_bytes())

    def test_history_reopens_with_the_result(self):
        with tempfile.TemporaryDirectory() as folder:
            target = Path(folder) / "result.xlsx"
            path = Path(folder) / "history.sqlite3"
            history = History(path)
            key = history.start("Кафе", "gornoaltaysk", 8, target)
            history.finish(key, "completed", 8, target, "")
            history.db.close()
            reopened = History(path)
            self.addCleanup(reopened.db.close)
            row, = reopened.list()
            self.assertEqual((row["query"], row["status"], row["rows"], row["file"]),
                             ("Кафе", "completed", 8, str(target)))
            reopened.db.close()

    def test_query_city_accepts_real_map_url(self):
        self.assertEqual(city_code("горно-алтайск"), "gornoaltaysk")
        self.assertEqual(city_code("https://2gis.ru/gornoaltaysk/search/test"), "gornoaltaysk")
        with self.assertRaises(ValueError):
            city_code("https://example.com/moscow")
        self.assertTrue(file_name("Кафе / Бар").endswith(".xlsx"))

    def test_offline_city_catalogue_uses_map_aliases(self):
        self.assertGreaterEqual(len(CITIES), 100)
        self.assertEqual(city_code("Нижний Новгород"), "n_novgorod")
        self.assertEqual(city_code("Питер"), "spb")
        self.assertEqual(city_code("Орел"), CITIES["Орёл"])
        self.assertEqual(city_code("Армавир"), "armawir")


if __name__ == "__main__":
    unittest.main()
