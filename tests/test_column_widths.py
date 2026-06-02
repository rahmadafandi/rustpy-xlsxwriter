import openpyxl
import pytest

from rustpy_xlsxwriter import write_worksheet, write_worksheets, FastExcel

RECORDS = [
    {"row": 1, "var": "age", "label": "Age in years"},
    {"row": 2, "var": "name", "label": "Full legal name"},
]


def widths(path, sheet=None):
    """{column_letter: width} for a sheet (active if sheet=None), expanding
    range-based <col min max> entries that set_column_range_width emits."""
    wb = openpyxl.load_workbook(str(path))
    ws = wb[sheet] if sheet else wb.active
    out = {}
    for dim in ws.column_dimensions.values():
        if not dim.width:
            continue
        for ci in range(dim.min or 1, (dim.max or dim.min or 1) + 1):
            out[openpyxl.utils.get_column_letter(ci)] = dim.width
    wb.close()
    return out


def approx_width(expected):
    """
    openpyxl reads back Excel column widths with a +0.7109375 rounding offset
    introduced by rust_xlsxwriter's pixel-to-character-unit conversion.  We
    accept anything within ±1 character unit of the requested value so the
    tests remain stable across Excel rendering engines.
    """
    return pytest.approx(expected, abs=0.75)


class TestSingleSheetWidthsFunctional:
    def test_uniform_width(self, tmp_path):
        p = tmp_path / "uniform.xlsx"
        write_worksheet(RECORDS, p, column_width=15, autofit=False)
        w = widths(p)
        assert w["A"] == approx_width(15)
        assert w["B"] == approx_width(15)
        assert w["C"] == approx_width(15)

    def test_dict_by_name(self, tmp_path):
        p = tmp_path / "dict.xlsx"
        write_worksheet(RECORDS, p, column_widths={"row": 7, "var": 22, "label": 40}, autofit=False)
        w = widths(p)
        assert w["A"] == approx_width(7)
        assert w["B"] == approx_width(22)
        assert w["C"] == approx_width(40)

    def test_list_positional(self, tmp_path):
        p = tmp_path / "list.xlsx"
        write_worksheet(RECORDS, p, column_widths=[7, 22, 40], autofit=False)
        w = widths(p)
        assert w["A"] == approx_width(7)
        assert w["B"] == approx_width(22)
        assert w["C"] == approx_width(40)

    def test_uniform_plus_override(self, tmp_path):
        p = tmp_path / "both.xlsx"
        write_worksheet(RECORDS, p, column_width=10, column_widths={"label": 40}, autofit=False)
        w = widths(p)
        assert w["A"] == approx_width(10)
        assert w["B"] == approx_width(10)
        assert w["C"] == approx_width(40)
        # Verify override actually changed C differently from A/B
        assert w["C"] > w["A"] + 20

    def test_explicit_beats_autofit(self, tmp_path):
        p = tmp_path / "vs_autofit.xlsx"
        write_worksheet(RECORDS, p, column_widths={"label": 50}, autofit=True)
        assert widths(p)["C"] == approx_width(50)


class TestMultiSheetWidths:
    def test_functional_keyed(self, tmp_path):
        p = tmp_path / "multi.xlsx"
        write_worksheets(
            [("Raw", RECORDS), ("Meta", RECORDS)],
            p,
            autofit=False,
            column_width={"general": 12, "Meta": 8},
            column_widths={"Meta": {"label": 40}},
        )
        raw = widths(p, "Raw")
        meta = widths(p, "Meta")
        assert raw["A"] == approx_width(12)
        assert raw["C"] == approx_width(12)
        assert meta["A"] == approx_width(8)
        assert meta["B"] == approx_width(8)
        assert meta["C"] == approx_width(40)


class TestSingleSheetWidthsBuilder:
    def test_uniform_width(self, tmp_path):
        p = tmp_path / "b_uniform.xlsx"
        FastExcel(p, autofit=False).sheet("S", RECORDS, column_width=15).save()
        w = widths(p)
        assert w["A"] == approx_width(15)
        assert w["C"] == approx_width(15)

    def test_dict_by_name(self, tmp_path):
        p = tmp_path / "b_dict.xlsx"
        FastExcel(p, autofit=False).sheet(
            "S", RECORDS, column_widths={"row": 7, "var": 22, "label": 40}
        ).save()
        w = widths(p)
        assert w["A"] == approx_width(7)
        assert w["B"] == approx_width(22)
        assert w["C"] == approx_width(40)

    def test_list_positional(self, tmp_path):
        p = tmp_path / "b_list.xlsx"
        FastExcel(p, autofit=False).sheet("S", RECORDS, column_widths=[7, 22, 40]).save()
        w = widths(p)
        assert w["A"] == approx_width(7)
        assert w["C"] == approx_width(40)

    def test_uniform_plus_override(self, tmp_path):
        p = tmp_path / "b_both.xlsx"
        FastExcel(p, autofit=False).sheet(
            "S", RECORDS, column_width=10, column_widths={"label": 40}
        ).save()
        w = widths(p)
        assert w["A"] == approx_width(10)
        assert w["C"] == approx_width(40)


class TestMultiSheetWidthsBuilder:
    def test_builder_per_sheet(self, tmp_path):
        p = tmp_path / "builder_multi.xlsx"
        (
            FastExcel(p, autofit=False)
            .sheet("Raw", RECORDS, column_width=12)
            .sheet("Meta", RECORDS, column_widths=[8, 8, 40])
            .save()
        )
        raw = widths(p, "Raw")
        meta = widths(p, "Meta")
        assert raw["A"] == approx_width(12)
        assert raw["C"] == approx_width(12)
        assert meta["A"] == approx_width(8)
        assert meta["B"] == approx_width(8)
        assert meta["C"] == approx_width(40)


class TestWidthWarningsAndIndex:
    def test_unknown_name_warns(self, tmp_path):
        p = tmp_path / "unknown.xlsx"
        with pytest.warns(UserWarning, match="unknown column 'nope'"):
            FastExcel(p, autofit=False).sheet(
                "S", RECORDS, column_widths={"row": 7, "nope": 9}
            ).save()
        assert widths(p)["A"] == approx_width(7)  # valid column still applied

    def test_list_too_long_warns(self, tmp_path):
        p = tmp_path / "toolong.xlsx"
        with pytest.warns(UserWarning, match=r"index 3 out of range"):
            FastExcel(p, autofit=False).sheet(
                "S", RECORDS, column_widths=[7, 22, 40, 99]
            ).save()
        assert widths(p)["C"] == approx_width(40)

    def test_invalid_width_warns(self, tmp_path):
        p = tmp_path / "invalid.xlsx"
        with pytest.warns(UserWarning, match=r"invalid width -5"):
            FastExcel(p, autofit=False).sheet(
                "S", RECORDS, column_widths={"row": -5}
            ).save()
        # Negative width skipped → column "A" has no explicit width applied.
        assert "A" not in widths(p)

    def test_index_columns_alignment(self, tmp_path):
        """index_columns only styles columns; it does not shift positions,
        so width indices still map to data column order."""
        p = tmp_path / "indexed.xlsx"
        write_worksheet(
            RECORDS, p, index_columns=["row"], column_widths={"label": 40},
            autofit=False,
        )
        assert widths(p)["C"] == approx_width(40)

    def test_unsupported_spec_type_raises(self, tmp_path):
        p = tmp_path / "badtype.xlsx"
        with pytest.raises(ValueError, match="must be a dict .* or a list"):
            FastExcel(p, autofit=False).sheet("S", RECORDS, column_widths=42).save()

    def test_sheet_absent_from_both_dicts(self, tmp_path):
        """A sheet with no sheet-specific and no 'general' width entry gets
        no explicit widths (autofit/default)."""
        p = tmp_path / "absent.xlsx"
        write_worksheets(
            [("Has", RECORDS), ("None", RECORDS)],
            p,
            autofit=False,
            column_width={"Has": 15},
        )
        has = widths(p, "Has")
        none = widths(p, "None")
        assert has["A"] == approx_width(15)
        assert none == {}  # no explicit width set for the unmatched sheet
