from rustpy_xlsxwriter import validate_sheet_name


class TestValidateSheetName:
    def test_valid_name(self):
        assert validate_sheet_name("Sheet1") is True

    def test_invalid_characters(self):
        for char in ["[", "]", ":", "*", "?", "/", "\\"]:
            assert validate_sheet_name(f"Test{char}") is False

    def test_empty_name(self):
        assert validate_sheet_name("") is False

    def test_max_length_31(self):
        assert validate_sheet_name("A" * 31) is True

    def test_exceeds_max_length(self):
        assert validate_sheet_name("A" * 32) is False

    def test_unicode_name(self):
        assert validate_sheet_name("シート1") is True

    def test_unicode_exceeds_length(self):
        assert validate_sheet_name("あ" * 32) is False

    def test_spaces_allowed(self):
        assert validate_sheet_name("My Sheet") is True
