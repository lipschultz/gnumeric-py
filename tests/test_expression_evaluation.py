"""
Gnumeric-py: Reading and writing gnumeric files with python
Copyright (C) 2017 Michael Lipschultz

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program. If not, see <http://www.gnu.org/licenses/>.
"""

import pytest

from gnumeric import Workbook
from gnumeric.evaluation_errors import EvaluationError
from gnumeric.expression_evaluation import evaluate, get_referenced_cells


@pytest.fixture()
def workbook_and_worksheet():
    workbook = Workbook()
    worksheet = workbook.create_sheet('Title')
    yield {'workbook': workbook, 'worksheet': worksheet}


@pytest.fixture()
def empty_worksheet():
    workbook = Workbook()
    worksheet = workbook.create_sheet('Title')
    yield worksheet


class TestOperatorAndConstant:
    ANY_CELL = None

    def test_it_can_start_with_a_plus_sign(self):
        actual = evaluate('+-2*3', self.ANY_CELL)
        assert -6 == actual

    def test_it_evaluates_integers(self):
        actual = evaluate('=54', self.ANY_CELL)
        assert 54 == actual
        assert isinstance(actual, int)

    def test_it_evaluates_floats(self):
        actual = evaluate('=5.4', self.ANY_CELL)
        assert 5.4 == actual

    def test_it_evaluates_text(self):
        actual = evaluate('="test"', self.ANY_CELL)
        assert 'test' == actual

    def test_it_evaluates_true(self):
        actual = evaluate('=TRUE', self.ANY_CELL)
        assert actual is True

    def test_it_evaluates_false(self):
        actual = evaluate('=FALSE', self.ANY_CELL)
        assert actual is False

    def test_it_evaluates_ref_error(self):
        actual = evaluate('=#REF!', self.ANY_CELL)
        assert EvaluationError.REF == actual

    @pytest.mark.parametrize(
        'formula, expected',
        [
            ('=-2', -2),
            ('=2+3', 5),
            ('=2-3', -1),
            ('=2*3', 6),
            ('=2/3', 2/3),
            ('=2^3', 2**3),
            ('=-2+3*4-10/5', 8),
            ('=2*(8-3)^2^3', 2*(8-3)**2**3),
            ('=TRUE+4', 5),
            ('=FALSE+4', 4),
            ('=1/0', EvaluationError.DIV0),
            ('=#REF!+1', EvaluationError.REF),
            ('=4+#REF!*3', EvaluationError.REF),
        ]
    )
    def test_basic_arithmetic_evaluation(self, formula, expected):
        actual = evaluate(formula, self.ANY_CELL)
        assert expected == actual, f'Result mismatch on {case}'

    @pytest.mark.parametrize(
        'formula',
        [
            '=4+"string"',
            '="string"+4',
            '=4-"string"',
            '="string"-4',
            '=4*"string"',
            '="string"*4',
            '=4/"string"',
            '="string"/4',
            '=4^"string"',
            '="string"^4',
        ]
    )
    def test_arithmetic_operations_between_numbers_and_strings_results_in_value_error(self, formula):
        actual = evaluate(formula, self.ANY_CELL)
        assert EvaluationError.VALUE == actual, f'Result mismatch on {case}'

    @pytest.mark.parametrize(
        'formula',
        [
            '=10000000000^10000000000',
            '=10000000000^1000',
            '=10^10000000000',
        ]
    )
    def test_arithmetic_operations_on_large_numbers_results_in_num_error(self, formula):
        actual = evaluate(formula, self.ANY_CELL)
        assert EvaluationError.NUM == actual, f'Result mismatch on {case}'

    @pytest.mark.parametrize(
        'formula, expected',
        [
            ('=2<3', True),
            ('=2<=3', True),
            ('=2>3', False),
            ('=2>=3', False),
            ('=2=3', False),
            ('=2<>3', True),
            ('=2<1+5', True),
            ('=2<#REF!', EvaluationError.REF),
        ]
    )
    def test_numeric_logical_evaluation(self, formula, expected):
        actual = evaluate(formula, self.ANY_CELL)
        assert expected == actual, f'Result mismatch on {case}'

    @pytest.mark.parametrize(
        'formula, expected',
        [
            ('="case"<"test"', True),
            ('="case"<="test"', True),
            ('="case">"test"', False),
            ('="case">="test"', False),
            ('="case"="test"', False),
            ('="case"<>"test"', True),
        ]
    )
    def test_string_logical_evaluation(self, formula, expected):
        actual = evaluate(formula, self.ANY_CELL)
        assert actual is expected, f'Result mismatch on {case}'

    @pytest.mark.parametrize(
        'formula, expected',
        [
            ('="case"<"CASE"', False),
            ('="case"<="CASE"', True),
            ('="case">"CASE"', False),
            ('="case">="CASE"', True),
            ('="case"="CASE"', True),
            ('="case"<>"CASE"', False),
        ]
    )
    def test_string_logical_evaluation_is_case_insensitive(self, formula, expected):
        actual = evaluate(formula, self.ANY_CELL)
        assert actual is expected, f'Result mismatch on {case}'

    @pytest.mark.parametrize(
        'formula, expected',
        [
            ('=2="2"', False),
            ('=2<>"2"', True),
            ('=50<"2"', True),
            ('=50<="2"', True),
            ('=2>"1"', False),
            ('=2>="1"', False),

            (f'=2=TRUE', False),
            (f'=2<>TRUE', True),
            (f'=50<TRUE', True),
            (f'=50<=TRUE', True),
            (f'=2>TRUE', False),
            (f'=2>=TRUE', False),

            (f'=2=FALSE', False),
            (f'=2<>FALSE', True),
            (f'=50<FALSE', True),
            (f'=50<=FALSE', True),
            (f'=2>FALSE', False),
            (f'=2>=FALSE', False),
        ]
    )
    def test_numbers_always_less_than_text_and_bools(self, formula, expected):
        actual = evaluate(formula, self.ANY_CELL)
        assert expected == actual, f'Result mismatch on {case}'

    @pytest.mark.parametrize(
        'formula, expected',
        [
            (f'="cat"=TRUE', False),
            (f'="cat"<>TRUE', True),
            (f'="cat"<TRUE', True),
            (f'="cat"<=TRUE', True),
            (f'="cat">TRUE', False),
            (f'="cat">=TRUE', False),

            (f'="cat"=FALSE', False),
            (f'="cat"<>FALSE', True),
            (f'="cat"<FALSE', True),
            (f'="cat"<=FALSE', True),
            (f'="cat">FALSE', False),
            (f'="cat">=FALSE', False),

            (f'=""=FALSE', False),
            (f'=""<>FALSE', True),
            (f'=""<FALSE', True),
            (f'=""<=FALSE', True),
            (f'="">FALSE', False),
            (f'="">=FALSE', False),
        ]
    )
    def test_text_always_less_than_bools(self, formula, expected):
        actual = evaluate(formula, self.ANY_CELL)
        assert actual is expected, f'Result mismatch on {case}'

    @pytest.mark.parametrize(
        'formula, expected',
        [
            (f'=TRUE=TRUE', True),
            (f'=TRUE<>TRUE', False),
            (f'=TRUE<TRUE', False),
            (f'=TRUE<=TRUE', True),
            (f'=TRUE>TRUE', False),
            (f'=TRUE>=TRUE', True),

            (f'=TRUE=FALSE', False),
            (f'=TRUE<>FALSE', True),
            (f'=TRUE<FALSE', False),
            (f'=TRUE<=FALSE', False),
            (f'=TRUE>FALSE', True),
            (f'=TRUE>=FALSE', True),

            (f'=FALSE=FALSE', True),
            (f'=FALSE<>FALSE', False),
            (f'=FALSE<FALSE', False),
            (f'=FALSE<=FALSE', True),
            (f'=FALSE>FALSE', False),
            (f'=FALSE>=FALSE', True),
        ]
    )
    def test_boolean_logical_evaluation(self, formula, expected):
        actual = evaluate(formula, self.ANY_CELL)
        assert actual is expected, f'Result mismatch on {case}'

    @pytest.mark.parametrize(
        'formula, expected',
        [
            ('="cat"&"dog"', 'catdog'),
            ('=2&"cat"', '2cat'),
            ('="cat"&2', 'cat2'),
            ('="cat"&-2+3*4-10/5', 'cat8'),
            ('=(2<3)&"cat"', 'TRUEcat'),
            ('=(2>3)&"cat"', 'FALSEcat'),
            ('=(2>3)&"cat"', 'FALSEcat'),
            ('=#REF!&"cat"', EvaluationError.REF),
        ]
    )
    def test_text_concatenation(self, formula, expected):
        actual = evaluate(formula, self.ANY_CELL)
        assert expected == actual, f'Result mismatch on {case}'


class TestCellReference:
    def test_referencing_string_cell_gets_string_value(self, empty_worksheet):
        expected_value = 'string'
        reference_cell = empty_worksheet.cell(0, 0)
        reference_cell.set_value(expected_value)

        test_cell = empty_worksheet.cell(0, 1)
        test_cell.set_value('=A1')

        actual_value = evaluate(test_cell.text, test_cell)
        assert expected_value == actual_value

    def test_referencing_string_cell_with_absolutes_gets_correct_value(self, empty_worksheet):
        expected_value = 'string'
        reference_cell = empty_worksheet.cell(0, 0)
        reference_cell.set_value(expected_value)

        test_cell = empty_worksheet.cell(0, 1)
        test_cell.set_value('=$A$1')

        actual_value = evaluate(test_cell.text, test_cell)
        assert expected_value == actual_value

    def test_referencing_non_existent_cell_returns_zero(self, empty_worksheet):
        test_cell = empty_worksheet.cell(0, 1)
        test_cell.set_value('=A1')

        actual_value = evaluate(test_cell.text, test_cell)
        assert 0 == actual_value

    def test_referencing_cell_in_another_sheet(self, workbook_and_worksheet):
        other_sheet = workbook_and_worksheet['workbook'].create_sheet('Other')
        expected_value = 'string'
        reference_cell = other_sheet.cell(0, 0)
        reference_cell.set_value(expected_value)

        test_cell = workbook_and_worksheet['worksheet'].cell(0, 1)
        test_cell.set_value('=Other!A1')

        actual_value = evaluate(test_cell.text, test_cell)
        assert expected_value == actual_value

    def test_referencing_cell_in_another_sheet_whose_name_contains_a_space(self, workbook_and_worksheet):
        other_sheet = workbook_and_worksheet['workbook'].create_sheet('Other Sheet')
        expected_value = 'string'
        reference_cell = other_sheet.cell(0, 0)
        reference_cell.set_value(expected_value)

        test_cell = workbook_and_worksheet['worksheet'].cell(0, 1)
        test_cell.set_value("='Other Sheet'!A1")

        actual_value = evaluate(test_cell.text, test_cell)
        assert expected_value == actual_value

    def test_referencing_cell_in_another_sheet_whose_name_contains_a_space_and_column_and_row_are_absolute_references(self, workbook_and_worksheet):
        other_sheet = workbook_and_worksheet['workbook'].create_sheet('Other Sheet')
        expected_value = 'string'
        reference_cell = other_sheet.cell(0, 0)
        reference_cell.set_value(expected_value)

        test_cell = workbook_and_worksheet['worksheet'].cell(0, 1)
        test_cell.set_value("='Other Sheet'!$A$1")

        actual_value = evaluate(test_cell.text, test_cell)
        assert expected_value == actual_value

    def test_referencing_cell_in_nonexistent_sheet_results_in_ref_error(self, empty_worksheet):
        test_cell = empty_worksheet.cell(0, 1)
        test_cell.set_value('=Other!A1')

        actual_value = evaluate(test_cell.text, test_cell)
        assert EvaluationError.REF == actual_value

    def test_referencing_cell_range_results_in_value_error(self, empty_worksheet):
        test_cell = empty_worksheet.cell(0, 1)
        test_cell.set_value('=A1:A5')

        actual_value = evaluate(test_cell.text, test_cell)
        assert EvaluationError.VALUE == actual_value

    def test_referencing_cell_range_in_function(self, empty_worksheet):
        expected_value = 0
        for row in range(5):
            empty_worksheet.cell(row, 0).set_value(row)
            expected_value += row

        test_cell = empty_worksheet.cell(0, 1)
        test_cell.set_value('=SUM(A1:A5)')

        actual_value = evaluate(test_cell.text, test_cell)
        assert expected_value == actual_value

    def test_referencing_cell_range_with_sheetname_in_function(self, empty_worksheet):
        expected_value = 0
        for row in range(5):
            empty_worksheet.cell(row, 0).set_value(row)
            expected_value += row

        test_cell = empty_worksheet.cell(0, 1)
        test_cell.set_value('=SUM(Title!A1:A5)')

        actual_value = evaluate(test_cell.text, test_cell)
        assert expected_value == actual_value

    def test_referencing_formula_cell_gets_formulas_result(self, empty_worksheet):
        reference_cell = empty_worksheet.cell(0, 0)
        reference_cell.set_value('=5+2')

        test_cell = empty_worksheet.cell(0, 1)
        test_cell.set_value('=A1')

        actual_value = evaluate(test_cell.text, test_cell)
        assert 7 == actual_value

    def test_function_argument_references_another_cell_uses_the_cells_value(self, empty_worksheet):
        reference_cell = empty_worksheet.cell(0, 0)
        reference_cell.set_value('=2-5')

        test_cell = empty_worksheet.cell(0, 1)
        test_cell.set_value('=ABS(A1)')

        actual_value = evaluate(test_cell.text, test_cell)
        assert 3 == actual_value

    def test_circular_references_end_and_use_zero_as_the_value(self, empty_worksheet):
        first_cell = empty_worksheet.cell(0, 0)
        first_cell.set_value('=B1')

        second_cell = empty_worksheet.cell(0, 1)
        second_cell.set_value('=A1')

        actual_first_cell = evaluate(first_cell.text, first_cell)
        actual_second_cell = evaluate(second_cell.text, second_cell)
        assert 0 == actual_first_cell
        assert 0 == actual_second_cell

    def test_creating_a_circular_reference_from_existing_cell_uses_cells_old_value_when_determining_cells_new_value(self, empty_worksheet):
        first_cell = empty_worksheet.cell(0, 0)
        first_cell.set_value(5)

        second_cell = empty_worksheet.cell(0, 1)
        second_cell.set_value('=A1')

        first_cell.set_value('=B1+5')

        actual_second_cell = evaluate(second_cell.text, second_cell)
        assert 5 == actual_second_cell

        actual_second_cell = evaluate(first_cell.text, first_cell)
        assert 10 == actual_second_cell

    @pytest.mark.skip('Not implemented yet')
    def test_updating_referenced_cell_updates_expression_cell_value(self):
        pass


class TestGetCellReference:
    def test_gets_cell_when_only_one_cell_referenced(self, empty_worksheet):
        for row in range(5):
            cell = empty_worksheet.cell(row, 0)
            cell.set_value(row)
        expected_cells = {empty_worksheet.cell(0, 0)}

        test_cell = empty_worksheet.cell(0, 1)
        test_cell.set_value('=5*A1')

        actual_cells = get_referenced_cells(test_cell.text, test_cell)
        assert expected_cells == actual_cells

    def test_get_all_cells_in_range_when_referencing_range_of_cell(self, empty_worksheet):
        for row in range(5):
            cell = empty_worksheet.cell(row, 0)
            cell.set_value(row)
        expected_cells = empty_worksheet.get_cell_collection('A1', 'A5')

        test_cell = empty_worksheet.cell(0, 1)
        test_cell.set_value('=SUM(A1:A5)')

        actual_cells = get_referenced_cells(test_cell.text, test_cell)
        assert set(expected_cells) == actual_cells

    def test_get_all_cells_in_range_when_referencing_multiple_cell_ranges(self, empty_worksheet):
        for row in range(5):
            cell = empty_worksheet.cell(row, 0)
            cell.set_value(row)
        expected_cells = empty_worksheet.get_cell_collection('A1', 'A3') + empty_worksheet.get_cell_collection('A5', 'A5')

        test_cell = empty_worksheet.cell(0, 1)
        test_cell.set_value('=SUM(A1:A3,A5)')

        actual_cells = get_referenced_cells(test_cell.text, test_cell)
        assert set(expected_cells) == actual_cells

    def test_get_all_cells_when_referencing_cell_range_inside_function_and_referencing_cells_outside(self, empty_worksheet):
        for row in range(5):
            cell = empty_worksheet.cell(row, 0)
            cell.set_value(row)
        test_cell = empty_worksheet.cell(0, 1)
        test_cell.set_value('=SUM(A1:A3,A5)+A10')

        actual_cells = get_referenced_cells(test_cell.text, test_cell)

        expected_cells = empty_worksheet.get_cell_collection('A1', 'A3') + empty_worksheet.get_cell_collection('A5', 'A5') + empty_worksheet.get_cell_collection('A10', 'A10', create_cells=True)
        assert set(expected_cells) == actual_cells


class TestFunctionEvaluation:
    ANY_CELL = None

    @pytest.mark.parametrize(
        'formula',
        ['=NAMEDOESNOTEXIST()', '=ABS']
    )
    def test_name_error(self, formula):
        assert EvaluationError.NAME == evaluate(formula, self.ANY_CELL)

    @pytest.mark.parametrize(
        'formula, expected',
        [
            ('=ABS(3)', 3),
            ('=ABS(-3)', 3),
            ('=ABS(5/3)', 5 / 3),
            ('=ABS(-5/3)', 5 / 3),
            ('=ABS(TRUE)', 1),
            ('=ABS(FALSE)', 0),
            ('=ABS("string")', EvaluationError.VALUE),
            ('=ABS(#REF!)', EvaluationError.REF),
            ('=ABS()', EvaluationError.NA),
        ]
    )
    def test_abs(self, formula, expected):
        assert expected == evaluate(formula, self.ANY_CELL)

    @pytest.mark.parametrize(
        'formula, expected',
        [
            ('=LEN("string")', 6),
            ('=LEN(TRUE)', 4),
            ('=LEN(FALSE)', 5),
            ('=LEN(12)', 2),
            ('=LEN(12/5)', 3),
            ('=LEN(5/3)', 18),
            ('=LEN(#REF!)', EvaluationError.REF),
        ]
    )
    def test_len(self, formula, expected):
        assert expected == evaluate(formula, self.ANY_CELL)
