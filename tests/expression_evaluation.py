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

import unittest

from gnumeric.expression_evaluation import evaluate


class EvaluationTests(unittest.TestCase):
    ANY_CELL = None

    def test_it_can_start_with_a_plus_sign(self):
        actual = evaluate('+-2*3', self.ANY_CELL)
        self.assertEqual(-6, actual)

    def test_it_evaluates_integers(self):
        actual = evaluate('=54', self.ANY_CELL)
        self.assertEqual(54, actual)

    def test_it_evaluates_floats(self):
        actual = evaluate('=5.4', self.ANY_CELL)
        self.assertEqual(5.4, actual)

    def test_it_evaluates_text(self):
        actual = evaluate(f'="test"', self.ANY_CELL)
        self.assertEqual('test', actual)

    def test_basic_arithmetic_evaluation(self):
        cases = (
            ('=-2', -2),
            ('=2+3', 5),
            ('=2-3', -1),
            ('=2*3', 6),
            ('=2/3', 2/3),
            ('=2^3', 2**3),
            ('=-2+3*4-10/5', 8),
            ('=2*(8-3)^2^3', 2*(8-3)**2**3),
        )
        for case, expected_result in cases:
            actual = evaluate(case, self.ANY_CELL)
            self.assertEqual(expected_result, actual, f'Result mismatch on {case}')

    def test_numeric_logical_evaluation(self):
        cases = (
            ('=2<3', True),
            ('=2<=3', True),
            ('=2>3', False),
            ('=2>=3', False),
            ('=2=3', False),
            ('=2<>3', True),
            ('=2<1+5', True),
        )
        for case, expected_result in cases:
            actual = evaluate(case, self.ANY_CELL)
            self.assertEqual(expected_result, actual, f'Result mismatch on {case}')

    def test_string_logical_evaluation(self):
        cases = (
            ('="case"<"test"', True),
            ('="case"<="test"', True),
            ('="case">"test"', False),
            ('="case">="test"', False),
            ('="case"="test"', False),
            ('="case"<>"test"', True),
        )
        for case, expected_result in cases:
            actual = evaluate(case, self.ANY_CELL)
            self.assertEqual(expected_result, actual, f'Result mismatch on {case}')

    def test_string_logical_evaluation_is_case_insensitive(self):
        cases = (
            ('="case"<"CASE"', False),
            ('="case"<="CASE"', True),
            ('="case">"CASE"', False),
            ('="case">="CASE"', True),
            ('="case"="CASE"', True),
            ('="case"<>"CASE"', False),
        )
        for case, expected_result in cases:
            actual = evaluate(case, self.ANY_CELL)
            self.assertEqual(expected_result, actual, f'Result mismatch on {case}')

    def test_numbers_always_less_than_text_and_bools(self):
        true = '(2<3)'
        false = '(2>3)'
        cases = (
            ('=2="2"', False),
            ('=2<>"2"', True),
            ('=50<"2"', True),
            ('=50<="2"', True),
            ('=2>"1"', False),
            ('=2>="1"', False),

            (f'=2={true}', False),
            (f'=2<>{true}', True),
            (f'=50<{true}', True),
            (f'=50<={true}', True),
            (f'=2>{true}', False),
            (f'=2>={true}', False),

            (f'=2={false}', False),
            (f'=2<>{false}', True),
            (f'=50<{false}', True),
            (f'=50<={false}', True),
            (f'=2>{false}', False),
            (f'=2>={false}', False),
        )
        for case, expected_result in cases:
            actual = evaluate(case, self.ANY_CELL)
            self.assertEqual(expected_result, actual, f'Result mismatch on {case}')

    def test_text_always_less_than_bools(self):
        true = '(2<3)'
        false = '(2>3)'
        cases = (
            (f'="cat"={true}', False),
            (f'="cat"<>{true}', True),
            (f'="cat"<{true}', True),
            (f'="cat"<={true}', True),
            (f'="cat">{true}', False),
            (f'="cat">={true}', False),

            (f'="cat"={false}', False),
            (f'="cat"<>{false}', True),
            (f'="cat"<{false}', True),
            (f'="cat"<={false}', True),
            (f'="cat">{false}', False),
            (f'="cat">={false}', False),

            (f'=""={false}', False),
            (f'=""<>{false}', True),
            (f'=""<{false}', True),
            (f'=""<={false}', True),
            (f'="">{false}', False),
            (f'="">={false}', False),
        )
        for case, expected_result in cases:
            actual = evaluate(case, self.ANY_CELL)
            self.assertEqual(expected_result, actual, f'Result mismatch on {case}')

    def test_boolean_logical_evaluation(self):
        # TODO: How to handle boolean constants
        true = '(2<3)'
        false = '(2>3)'
        cases = (
            (f'={true}={true}', True),
            (f'={true}<>{true}', False),
            (f'={true}<{true}', False),
            (f'={true}<={true}', True),
            (f'={true}>{true}', False),
            (f'={true}>={true}', True),

            (f'={true}={false}', False),
            (f'={true}<>{false}', True),
            (f'={true}<{false}', False),
            (f'={true}<={false}', False),
            (f'={true}>{false}', True),
            (f'={true}>={false}', True),

            (f'={false}={false}', True),
            (f'={false}<>{false}', False),
            (f'={false}<{false}', False),
            (f'={false}<={false}', True),
            (f'={false}>{false}', False),
            (f'={false}>={false}', True),
        )
        for case, expected_result in cases:
            actual = evaluate(case, self.ANY_CELL)
            self.assertEqual(expected_result, actual, f'Result mismatch on {case}')

    def test_text_concatenation(self):
        cases = (
            ('="cat"&"dog"', 'catdog'),
            ('=2&"cat"', '2cat'),
            ('="cat"&2', 'cat2'),
            ('="cat"&-2+3*4-10/5', 'cat8'),
            ('=(2<3)&"cat"', 'TRUEcat'),
            ('=(2>3)&"cat"', 'FALSEcat'),
        )
        for case, expected_result in cases:
            actual = evaluate(case, self.ANY_CELL)
            self.assertEqual(expected_result, actual, f'Result mismatch on {case}')
