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

    def test_logical_evaluation(self):
        cases = (
            ('=2<3', True),
            ('=2<=3', True),
            ('=2>3', False),
            ('=2>=3', False),
            ('=2=3', False),
            ('=2<>3', True),
        )
        for case, expected_result in cases:
            actual = evaluate(case, self.ANY_CELL)
            self.assertEqual(expected_result, actual, f'Result mismatch on {case}')
